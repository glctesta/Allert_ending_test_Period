# main.py
import logging
from datetime import datetime
from db_connection import DatabaseConnection
from config_manager import ConfigManager
from utils import get_email_recipients, send_email
import pyodbc

# Configurazione logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('employee_notifications.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("TraceabilityRS")


def get_employees_with_upcoming_test_end(conn):
    """Recupera i dipendenti con periodo di prova in scadenza"""
    query = """
            select h.employeehirehistoryid,
                   e.EmployeeSurname + ' ' + e.employeename + ' [CNP: ' + e.EmployeeNID + ']' as Employee,
                   format(hiredate, 'd', 'ro-ro')                                             AS HireDate,
                   cast(TestPeriod as nvarchar(2)) + ' days'                                  as TestPeriod,
                   format(dateadd(DAY, [testPeriod], HireDate), 'd', 'ro-ro')                 as LastTestDate,
                   abs(datediff(DAY, dateadd(DAY, [testPeriod], HireDate), getdate()))        as MissingDayAtEndTestDate
            from employee.dbo.employees e
                     inner join employee.dbo.employeehirehistory h on e.employeeid = h.EmployeeId and h.EmployeerId = 2
            left join [Employee].[dbo].[EmployeeEvaluationHistory] ev on ev.EmployeeHireHistoryId=h.EmployeeHireHistoryId
where datediff(DAY,dateadd(DAY, [testPeriod],HireDate),getdate()) between -30 and 0
and ev.IdValutazioneStorico is null;\
            """

    try:
        with conn.cursor() as cursor:
            cursor.execute(query)
            results = cursor.fetchall()

        employees = []
        for row in results:
            employee = {
                'employeehirehistoryid': row[0],
                'Employee': row[1],
                'HireDate': row[2],
                'TestPeriod': row[3],
                'LastTestDate': row[4],
                'MissingDayAtEndTestDate': row[5]
            }
            employees.append(employee)

        logger.info(f"Trovati {len(employees)} dipendenti con periodo di prova in scadenza")
        return employees

    except Exception as e:
        logger.error(f"Errore nell'esecuzione della query: {str(e)}")
        raise


def get_manager_emails(conn, employee_ids):
    """Recupera gli indirizzi email dei manager tramite stored procedure"""
    if not employee_ids:
        return []

    try:
        # Creiamo una stringa con gli ID separati da virgola
        id_list = ",".join([str(id) for id in employee_ids])

        # Query alternativa che non richiede il tipo di tabella
        # Assumendo che la SP accetti una stringa di ID separati da virgola
        sp_query = f"""
        EXEC GetManager @EmployeeIds = ?
        """

        with conn.cursor() as cursor:
            cursor.execute(sp_query, id_list)
            results = cursor.fetchall()

        emails = [row[0] for row in results if row[0] and '@' in row[0]]
        logger.info(f"Trovati {len(emails)} indirizzi email manager")
        return emails

    except Exception as e:
        logger.error(f"Errore nel recupero email manager: {str(e)}")

        # Fallback: prova a recuperare email dai settings
        try:
            logger.info("Tentativo fallback: recupero email dai settings")
            return get_email_recipients(conn, 'Sys_Email_Purchase')
        except Exception as fallback_error:
            logger.error(f"Anche il fallback ha fallito: {str(fallback_error)}")
            return []


def create_email_body(employees):
    """Crea il corpo HTML dell'email con il logo"""
    html_body = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            .header {{ background-color: #f8f9fa; padding: 15px; border-left: 4px solid #007cba; }}
            .logo {{ max-width: 200px; margin-bottom: 15px; }}
            .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            .table th, .table td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
            .table th {{ background-color: #f2f2f2; }}
            .urgent {{ color: #d9534f; font-weight: bold; }}
            .warning {{ color: #f0ad4e; }}
            .info {{ color: #5bc0de; }}
        </style>
    </head>
    <body>
        <div class="header">
            <!-- Logo incorporato -->
            <img src="cid:Logo.png" alt="Company Logo" class="logo">
            <h2>Notifica Scadenza Periodo di Prova</h2>
            <p>Data generazione report: {datetime.now().strftime('%d.%m.%Y %H:%M')}</p>
        </div>

        <p>Si informa che i seguenti dipendenti hanno il periodo di prova in scadenza:</p>

        <table class="table">
            <thead>
                <tr>
                    <th>Dipendente</th>
                    <th>Data Assunzione</th>
                    <th>Periodo di Prova</th>
                    <th>Data Fine Prova</th>
                    <th>Giorni Mancanti</th>
                </tr>
            </thead>
            <tbody>
    """

    for emp in employees:
        days_left = emp['MissingDayAtEndTestDate']
        if days_left <= 3:
            days_class = "urgent"
        elif days_left <= 7:
            days_class = "warning"
        else:
            days_class = "info"

        html_body += f"""
                <tr>
                    <td>{emp['Employee']}</td>
                    <td>{emp['HireDate']}</td>
                    <td>{emp['TestPeriod']}</td>
                    <td>{emp['LastTestDate']}</td>
                    <td class="{days_class}">{days_left} giorni</td>
                </tr>
        """

    html_body += """
            </tbody>
        </table>

        <p><strong>Note:</strong></p>
        <ul>
            <li class="urgent">● Rosso: scadenza entro 3 giorni</li>
            <li class="warning">● Arancione: scadenza entro 7 giorni</li>
            <li class="info">● Blu: scadenza oltre 7 giorni</li>
        </ul>

        <p>Si prega di prendere le necessarie azioni per la valutazione del periodo di prova.</p>

        <hr>
        <p style="color: #666; font-size: 12px;">
            Questo è un messaggio automatico, si prega di non rispondere.
        </p>
    </body>
    </html>
    """

    return html_body


def main():
    """Funzione principale"""
    try:
        # Inizializzazione configurazione e connessione
        config_manager = ConfigManager()
        db_connection = DatabaseConnection(config_manager)

        # Connessione al database
        conn = db_connection.connect()

        # Recupera dipendenti con periodo di prova in scadenza
        employees = get_employees_with_upcoming_test_end(conn)

        if not employees:
            logger.info("Nessun dipendente con periodo di prova in scadenza trovato")
            return

        # Recupera gli ID per la stored procedure
        employee_ids = [emp['employeehirehistoryid'] for emp in employees]

        # Recupera email dei manager
        manager_emails = get_manager_emails(conn, employee_ids)

        if not manager_emails:
            logger.warning("Nessun indirizzo email manager trovato")
            return

        # Crea e invia l'email
        subject = f"Notifica Scadenza Periodo di Prova - {datetime.now().strftime('%d.%m.%Y')}"
        body = create_email_body(employees)

        # Invia l'email come HTML
        send_email(
            recipients=manager_emails,
            subject=subject,
            body=body,
            is_html=True
        )

        logger.info(f"Email inviata con successo a {len(manager_emails)} destinatari")

    except Exception as e:
        logger.error(f"Errore durante l'esecuzione del programma: {str(e)}")
        raise
    finally:
        # Chiude la connessione
        if 'db_connection' in locals():
            db_connection.disconnect()


if __name__ == "__main__":
    main()