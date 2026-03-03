import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_finish_notification(dest_email, client_name, ot_num, items_count):
    """
    Envía un correo electrónico al cliente notificando que su parte ha sido finalizado.
    """
    smtp_server = os.environ.get("SMTP_SERVER")
    smtp_port = os.environ.get("SMTP_PORT", "587")
    smtp_user = os.environ.get("SMTP_USER")
    smtp_pass = os.environ.get("SMTP_PASS")
    smtp_from = os.environ.get("SMTP_FROM", smtp_user)

    if not all([smtp_server, smtp_user, smtp_pass]):
        print("Error: SMTP configuracion incompleta (SMTP_SERVER, SMTP_USER, SMTP_PASS)")
        return False, "Configuración de correo incompleta"

    try:
        msg = MIMEMultipart()
        msg['From'] = f"Xurgical SAT <{smtp_from}>"
        msg['To'] = dest_email
        msg['Subject'] = f"Finalización de revisión - Parte {ot_num}"

        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
            <div style="max-width: 600px; margin: 0 auto; border: 1px solid #eee; padding: 20px; border-radius: 10px;">
                <h2 style="color: #2c3e50;">¡Hola {client_name}!</h2>
                <p>Te informamos que hemos finalizado la revisión de los instrumentos del <b>Parte {ot_num}</b>.</p>
                <p>Se han revisado un total de <b>{items_count}</b> instrumentos.</p>
                <p>Ya puedes consultar todos los detalles, fotos y estados en nuestra aplicación web.</p>
                
                <div style="text-align: center; margin: 30px 0;">
                    <a href="https://xurgical-sat.onrender.com/" 
                       style="background-color: #f39c12; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold;">
                       Acceder a la App
                    </a>
                </div>

                <p style="font-size: 0.9rem; color: #666;">Si aún no tienes acceso o tienes dudas, por favor contacta con nosotros.</p>
                <hr style="border: 0; border-top: 1px solid #eee;">
                <p style="font-size: 0.8rem; color: #999;">Saludos,<br>El equipo de Xurgical SAT</p>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
        server.quit()
        return True, "Email enviado correctamente"
    except Exception as e:
        print(f"Error enviando email: {e}")
        return False, str(e)
