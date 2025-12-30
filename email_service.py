import os
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

logger = logging.getLogger(__name__)

# TransIP SMTP Configuration
SMTP_HOST = os.environ.get('SMTP_HOST', 'smtp.transip.email')
SMTP_PORT = int(os.environ.get('SMTP_PORT', 465))
SMTP_USERNAME = os.environ.get('SMTP_USERNAME', '')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD', '')
SMTP_FROM = os.environ.get('SMTP_FROM', 'info@theglobal-bedrijfsdiensten.nl')
SMTP_SECURE = os.environ.get('SMTP_SECURE', 'ssl')
FRONTEND_URL = os.environ.get('FRONTEND_URL', 'https://mandagenstaat-export.preview.emergentagent.com')

def send_email(to_email: str, subject: str, html_content: str, text_content: str = None):
    """
    Send email via TransIP SMTP
    """
    if not SMTP_USERNAME or not SMTP_PASSWORD:
        logger.warning(f"Email not sent - SMTP not configured. Would send to: {to_email}")
        logger.info(f"Subject: {subject}")
        return False
    
    try:
        # Create message
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = f"The Global Urenregistratie <{SMTP_FROM}>"
        msg['To'] = to_email
        
        # Add text and HTML parts
        if text_content:
            part1 = MIMEText(text_content, 'plain', 'utf-8')
            msg.attach(part1)
        
        part2 = MIMEText(html_content, 'html', 'utf-8')
        msg.attach(part2)
        
        # Send via SMTP
        if SMTP_SECURE == 'ssl':
            # SSL connection (port 465)
            server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT)
        else:
            # TLS connection (port 587)
            server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
            server.starttls()
        
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        server.send_message(msg)
        server.quit()
        
        logger.info(f"Email sent successfully to {to_email}")
        return True
        
    except Exception as e:
        logger.error(f"Error sending email: {str(e)}")
        return False


def send_invitation_email(to_email: str, token: str):
    """
    Send invitation email to new employee
    """
    register_link = f"{FRONTEND_URL}/register/{token}"
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
            .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
            .header {{ background: linear-gradient(135deg, #16a085 0%, #1abc9c 100%); color: white; padding: 30px; text-align: center; border-radius: 10px 10px 0 0; }}
            .content {{ background: #f9f9f9; padding: 30px; border-radius: 0 0 10px 10px; }}
            .button {{ display: inline-block; background: #16a085; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; margin: 20px 0; font-weight: bold; }}
            .footer {{ text-align: center; margin-top: 30px; color: #666; font-size: 12px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>Welkom bij The Global Urenregistratie</h1>
            </div>
            <div class="content">
                <p>Beste medewerker,</p>
                <p>Je bent uitgenodigd om een account aan te maken voor ons urenregistratie systeem.</p>
                <p>Klik op de onderstaande knop om je account te activeren:</p>
                <p style="text-align: center;">
                    <a href="{register_link}" class="button">Account Aanmaken</a>
                </p>
                <p>Of kopieer deze link in je browser:</p>
                <p style="word-break: break-all; background: white; padding: 10px; border-radius: 5px;">
                    {register_link}
                </p>
                <p>Deze uitnodiging is eenmalig te gebruiken.</p>
                <p>Met vriendelijke groet,<br>
                The Global Bedrijfsdiensten</p>
            </div>
            <div class="footer">
                <p>© 2025 The Global Bedrijfsdiensten. Alle rechten voorbehouden.</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    text_content = f"""
    Welkom bij The Global Urenregistratie
    
    Je bent uitgenodigd om een account aan te maken voor ons urenregistratie systeem.
    
    Gebruik deze link om je account te activeren:
    {register_link}
    
    Deze uitnodiging is eenmalig te gebruiken.
    
    Met vriendelijke groet,
    The Global Bedrijfsdiensten
    """
    
    return send_email(to_email, "Uitnodiging - The Global Urenregistratie", html_content, text_content)


def send_password_reset_email(to_email: str, token: str):
    """
    Send password reset email
    """
    reset_link = f"{FRONTEND_URL}/reset-password/{token}"
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
            .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
            .header {{ background: linear-gradient(135deg, #16a085 0%, #1abc9c 100%); color: white; padding: 30px; text-align: center; border-radius: 10px 10px 0 0; }}
            .content {{ background: #f9f9f9; padding: 30px; border-radius: 0 0 10px 10px; }}
            .button {{ display: inline-block; background: #16a085; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; margin: 20px 0; font-weight: bold; }}
            .footer {{ text-align: center; margin-top: 30px; color: #666; font-size: 12px; }}
            .warning {{ background: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 20px 0; border-radius: 5px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>Wachtwoord Resetten</h1>
            </div>
            <div class="content">
                <p>Beste gebruiker,</p>
                <p>We hebben een verzoek ontvangen om je wachtwoord te resetten.</p>
                <p>Klik op de onderstaande knop om een nieuw wachtwoord in te stellen:</p>
                <p style="text-align: center;">
                    <a href="{reset_link}" class="button">Wachtwoord Resetten</a>
                </p>
                <p>Of kopieer deze link in je browser:</p>
                <p style="word-break: break-all; background: white; padding: 10px; border-radius: 5px;">
                    {reset_link}
                </p>
                <div class="warning">
                    <strong>Let op:</strong> Deze link is 1 uur geldig. Als je dit verzoek niet hebt gedaan, negeer deze email dan.
                </div>
                <p>Met vriendelijke groet,<br>
                The Global Bedrijfsdiensten</p>
            </div>
            <div class="footer">
                <p>© 2025 The Global Bedrijfsdiensten. Alle rechten voorbehouden.</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    text_content = f"""
    Wachtwoord Resetten - The Global Urenregistratie
    
    We hebben een verzoek ontvangen om je wachtwoord te resetten.
    
    Gebruik deze link om een nieuw wachtwoord in te stellen:
    {reset_link}
    
    Deze link is 1 uur geldig. Als je dit verzoek niet hebt gedaan, negeer deze email dan.
    
    Met vriendelijke groet,
    The Global Bedrijfsdiensten
    """
    
    return send_email(to_email, "Wachtwoord Reset - The Global Urenregistratie", html_content, text_content)


def send_weekly_reminder_email(to_email: str, user_name: str):
    """
    Send weekly reminder to fill time entries
    """
    login_link = f"{FRONTEND_URL}/login"
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
            .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
            .header {{ background: linear-gradient(135deg, #16a085 0%, #1abc9c 100%); color: white; padding: 30px; text-align: center; border-radius: 10px 10px 0 0; }}
            .content {{ background: #f9f9f9; padding: 30px; border-radius: 0 0 10px 10px; }}
            .button {{ display: inline-block; background: #16a085; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; margin: 20px 0; font-weight: bold; }}
            .footer {{ text-align: center; margin-top: 30px; color: #666; font-size: 12px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>📝 Wekelijkse Herinnering</h1>
            </div>
            <div class="content">
                <p>Beste {user_name},</p>
                <p>Dit is een vriendelijke herinnering om je gewerkte uren van deze week in te vullen in het urenregistratie systeem.</p>
                <p style="text-align: center;">
                    <a href="{login_link}" class="button">Uren Invullen</a>
                </p>
                <p>Het invullen van je uren zorgt ervoor dat we een accuraat overzicht hebben van alle werkzaamheden.</p>
                <p>Met vriendelijke groet,<br>
                The Global Bedrijfsdiensten</p>
            </div>
            <div class="footer">
                <p>© 2025 The Global Bedrijfsdiensten. Alle rechten voorbehouden.</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    return send_email(to_email, "Herinnering: Uren invullen - The Global", html_content)
