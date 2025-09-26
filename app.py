import os
import smtplib
import time
import json
from flask import Flask, request, render_template, redirect, url_for, jsonify
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from werkzeug.utils import secure_filename
import logging
from dotenv import load_dotenv
from docx import Document
import pandas as pd
import threading

# Load environment variables
load_dotenv()

# Initialize Flask
app = Flask(__name__)

# Upload folder
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
CONFIG_FILE = 'email_config.json'

# Logging
logging.basicConfig(level=logging.INFO)

# Load/save config functions
def load_configurations():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return []

def save_configuration(new_config):
    configs = load_configurations()
    configs.append(new_config)
    with open(CONFIG_FILE, 'w') as f:
        json.dump(configs, f)

# Extract pages from Word (if needed)
def extract_pages_from_word(doc_path):
    document = Document(doc_path)
    pages = []
    page_content = []

    for para in document.paragraphs:
        if 'w:br' in para._element.xml and 'w:type="page"' in para._element.xml:
            if page_content:
                pages.append('\n'.join(page_content))
            page_content = []
        else:
            page_content.append(para.text)

    if page_content:
        pages.append('\n'.join(page_content))

    return pages

# Global state
paused = False
stop_sending = False
current_index = 0
total_emails = 0
sent_emails = 0
email_thread = None
state_lock = threading.Lock()

# Control routes
@app.route('/pause', methods=['POST'])
def pause_sending():
    global paused
    with state_lock:
        paused = True
    return jsonify({"status": "paused"})

@app.route('/resume', methods=['POST'])
def resume_sending():
    global paused
    with state_lock:
        paused = False
    logging.info("Resuming email sending...")
    return jsonify({"status": "resumed"})

@app.route('/stop', methods=['POST'])
def stop_sending_emails():
    global stop_sending
    with state_lock:
        stop_sending = True
    return jsonify({"status": "stopped"})

@app.route('/progress', methods=['GET'])
def get_progress():
    global total_emails, sent_emails
    with state_lock:
        return jsonify({
            "total_emails": total_emails,
            "sent_emails": sent_emails
        })

# Bulk email sender
def send_bulk_emails(excel_file, delay, min_limit, max_limit):
    global paused, stop_sending, current_index, total_emails, sent_emails
    email_configs = load_configurations()

    if not email_configs:
        logging.error("No email configurations found.")
        return

    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        logging.error(f"Failed to read Excel file: {e}")
        return

    # Adjust limits safely
    min_limit = max(1, min_limit)
    max_limit = min(len(df), max_limit)
    if min_limit > max_limit:
        logging.warning("min_limit > max_limit, adjusting values")
        min_limit, max_limit = max_limit, min_limit

    with state_lock:
        total_emails = max_limit - min_limit + 1
        sent_emails = 0

    config_index = 0
    count = 0

    for i in range(current_index, max_limit):
        with state_lock:
            if stop_sending:
                logging.info("Email sending stopped.")
                break
            while paused:
                logging.info("Email sending paused. Waiting...")
                time.sleep(1)

        row = df.iloc[i]
        config = email_configs[config_index]

        sender_email = config['sender_email']
        smtp_server = config['smtp_server']
        smtp_port = config['smtp_port']
        sender_name = config.get('sender_name', '')
        sender_password = config.get('sender_password', None)

        recipient_email = row.get('email', '').strip()
        first_name = str(row.get('first_name', '')).strip()
        last_name = str(row.get('last_name', '')).strip()
        company_name = str(row.get('company_name', '')).strip()
        subject = str(row.get('subject', '')).strip()
        body = str(row.get('body', '')).strip()

        if not recipient_email:
            continue

        # Personalization
        subject = subject.replace("{first_name}", first_name)\
                         .replace("{last_name}", last_name)\
                         .replace("{company_name}", company_name)
        body = body.replace("{first_name}", first_name)\
                   .replace("{last_name}", last_name)\
                   .replace("{company_name}", company_name)\
                   .replace("{sender_name}", sender_name)

        try:
            html_body = body.replace("\n•", "<br>&bull;").replace("\n", "<br>")
            html_content = f"<html><body>{html_body}</body></html>"

            msg = MIMEMultipart()
            msg['From'] = f"{sender_name} <{sender_email}>"
            msg['To'] = recipient_email
            msg['Subject'] = subject
            msg.attach(MIMEText(html_content, 'html', 'utf-8'))

            # SMTP/PMTA connection
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.set_debuglevel(1)

            # Login only if password exists
            if sender_password:
                server.starttls()
                server.login(sender_email, sender_password)

            server.sendmail(sender_email, recipient_email, msg.as_string())
            server.quit()

            logging.info(f"✅ Sent email to {recipient_email} via {sender_email}")
            time.sleep(delay)

            with state_lock:
                sent_emails += 1
                current_index = i + 1

            config_index = (config_index + 1) % len(email_configs)
            count += 1

        except Exception as e:
            logging.error(f"❌ Failed to send email to {recipient_email} using {sender_email}: {e}")
            continue

    logging.info(f"Finished sending emails. Total sent: {count}")

# Routes
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        return redirect(url_for('config'))
    return render_template('index.html')

@app.route('/config', methods=['GET', 'POST'])
def config():
    if request.method == 'POST':
        try:
            sender_email = request.form['sender_email']
            smtp_server = request.form['smtp_server']
            smtp_port = int(request.form['smtp_port'])
            sender_password = request.form.get('sender_password', '').strip()

            new_config = {
                'sender_email': sender_email,
                'smtp_server': smtp_server,
                'smtp_port': smtp_port,
                'sender_password': sender_password if sender_password else None
            }
            save_configuration(new_config)
            return redirect(url_for('send_email'))
        except Exception as e:
            return f"Error: {str(e)}", 500

    return render_template('configure.html', configurations=load_configurations())

@app.route('/send', methods=['GET', 'POST'])
def send_email():
    global paused, stop_sending, current_index, total_emails, sent_emails, email_thread
    configurations = load_configurations()

    if request.method == 'POST':
        try:
            delay = float(request.form.get('delay', '5'))
            min_limit = int(request.form.get('min_limit', '10'))
            max_limit = int(request.form.get('max_limit', '100'))

            excel_file = request.files.get('excel_file')
            excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(excel_file.filename))
            excel_file.save(excel_file_path)

            if email_thread and email_thread.is_alive():
                logging.info("Already sending emails.")
                return render_template('send.html', configurations=configurations, sending=True)

            paused = False
            stop_sending = False
            current_index = min_limit - 1
            total_emails = 0
            sent_emails = 0

            email_thread = threading.Thread(target=send_bulk_emails, args=(excel_file_path, delay, min_limit, max_limit))
            email_thread.start()

            return render_template('send.html', configurations=configurations, sending=True)
        except Exception as e:
            logging.error(f"Error: {e}")
            return f"Error: {e}", 500

    return render_template('send.html', configurations=configurations, sending=email_thread and email_thread.is_alive())

if __name__ == "__main__":
    app.run()
