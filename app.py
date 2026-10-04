import os
import io
import re
import time
import uuid
import threading
import builtins
from datetime import datetime
import openpyxl
from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for
import Disussion_Automate as da
import sheets_helper

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'edgewood_secret_key_123')

# Global state for live updates
task_progress = {}
task_results = {}
current_task_id = None

# Intercept print to stream live updates natively
original_print = builtins.print

def custom_print(*args, **kwargs):
    global current_task_id
    text = " ".join(str(a) for a in args)
    if current_task_id and text.strip():
        clean_text = text.strip()
        # Make the prints look cooler and more readable for the UI
        if "Processing thread" in clean_text:
            clean_text = clean_text.replace("Processing thread:", "Analyzing Topic:")
        elif "Fetching" in clean_text:
            clean_text = f"Connecting to Canvas: {clean_text}"
            
        # Ignore empty lines or simple separators
        if not set(clean_text).issubset({'-', '=', ' '}):
            task_progress[current_task_id] = clean_text
            
    original_print(*args, **kwargs)

builtins.print = custom_print

def run_audit(task_id, course_codes_input):
    global current_task_id
    current_task_id = task_id
    task_progress[task_id] = "Initializing Canvas API Connection..."
    
    try:
        if "," in course_codes_input:
            sis_ids = [s.strip() for s in course_codes_input.split(",") if s.strip()]
        else:
            sis_ids = [s.strip() for s in course_codes_input.split() if s.strip()]

        if not sis_ids:
            task_progress[task_id] = "ERROR: No valid SIS IDs found in input."
            return

        if not da.check_connection():
            task_progress[task_id] = "ERROR: API connection failed. Check your token."
            return

        # Fetch latest TA directory from Google Sheet to ensure new TAs are immediately active
        task_progress[task_id] = "Syncing Active TA Directory from Google Sheet..."
        sheets_helper.get_import_data_ta(force_refresh=True)

        wb = openpyxl.Workbook()
        processed_count = 0

        for sis_id in sis_ids:
            # Education Tracker specific filter (handles ADO typo / speech-to-text)
            clean_sis = sis_id.strip()
            clean_sis_fixed = re.sub(r'\bADO\b', 'EDU', clean_sis, flags=re.IGNORECASE).replace("ADO_", "EDU_").replace("ADO-", "EDU-")
            if not clean_sis_fixed.upper().startswith("EDU"):
                task_progress[task_id] = f"Skipping {sis_id} (Not an Education course)."
                time.sleep(1)
                continue
            sis_id = clean_sis_fixed

            task_progress[task_id] = f"Locating Course: {sis_id}..."
            course = da.find_course(sis_id)
            if not course:
                task_progress[task_id] = f"Could not find Course: {sis_id}. Skipping..."
                time.sleep(1)
                continue
                
            course_id = course["id"]
            course_name = course.get("name", "Unknown Course")
            
            # This calls the original function which has all the print statements we intercept!
            flat_rows = da.collect_consolidated_rows(course_id, sis_id, course_name)
            
            has_real_data = any(r.get("Learner Name") not in ("(No student posts yet)", "N/A", "Unknown", "User#None", "") for r in flat_rows)

            if not flat_rows or not has_real_data:
                task_progress[task_id] = f"No discussions found in {sis_id}. Skipping..."
                time.sleep(1)
                continue

            task_progress[task_id] = "Compiling Data into Excel Structure..."
            da.build_excel_sheet(wb, flat_rows, sis_id, course_name)
            try:
                task_progress[task_id] = "Syncing Report to Google Sheet Database..."
                sheets_helper.push_to_google_sheet(flat_rows, sis_id, course_name)
                # Bust the dashboard cache so next fetch reflects new audit data
                sheets_helper._dashboard_cache["data"] = None
            except Exception as e:
                original_print(f"Error pushing to Google Sheets: {e}")
            processed_count += 1


        if processed_count == 0:
            task_progress[task_id] = "ERROR: No student discussions found in any course."
            return

        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb["Sheet"]

        task_progress[task_id] = "Finalizing Output File..."
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        task_results[task_id] = {
            "file": output,
            "filename": f"MultiCourse_Discussion_Audit_{sis_ids[0].replace(' ', '_')}.xlsx"
        }
        
        time.sleep(0.5) # Slight pause to let UI catch up
        task_progress[task_id] = "COMPLETE"
    except Exception as e:
        task_progress[task_id] = f"ERROR: {str(e)}"
    finally:
        current_task_id = None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/start', methods=['POST'])
def start_audit():
    data = request.json
    course_codes_input = data.get('course_codes', '')
    if not course_codes_input:
        return jsonify({"error": "Please enter at least one course code."}), 400
        
    task_id = str(uuid.uuid4())
    task_progress[task_id] = "Warming up Audit Engine..."
    
    # Run the scraping in a background thread
    thread = threading.Thread(target=run_audit, args=(task_id, course_codes_input))
    thread.start()
    
    return jsonify({"task_id": task_id})

@app.route('/status/<task_id>')
def status(task_id):
    state = task_progress.get(task_id, "Unknown Task")
    return jsonify({"status": state})

@app.route('/download/<task_id>')
def download(task_id):
    if task_id in task_results:
        res = task_results[task_id]
        return send_file(
            res["file"],
            as_attachment=True,
            download_name=res["filename"],
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return "File not found or expired.", 404

@app.route('/api/dashboard')
def api_dashboard():
    try:
        data = sheets_helper.get_dashboard_data()
        return jsonify(data)
    except Exception as e:
        original_print(f"Dashboard data fetch failed: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/api/verify-credentials', methods=['POST'])
def verify_credentials():
    data = request.json or {}
    sender_email = (data.get('sender_email') or '').strip()
    app_password = (data.get('app_password') or '').strip().replace(' ', '')
    if not sender_email or not app_password:
        return jsonify({"valid": False, "error": "Email and App Password are required."}), 400

    import smtplib, ssl
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context, timeout=10) as server:
            server.login(sender_email, app_password)
        return jsonify({"valid": True, "message": "Credentials verified via Google SMTP."})
    except smtplib.SMTPAuthenticationError:
        return jsonify({"valid": False, "error": "Authentication failed. Check your email or 16-character Google App Password."}), 401
    except Exception as e:
        return jsonify({"valid": False, "error": f"Verification error: {str(e)}"}), 500

@app.route('/api/scan-pending-ta-emails', methods=['POST'])
def scan_pending_ta_emails():
    data = request.json or {}
    course_shell = (data.get('course_shell') or '').strip()
    if not course_shell:
        return jsonify({"error": "Please provide a valid Course Shell ID."}), 400

    try:
        summary = sheets_helper.get_pending_ta_summary(course_shell)
        return jsonify(summary)
    except Exception as e:
        original_print(f"Error scanning pending TA emails: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/api/send-ta-email', methods=['POST'])
def send_ta_email():
    data = request.json or {}
    sender_email  = (data.get('sender_email') or '').strip()
    app_password  = (data.get('app_password') or '').strip().replace(' ', '')
    to_email      = (data.get('to_email') or '').strip()
    cc_emails_raw = (data.get('cc_emails') or '').strip()
    subject       = (data.get('subject') or '').strip()
    body_html     = data.get('body_html', '')
    body_text     = data.get('body_text', '')
    ta_name       = (data.get('ta_name') or 'TA').strip()
    is_test       = bool(data.get('is_test', False))

    if not sender_email or not app_password:
        return jsonify({"success": False, "error": "Sender email and App Password must be configured in Settings."}), 400
    if not to_email:
        return jsonify({"success": False, "error": "Recipient TA email is required."}), 400

    import smtplib, ssl
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.utils import formataddr

    # ── Test Mode: redirect to sender's own inbox ──────────────────────────────
    actual_to = sender_email if is_test else to_email
    actual_cc = []           if is_test else [c.strip() for c in cc_emails_raw.replace(';', ',').split(',') if c.strip()]
    effective_subject = f"[TEST PREVIEW] {subject}" if is_test else subject or "[Action Required] Pending Discussion Forum Queries"
    test_banner_html = ""
    test_banner_text = ""
    if is_test:
        test_banner_html = """
        <div style="background:#fef3c7;border:2px dashed #d97706;border-radius:8px;padding:12px 16px;margin-bottom:20px;font-family:sans-serif;">
          <strong style="color:#92400e;">⚠ TEST PREVIEW — This email was NOT sent to the TA.</strong><br>
          <span style="font-size:12px;color:#78350f;">Original recipient: <code>{to}</code>. You are previewing how this email renders in your inbox.</span>
        </div>""".format(to=to_email)
        test_banner_text = f"\n--- TEST PREVIEW (original recipient: {to_email}) ---\n\n"

    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = effective_subject
        sender_display = data.get('full_name') or "Edgewood Operations"
        msg['From'] = formataddr((sender_display, sender_email))
        msg['To'] = formataddr(("[TEST] " + ta_name if is_test else ta_name, actual_to))
        if actual_cc:
            msg['Cc'] = ", ".join(actual_cc)

        if body_text:
            msg.attach(MIMEText(test_banner_text + body_text, 'plain', 'utf-8'))
        if body_html:
            if is_test:
                import re
                if '<body' in body_html.lower():
                    injected_html = re.sub(r'(<body[^>]*>)', r'\1' + test_banner_html, body_html, count=1, flags=re.IGNORECASE)
                else:
                    injected_html = test_banner_html + body_html
            else:
                injected_html = body_html
            msg.attach(MIMEText(injected_html, 'html', 'utf-8'))

        context = ssl.create_default_context()
        recipients = [actual_to] + actual_cc
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context, timeout=15) as server:
            server.login(sender_email, app_password)
            server.sendmail(sender_email, recipients, msg.as_string())

        now_str = datetime.now().strftime("%I:%M:%S %p")
        return jsonify({
            "success": True,
            "is_test": is_test,
            "recipient": actual_to,
            "original_recipient": to_email,
            "cc": actual_cc,
            "sent_at": now_str,
            "message": f"[TEST] Email preview sent to {sender_email}" if is_test else f"Email successfully sent to {ta_name} ({to_email})"
        })
    except smtplib.SMTPAuthenticationError:
        return jsonify({"success": False, "error": "SMTP Authentication failed. Check your App Password."}), 401
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.getenv("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port, threaded=True)
