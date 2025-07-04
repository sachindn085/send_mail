from flask import Flask, request, jsonify
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

app = Flask(__name__)

@app.route("/send-emails", methods=["POST"])
def send_emails():
    file = request.files.get("file")
    gmail_user = request.form.get("gmail")
    gmail_app_password = request.form.get("password")

    if not file:
        return jsonify({"error": "XLSX file is required."}), 400
    if not gmail_user or not gmail_app_password:
        return jsonify({"error": "Both 'gmail' and 'password' are required."}), 400

    try:
        df = pd.read_excel(file)

        required_columns = ["email", "email_subject", "email_body", "score"]
        if not all(col in df.columns for col in required_columns):
            return jsonify({"error": f"Missing required columns. Found columns: {list(df.columns)}"}), 400

        success = []
        failed = []
        skipped_due_to_score = []

        for _, row in df.iterrows():
            to_email = row.get("email")
            subject = row.get("email_subject")
            body = row.get("email_body")
            score = row.get("score")

            if pd.isna(to_email) or pd.isna(subject) or pd.isna(body) or pd.isna(score):
                failed.append({"email": to_email, "error": "Missing email, subject, body, or score."})
                continue

            if float(score) < 70:
                skipped_due_to_score.append({"email": to_email, "score": score})
                continue

            msg = MIMEMultipart()
            msg["From"] = gmail_user
            msg["To"] = to_email
            msg["Subject"] = subject
            msg.attach(MIMEText(str(body), "plain"))

            try:
                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                    server.login(gmail_user, gmail_app_password)
                    server.sendmail(gmail_user, to_email, msg.as_string())
                success.append(to_email)
            except Exception as e:
                failed.append({"email": to_email, "error": str(e)})

        return jsonify({
            "sent": success,
            "failed": failed,
            "skipped_due_to_low_score": skipped_due_to_score
        }), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
