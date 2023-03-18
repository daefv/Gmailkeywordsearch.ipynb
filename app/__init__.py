import os
from pathlib import Path

from flask import (
    Flask,
    redirect,
    render_template,
    request,
    send_from_directory,
    url_for,
)

from .email import generate_excel_file

app = Flask(__name__)


@app.route("/")
def index():
    values = request.form
    error = request.args.get("error")

    return render_template("index.html", values=values, error=error)


@app.route("/download", methods=["POST"])
def download():
    values = request.form

    try:
        # return send_from_directory(".", "sender_emails.xlsx")

        # ensure instance directory exists
        Path(app.instance_path).mkdir(parents=True, exist_ok=True)

        # remove old excel file
        Path(f"{app.instance_path}/sender_emails.xlsx").unlink(missing_ok=True)

        # create excel
        generate_excel_file(
            email_address=values.get("email"),
            password=values.get("app_password"),
            keyword_to_search=values.get("keyword"),
            emails_to_search=values.get("amount"),
            filepath=f"{app.instance_path}/sender_emails.xlsx",
        )

        # return excel for download
        return send_from_directory(".", "sender_emails.xlsx")

    except Exception as exc:
        return redirect(url_for("index", error=str(exc)))
