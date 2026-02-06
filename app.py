from flask import Flask, render_template, redirect, url_for, request, send_file
import os

app = Flask(__name__)


@app.route("/")
@app.route("/login", methods=["GET"])
def login_page():
    return render_template("login.html")


@app.route("/login", methods=["POST"])
def login_submit():
    return redirect(url_for("download_page"))


@app.route("/download")
def download_page():
    return render_template("download.html")


@app.route("/download/csv")
def download_csv():
    csv_path = os.path.join(app.static_folder, "sample.csv")
    return send_file(csv_path, as_attachment=True, download_name="sample.csv")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
