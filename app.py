from flask import Flask, render_template, request, redirect, url_for, send_file, session
from NOMEACAO_NOVA import gerar_nomeacoes_web
import os

app = Flask(__name__)
app.secret_key = "supersecretkey123"

USERNAME = "admin"
PASSWORD = "123456"

EXCEL_FILE = "NOMEACOES_CEJUSC.xlsx"

# ------------------------
# LOGIN
# ------------------------
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        if username == USERNAME and password == PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("index"))
        else:
            return render_template("login.html", error="Usuário ou senha incorretos")

    return render_template("login.html")


# ------------------------
# PÁGINA PRINCIPAL
# ------------------------
@app.route("/index", methods=["GET", "POST"])
def index():
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    existentes = ""
    novos = ""
    message = None

    if request.method == "POST":
        existentes = request.form.get("existentes", "")
        novos = request.form.get("novos", "")

        try:
            gerar_nomeacoes_web(existentes, novos)

            # 🔥 FAZ DOWNLOAD AUTOMÁTICO
            return send_file(
                EXCEL_FILE,
                as_attachment=True
            )

        except Exception as e:
            message = f"Erro: {e}"

    return render_template(
        "index.html",
        message=message,
        existentes=existentes,
        novos=novos
    )


# ------------------------
# LOGOUT
# ------------------------
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ------------------------
# RENDER + LOCAL
# ------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)