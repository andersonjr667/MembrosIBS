from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import json, os
import pandas as pd
from datetime import datetime, timedelta

app = Flask(__name__)
app.secret_key = "segredo123"
app.permanent_session_lifetime = timedelta(minutes=30)

USERS_FILE = "data/usuarios.json"
MEMBROS_FILE = "data/membros.xlsx"

def carregar_usuarios():
    if not os.path.exists(USERS_FILE):
        return []
    with open(USERS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def carregar_membros():
    if not os.path.exists(MEMBROS_FILE):
        cols = ["Nome","Pai","Mãe","Data de nascimento","Natural de","Nacionalidade",
                "Estado","País","Estado civil","Nome do conjuge","Endereço","Bairro",
                "Cidade","CEP","Telefone residencial","Celular"]
        df = pd.DataFrame(columns=cols)
        df.to_excel(MEMBROS_FILE, index=False)
    return pd.read_excel(MEMBROS_FILE)

def salvar_membros(df):
    df.to_excel(MEMBROS_FILE, index=False)

# --- Login ---
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form["usuario"]
        senha = request.form["senha"]
        usuarios = carregar_usuarios()
        for u in usuarios:
            if u["usuario"] == usuario and u["senha"] == senha:
                session["usuario"] = usuario
                flash("Login realizado com sucesso!", "success")
                return redirect(url_for("index"))
        flash("Usuário ou senha inválidos!", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.pop("usuario", None)
    flash("Você saiu do sistema.", "info")
    return redirect(url_for("login"))

@app.before_request
def proteger_rotas():
    rotas_livres = ["login", "static"]
    if "usuario" not in session and request.endpoint not in rotas_livres:
        return redirect(url_for("login"))

# --- Páginas ---
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/membros")
def membros():
    df = carregar_membros()
    return render_template("membros.html", membros=df.to_dict(orient="records"))

@app.route("/novo_membro", methods=["GET","POST"])
def novo_membro():
    if request.method == "POST":
        df = carregar_membros()
        novo = {col: request.form.get(col,"") for col in df.columns}
        df = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
        salvar_membros(df)
        flash("Membro adicionado com sucesso!", "success")
        return redirect(url_for("membros"))
    df = carregar_membros()
    return render_template("novo_membro.html", campos=df.columns)

@app.route("/aniversariantes", methods=["GET","POST"])
def aniversariantes():
    df = carregar_membros()
    aniversarios = []
    mes = None
    if request.method == "POST":
        mes = int(request.form["mes"])
        for _, row in df.iterrows():
            try:
                data = pd.to_datetime(row["Data de nascimento"], errors="coerce")
                if not pd.isna(data) and data.month == mes:
                    aniversarios.append(row.to_dict())
            except:
                pass
    return render_template("aniversariantes.html", membros=aniversarios, mes=mes)

@app.route("/baixar_aniversariantes/<int:mes>")
def baixar_aniversariantes(mes):
    df = carregar_membros()
    aniversarios = []
    for _, row in df.iterrows():
        try:
            data = pd.to_datetime(row["Data de nascimento"], errors="coerce")
            if not pd.isna(data) and data.month == mes:
                aniversarios.append(row.to_dict())
        except:
            pass
    if aniversarios:
        df_mes = pd.DataFrame(aniversarios)
        arquivo = f"data/aniversariantes_{mes}.xlsx"
        df_mes.to_excel(arquivo, index=False)
        return send_file(arquivo, as_attachment=True)
    flash("Nenhum aniversariante encontrado para este mês.", "warning")
    return redirect(url_for("aniversariantes"))

if __name__ == "__main__":
    app.run(debug=True)
