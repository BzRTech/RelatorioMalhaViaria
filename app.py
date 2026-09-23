# -*- coding: utf-8 -*-
"""
Webapp do Relatório de Malha Viária (Flask) - pronto para o Render.

Fluxo: o usuário envia o shapefile (.zip ou os arquivos .shp/.dbf/.shx) ou um
CSV, informa o município e recebe para download o PDF, o Excel, os gráficos e
a apresentação (.pptx).

Variáveis de ambiente:
  APP_PASSWORD  (opcional) exige login com essa senha (usuário: qualquer)
  JOBS_DIR      (opcional) pasta onde os resultados ficam temporariamente
  JOB_TTL_HORAS (opcional) horas até apagar os resultados (padrão: 2)
"""

import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import uuid
import zipfile
from functools import wraps

from flask import (Flask, Response, abort, flash, redirect, render_template,
                   request, send_from_directory, url_for)
from werkzeug.utils import secure_filename

from tarefa import ProcessoInterrompido, gerar_isolado

LOGO_PADRAO = "logomarca-eixo-cores-2.png"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JOBS_DIR = os.environ.get("JOBS_DIR") or os.path.join(tempfile.gettempdir(), "malha_viaria_jobs")
JOB_TTL = float(os.environ.get("JOB_TTL_HORAS", "2")) * 3600
SENHA = os.environ.get("APP_PASSWORD", "")
EXTENSOES = {".zip", ".rar", ".csv", ".shp", ".dbf", ".shx", ".prj", ".cpg",
             ".sbn", ".sbx", ".xml", ".qmd"}
EXT_LOGO = {".png", ".jpg", ".jpeg"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024  # 80 MB
app.secret_key = os.environ.get("SECRET_KEY") or os.urandom(24)

# Um relatório por vez: cada um roda num processo próprio de ~400 MB e o
# plano free do Render tem 512 MB
_trava = threading.Lock()


# --------------------------------------------------------------------------
# Autenticação simples (opcional)
# --------------------------------------------------------------------------

def exige_senha(view):
    @wraps(view)
    def envolvida(*args, **kwargs):
        if SENHA:
            auth = request.authorization
            if not auth or auth.password != SENHA:
                return Response("Acesso restrito.", 401,
                                {"WWW-Authenticate": 'Basic realm="Malha Viaria"'})
        return view(*args, **kwargs)
    return envolvida


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------

def _limpar_antigos():
    if not os.path.isdir(JOBS_DIR):
        return
    agora = time.time()
    for nome in os.listdir(JOBS_DIR):
        caminho = os.path.join(JOBS_DIR, nome)
        if os.path.isdir(caminho) and agora - os.path.getmtime(caminho) > JOB_TTL:
            shutil.rmtree(caminho, ignore_errors=True)


def _pasta_job(job_id):
    if not re.fullmatch(r"[0-9a-f]{32}", job_id or ""):
        abort(404)
    pasta = os.path.join(JOBS_DIR, job_id, "saida")
    if not os.path.isdir(pasta):
        abort(404)
    return pasta


def _definir_entrada(pasta_entrada, nomes):
    """Um único .zip/.rar/.csv vira a entrada; vários arquivos = a pasta (shapefile)."""
    if len(nomes) == 1:
        return os.path.join(pasta_entrada, nomes[0])
    exts = {os.path.splitext(n)[1].lower() for n in nomes}
    if ".shp" not in exts and ".dbf" not in exts and ".csv" not in exts:
        raise RuntimeError("Envie um .zip com o shapefile, os arquivos .shp + .dbf + .shx "
                           "juntos, ou um .csv.")
    return pasta_entrada


# --------------------------------------------------------------------------
# Rotas
# --------------------------------------------------------------------------

@app.get("/")
@exige_senha
def index():
    return render_template("index.html")


@app.post("/gerar")
@exige_senha
def gerar():
    _limpar_antigos()
    municipio = (request.form.get("municipio") or "").strip()
    arquivos = [f for f in request.files.getlist("dados") if f and f.filename]
    if not municipio:
        flash("Informe o nome do município.")
        return redirect(url_for("index"))
    if not arquivos:
        flash("Selecione o arquivo de dados.")
        return redirect(url_for("index"))

    job_id = uuid.uuid4().hex
    raiz = os.path.join(JOBS_DIR, job_id)
    pasta_entrada = os.path.join(raiz, "entrada")
    pasta_saida = os.path.join(raiz, "saida")
    os.makedirs(pasta_entrada)

    try:
        nomes = []
        for f in arquivos:
            nome = secure_filename(f.filename) or "arquivo"
            if os.path.splitext(nome)[1].lower() not in EXTENSOES:
                raise RuntimeError(f"Tipo de arquivo não aceito: {f.filename}")
            f.save(os.path.join(pasta_entrada, nome))
            nomes.append(nome)
        entrada = _definir_entrada(pasta_entrada, nomes)

        logo = os.path.join(BASE_DIR, LOGO_PADRAO)
        arq_logo = request.files.get("logo")
        if arq_logo and arq_logo.filename:
            ext = os.path.splitext(arq_logo.filename)[1].lower()
            if ext not in EXT_LOGO:
                raise RuntimeError("A logomarca deve ser .png ou .jpg.")
            logo = os.path.join(raiz, "logo" + ext)
            arq_logo.save(logo)

        with _trava:
            res = gerar_isolado(entrada, municipio, pasta_saida, logo_path=logo)
    except Exception as e:  # noqa: BLE001 - mostra o erro ao usuário
        app.logger.exception("Falha ao gerar relatório")
        shutil.rmtree(raiz, ignore_errors=True)
        if isinstance(e, ProcessoInterrompido):
            msg = ("o processamento foi interrompido pelo servidor (provavelmente falta "
                   "de memória). Tente de novo; se persistir, o arquivo é grande demais "
                   "para o plano atual.")
        elif isinstance(e, subprocess.TimeoutExpired):
            msg = "o processamento demorou demais e foi cancelado."
        else:
            msg = str(e)
        flash(f"Não foi possível gerar o relatório: {msg}")
        return redirect(url_for("index"))
    finally:
        shutil.rmtree(pasta_entrada, ignore_errors=True)

    # Pacote com tudo
    with zipfile.ZipFile(os.path.join(pasta_saida, "tudo.zip"), "w", zipfile.ZIP_DEFLATED) as z:
        for arq in res["arquivos"].values():
            z.write(arq, os.path.basename(arq))

    info = {"municipio": municipio, "data": res["data"], "resumo": res["resumo"],
            "arquivos": {k: os.path.basename(v) for k, v in res["arquivos"].items()}}
    with open(os.path.join(pasta_saida, "info.json"), "w", encoding="utf-8") as fh:
        json.dump(info, fh, ensure_ascii=False)
    return redirect(url_for("resultado", job_id=job_id))


@app.get("/resultado/<job_id>")
@exige_senha
def resultado(job_id):
    pasta = _pasta_job(job_id)
    with open(os.path.join(pasta, "info.json"), encoding="utf-8") as fh:
        info = json.load(fh)
    return render_template("resultado.html", job_id=job_id, info=info,
                           ttl_horas=int(JOB_TTL // 3600))


@app.get("/baixar/<job_id>/<path:nome>")
@exige_senha
def baixar(job_id, nome):
    pasta = _pasta_job(job_id)
    if nome == "tudo.zip":
        with open(os.path.join(pasta, "info.json"), encoding="utf-8") as fh:
            municipio = json.load(fh)["municipio"]
        return send_from_directory(pasta, nome, as_attachment=True,
                                   download_name=f"Malha_Viaria_{secure_filename(municipio)}.zip")
    return send_from_directory(pasta, nome, as_attachment=True)


@app.get("/saude")
def saude():
    return {"ok": True}


@app.errorhandler(413)
def muito_grande(_):
    flash("Arquivo muito grande (limite de 80 MB).")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))
