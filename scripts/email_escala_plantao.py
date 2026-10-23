"""Manda por e-mail a escala de plantão da semana (segunda a domingo).
Usado pelo workflow email-escala-plantao.yml toda segunda de manhã.

Variáveis de ambiente:
  SUPABASE_DB_URL  conexão com o banco (mesmo secret do manter-acordado)
  SMTP_USER        e-mail que envia (login no servidor SMTP)
  SMTP_PASSWORD    senha desse e-mail
  EMAIL_PARA       destinatários separados por vírgula
  SMTP_HOST        padrão smtp.office365.com (Outlook / Microsoft 365)
  SMTP_PORT        padrão 587 (STARTTLS)
Com --previa, só grava o HTML em escala_semana.html, sem enviar.
"""
import datetime
import html
import os
import smtplib
import sys
from email.message import EmailMessage

import psycopg

FUSO_BRASILIA = datetime.timezone(datetime.timedelta(hours=-3))
DIAS = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
CATEGORIAS = {1: "Semana", 2: "Fim de semana", 3: "Feriado / ponte"}


def semana_atual():
    hoje = datetime.datetime.now(FUSO_BRASILIA).date()
    inicio = hoje - datetime.timedelta(days=hoje.weekday())
    return inicio, inicio + datetime.timedelta(days=6)


def carregar_semana(inicio, fim):
    url = os.environ["SUPABASE_DB_URL"].replace("postgresql+psycopg://", "postgresql://")
    with psycopg.connect(url, prepare_threshold=None, connect_timeout=30) as conn:
        return conn.execute(
            """SELECT data, colaborador, categoria, escalado_original, motivo
               FROM plantao_escala WHERE data BETWEEN %s AND %s ORDER BY data""",
            (inicio, fim),
        ).fetchall()


def montar_html(inicio, fim, linhas):
    if not linhas:
        corpo = "<p>A escala desta semana ainda não está cadastrada no site.</p>"
    else:
        celula = "padding:8px 12px;border-bottom:1px solid #E6E8EC;"
        tabela = []
        for data, colaborador, categoria, original, motivo in linhas:
            observacao = []
            if original:
                observacao.append(f"no lugar de {original}")
            if motivo:
                observacao.append(motivo)
            tabela.append(
                f"<tr><td style='{celula}'>{DIAS[data.weekday()]}, {data:%d/%m}</td>"
                f"<td style='{celula}font-weight:600'>{html.escape(colaborador)}</td>"
                f"<td style='{celula}color:#6B7280'>{CATEGORIAS.get(categoria, '')}</td>"
                f"<td style='{celula}color:#6B7280'>{html.escape(' · '.join(observacao))}</td></tr>"
            )
        cabecalho = "".join(
            f"<th style='text-align:left;padding:8px 12px;color:#6B7280;font-weight:600;"
            f"border-bottom:2px solid #FF6A13'>{t}</th>"
            for t in ("Dia", "Plantão", "Categoria", "Observação")
        )
        corpo = (
            "<table style='border-collapse:collapse;font-size:14px'>"
            f"<tr>{cabecalho}</tr>{''.join(tabela)}</table>"
        )
    return (
        "<div style='font-family:Segoe UI,Arial,sans-serif;color:#1F2430'>"
        f"<p>Bom dia! Segue a escala de plantão da Engenharia de "
        f"{inicio:%d/%m} a {fim:%d/%m/%Y}:</p>{corpo}"
        "<p style='color:#6B7280;font-size:12px;margin-top:16px'>Trocas e detalhes na tela "
        "Escala de plantão do app Pesagem e Balanceamento. E-mail automático.</p></div>"
    )


def enviar(assunto, conteudo_html):
    usuario = os.environ["SMTP_USER"]
    destinatarios = [e.strip() for e in os.environ["EMAIL_PARA"].split(",") if e.strip()]
    mensagem = EmailMessage()
    mensagem["Subject"] = assunto
    mensagem["From"] = usuario
    mensagem["To"] = ", ".join(destinatarios)
    mensagem.set_content("Escala de plantão da semana (abra em um leitor com HTML).")
    mensagem.add_alternative(conteudo_html, subtype="html")
    host = os.environ.get("SMTP_HOST") or "smtp.office365.com"
    porta = int(os.environ.get("SMTP_PORT") or 587)
    with smtplib.SMTP(host, porta, timeout=60) as smtp:
        smtp.starttls()
        smtp.login(usuario, os.environ["SMTP_PASSWORD"])
        smtp.send_message(mensagem)
    print(f"E-mail enviado para {len(destinatarios)} destinatário(s).")


def main():
    inicio, fim = semana_atual()
    linhas = carregar_semana(inicio, fim)
    conteudo = montar_html(inicio, fim, linhas)
    assunto = f"Escala de plantão da Engenharia: {inicio:%d/%m} a {fim:%d/%m}"
    if "--previa" in sys.argv:
        with open("escala_semana.html", "w", encoding="utf-8") as arquivo:
            arquivo.write(conteudo)
        print(f"{assunto}: {len(linhas)} dia(s). Prévia em escala_semana.html")
        return
    enviar(assunto, conteudo)


if __name__ == "__main__":
    main()
