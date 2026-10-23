"""Manda por e-mail a escala de plantão da semana (segunda a domingo).
Usado pelo workflow email-escala-plantao.yml toda segunda de manhã.

Variáveis de ambiente:
  SUPABASE_DB_URL  conexão com o banco (mesmo secret do manter-acordado)
  EMAIL_PARA       destinatários separados por vírgula
  POWER_AUTOMATE_URL  endereço do gatilho "Quando uma solicitação HTTP é
                   recebida" do Power Automate. Com ele, o e-mail sai pela conta
                   Microsoft 365 da GOL (o SMTP é bloqueado no tenant) e os
                   campos SMTP_* não são usados.
  SMTP_USER        e-mail que envia (login no servidor SMTP)
  SMTP_PASSWORD    senha desse e-mail
  SMTP_HOST        padrão smtp.office365.com (Outlook / Microsoft 365); para
                   Gmail use smtp.gmail.com com uma senha de app
  SMTP_PORT        padrão 587 (STARTTLS)
Com --previa, só grava o HTML em escala_semana.html, sem enviar.
Com --rss PASTA, grava PASTA/escala.xml (feed RSS com a escala da semana como
único item), sem enviar. O workflow publica esse feed no GitHub Pages num
endereço secreto e um fluxo padrão do Power Automate (conector RSS, sem
Premium) lê o feed toda segunda e manda o e-mail pela conta da GOL.
"""
import datetime
import email.utils
import html
import json
import os
import smtplib
import sys
from email.message import EmailMessage

import psycopg

FUSO_BRASILIA = datetime.timezone(datetime.timedelta(hours=-3))
DIAS = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
DIAS_CURTOS = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]
CATEGORIAS = {1: "Semana", 2: "Fim de semana", 3: "Feriado"}
# Mesmas cores do app (CORES_COLABORADORES_PLANTAO), na mesma ordem alfabética.
CORES = [
    "#4E79A7", "#E15759", "#59A14F", "#B07AA1", "#76B7B2",
    "#EDC948", "#9C755F", "#FF9DA7", "#2F4B7C", "#BAB0AC",
]


def semana_atual():
    hoje = datetime.datetime.now(FUSO_BRASILIA).date()
    inicio = hoje - datetime.timedelta(days=hoje.weekday())
    return inicio, inicio + datetime.timedelta(days=6)


def carregar_semana(inicio, fim):
    url = os.environ["SUPABASE_DB_URL"].replace("postgresql+psycopg://", "postgresql://")
    with psycopg.connect(url, prepare_threshold=None, connect_timeout=30) as conn:
        linhas = conn.execute(
            """SELECT data, colaborador, categoria, escalado_original, motivo
               FROM plantao_escala WHERE data BETWEEN %s AND %s ORDER BY data""",
            (inicio, fim),
        ).fetchall()
        nomes = [n for (n,) in conn.execute("SELECT DISTINCT colaborador FROM plantao_escala")]
    return linhas, sorted(nomes, key=str.lower)


def cor(nome, nomes):
    indice = nomes.index(nome) if nome in nomes else len(nomes)
    return CORES[indice % len(CORES)]


def agrupar(linhas):
    """Dias seguidos da mesma pessoa viram um plantão só."""
    blocos = []
    for data, colaborador, categoria, original, motivo in linhas:
        ultimo = blocos[-1] if blocos else None
        if ultimo and ultimo["nome"] == colaborador and (data - ultimo["fim"]).days == 1:
            ultimo["fim"] = data
        else:
            ultimo = {"nome": colaborador, "inicio": data, "fim": data,
                      "categorias": set(), "trocas": set()}
            blocos.append(ultimo)
        ultimo["categorias"].add(categoria)
        if original:
            ultimo["trocas"].add(f"no lugar de {original}")
        elif motivo:
            ultimo["trocas"].add(motivo)
    return blocos


def montar_html(inicio, fim, linhas, nomes):
    fonte = "font-family:Segoe UI,Arial,Helvetica,sans-serif;"
    if not linhas:
        conteudo = (
            f"<tr><td style='padding:24px 28px;{fonte}color:#6B7280;font-size:14px'>"
            "A escala desta semana ainda não está cadastrada no site.</td></tr>"
        )
    else:
        itens = []
        for bloco in agrupar(linhas):
            c = cor(bloco["nome"], nomes)
            i, f = bloco["inicio"], bloco["fim"]
            if i == f:
                quando = f"{DIAS[i.weekday()]}, {i:%d/%m}"
            else:
                # Todos os dias por extenso: "Terça, quarta e quinta · 20/10 a 22/10".
                dias = [DIAS[(i + datetime.timedelta(days=n)).weekday()].lower()
                        for n in range((f - i).days + 1)]
                nomes_dias = ", ".join(dias[:-1]) + " e " + dias[-1]
                quando = f"{nomes_dias.capitalize()} · {i:%d/%m} a {f:%d/%m}"
            etiquetas = "".join(
                "<span style='display:inline-block;margin-left:6px;padding:2px 8px;border-radius:10px;"
                "font-size:11px;font-weight:600;"
                + ("background:#FFF1E8;color:#C2410C;" if cat == 3 else "background:#F1F3F6;color:#6B7280;")
                + f"'>{CATEGORIAS.get(cat, '')}</span>"
                for cat in sorted(bloco["categorias"]) if cat != 1
            )
            troca = (
                f"<div style='font-size:12px;color:#8A5A00;margin-top:3px'>⇄ "
                f"{html.escape(', '.join(sorted(bloco['trocas'])))}</div>"
                if bloco["trocas"] else ""
            )
            itens.append(
                "<tr><td style='padding:6px 28px'>"
                "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' "
                f"style='background:{c}24;border-left:4px solid {c};border-radius:8px'>"
                f"<tr><td style='padding:12px 16px;{fonte}'>"
                f"<div style='font-size:12px;color:#6B7280;text-transform:uppercase;letter-spacing:.04em'>{quando}</div>"
                f"<div style='font-size:16px;font-weight:600;color:#1F2430;margin-top:2px'>"
                f"{html.escape(bloco['nome'])}{etiquetas}</div>{troca}"
                "</td></tr></table></td></tr>"
            )
        conteudo = "".join(itens)
    return f"""<!doctype html><html><body style="margin:0;padding:0;background:#F4F5F7">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F4F5F7;padding:24px 0">
<tr><td align="center">
<table role="presentation" width="560" cellpadding="0" cellspacing="0"
 style="max-width:560px;width:100%;background:#FFFFFF;border-radius:12px;overflow:hidden;border:1px solid #E6E8EC">
<tr><td style="height:4px;background:#FF6A13;font-size:0;line-height:0">&nbsp;</td></tr>
<tr><td style="padding:24px 28px 8px;{fonte}">
<div style="font-size:12px;color:#FF6A13;font-weight:700;letter-spacing:.08em;text-transform:uppercase">Engenharia GOL</div>
<div style="font-size:22px;font-weight:700;color:#1F2430;margin-top:4px">Plantão da semana</div>
<div style="font-size:14px;color:#6B7280;margin-top:2px">{inicio:%d/%m} a {fim:%d/%m/%Y}</div>
</td></tr>
<tr><td style="height:10px;font-size:0;line-height:0">&nbsp;</td></tr>
{conteudo}
<tr><td style="height:18px;font-size:0;line-height:0">&nbsp;</td></tr>
</table></td></tr></table></body></html>"""


def destinatarios():
    return [e.strip() for e in os.environ["EMAIL_PARA"].split(",") if e.strip()]


def enviar_power_automate(assunto, conteudo_html):
    import urllib.request

    corpo = json.dumps({
        "assunto": assunto,
        "html": conteudo_html,
        "para": "; ".join(destinatarios()),
    }).encode("utf-8")
    pedido = urllib.request.Request(
        os.environ["POWER_AUTOMATE_URL"], data=corpo, method="POST",
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(pedido, timeout=60) as resposta:
        print(f"Power Automate respondeu HTTP {resposta.status}.")


def enviar(assunto, conteudo_html):
    if os.environ.get("POWER_AUTOMATE_URL"):
        enviar_power_automate(assunto, conteudo_html)
        return
    usuario = os.environ["SMTP_USER"]
    destinatarios_email = destinatarios()
    mensagem = EmailMessage()
    mensagem["Subject"] = assunto
    mensagem["From"] = usuario
    mensagem["To"] = ", ".join(destinatarios_email)
    mensagem.set_content("Escala de plantão da semana (abra em um leitor com HTML).")
    mensagem.add_alternative(conteudo_html, subtype="html")
    # Cópia do HTML em anexo: o fluxo do Power Automate que reencaminha pela conta
    # da GOL usa o anexo como corpo, então o aviso de "e-mail externo" que o
    # Outlook coloca no corpo da mensagem recebida não vai junto.
    mensagem.add_attachment(
        conteudo_html.encode("utf-8"), maintype="text", subtype="html",
        filename="escala.html",
    )
    host = os.environ.get("SMTP_HOST") or "smtp.office365.com"
    porta = int(os.environ.get("SMTP_PORT") or 587)
    with smtplib.SMTP(host, porta, timeout=60) as smtp:
        smtp.starttls()
        smtp.login(usuario, os.environ["SMTP_PASSWORD"])
        smtp.send_message(mensagem)
    print(f"E-mail enviado para {len(destinatarios_email)} destinatário(s).")


def gravar_rss(pasta, assunto, conteudo_html, inicio):
    publicado = datetime.datetime.combine(inicio, datetime.time(6), FUSO_BRASILIA)
    # O HTML vai num CDATA; "]]>" fecharia o bloco antes da hora.
    descricao = conteudo_html.replace("]]>", "]]]]><![CDATA[>")
    feed = f"""<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0"><channel>
<title>Escala de plantão da Engenharia</title>
<link>https://github.com/</link>
<description>Escala de plantão da semana</description>
<item>
<title>{html.escape(assunto)}</title>
<guid isPermaLink="false">plantao-{inicio:%Y-%m-%d}</guid>
<pubDate>{email.utils.format_datetime(publicado)}</pubDate>
<description><![CDATA[{descricao}]]></description>
</item>
</channel></rss>
"""
    os.makedirs(pasta, exist_ok=True)
    with open(os.path.join(pasta, "escala.xml"), "w", encoding="utf-8") as arquivo:
        arquivo.write(feed)


def main():
    inicio, fim = semana_atual()
    linhas, nomes = carregar_semana(inicio, fim)
    conteudo = montar_html(inicio, fim, linhas, nomes)
    assunto = f"Escala de plantão da Engenharia: {inicio:%d/%m} a {fim:%d/%m}"
    if "--previa" in sys.argv:
        with open("escala_semana.html", "w", encoding="utf-8") as arquivo:
            arquivo.write(conteudo)
        print(f"{assunto}: {len(linhas)} dia(s). Prévia em escala_semana.html")
        return
    if "--rss" in sys.argv:
        pasta = sys.argv[sys.argv.index("--rss") + 1]
        gravar_rss(pasta, assunto, conteudo, inicio)
        print(f"{assunto}: {len(linhas)} dia(s). Feed gravado.")
        return
    enviar(assunto, conteudo)


if __name__ == "__main__":
    main()
