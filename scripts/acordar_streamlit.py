"""Abre o app no Streamlit Cloud e clica em "Yes, get this app back up!"
quando ele estiver dormindo. Usado pelo workflow manter-acordado.yml."""
import re
import sys

from playwright.sync_api import sync_playwright

BOTAO = re.compile(r"get this app back up", re.IGNORECASE)


def procurar_botao(pagina):
    for quadro in pagina.frames:
        botao = quadro.get_by_role("button", name=BOTAO)
        if botao.count():
            return botao.first
    return None


def main(url):
    with sync_playwright() as p:
        navegador = p.chromium.launch()
        pagina = navegador.new_page()
        pagina.goto(url, wait_until="domcontentloaded", timeout=120_000)
        pagina.wait_for_timeout(15_000)
        botao = procurar_botao(pagina)
        if botao is None:
            print("App já estava acordado.")
        else:
            print("App dormindo; clicando para acordar.")
            botao.click()
            for _ in range(36):  # até 3 min
                pagina.wait_for_timeout(5_000)
                if procurar_botao(pagina) is None:
                    break
            else:
                print("O botão continuou na tela depois de 3 min.")
                return 1
            print("App acordando.")
        # Fica um pouco na página para a visita contar como uso.
        pagina.wait_for_timeout(30_000)
        navegador.close()
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
