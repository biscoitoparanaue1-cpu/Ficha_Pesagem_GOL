import argparse
import json
import os
import sqlite3
import sys
from contextlib import closing
from pathlib import Path

TABLES = (
    "pesagens",
    "usuarios",
    "fluxo_fichas",
    "campos_com_erro",
    "assinaturas_usuarios",
)
DEPENDENT_TABLES = (
    "fluxo_fichas",
    "campos_com_erro",
    "assinaturas_usuarios",
)
REQUIRED_COLUMNS = {
    "usuarios": {
        "usuario",
        "nome",
        "senha_hash",
        "salt",
        "nivel_acesso",
        "criado_em",
    },
    "fluxo_fichas": {
        "prefixo",
        "pesagem",
        "revisao",
        "gerado_por",
        "aprovado_por",
        "criado_em",
        "aprovado_em",
    },
    "campos_com_erro": {
        "prefixo",
        "pesagem",
        "revisao",
        "campo",
        "descricao",
        "marcado_por",
        "marcado_em",
    },
    "assinaturas_usuarios": {"usuario", "imagem_png", "atualizado_em"},
}


class MigrationError(Exception):
    pass


def ler_banco_sqlite(caminho):
    uri = f"{caminho.resolve().as_uri()}?mode=ro"
    try:
        with closing(sqlite3.connect(uri, uri=True)) as origem:
            integridade = origem.execute("PRAGMA integrity_check").fetchone()[0]
            if integridade != "ok":
                raise MigrationError("O SQLite falhou na verificação de integridade.")

            tabelas_disponiveis = {
                linha[0]
                for linha in origem.execute(
                    "SELECT name FROM sqlite_master WHERE type = 'table'"
                )
            }
            ausentes = set(TABLES) - tabelas_disponiveis
            if ausentes:
                raise MigrationError(
                    "O SQLite não contém as tabelas esperadas: "
                    + ", ".join(sorted(ausentes))
                )

            dados = {}
            for tabela in TABLES:
                cursor = origem.execute(f'SELECT * FROM "{tabela}"')
                colunas = [item[0] for item in cursor.description]
                obrigatorias = REQUIRED_COLUMNS.get(tabela, set())
                if tabela == "pesagens":
                    if not {"Prefixo", "Pesagem"} <= set(colunas) or not (
                        {"Revisao", "revisao"} & set(colunas)
                    ):
                        raise MigrationError(
                            "A tabela SQLite pesagens não contém as colunas "
                            "de identificação esperadas."
                        )
                elif not obrigatorias <= set(colunas):
                    raise MigrationError(
                        f"A tabela SQLite {tabela} não contém as colunas esperadas."
                    )
                dados[tabela] = [
                    dict(zip(colunas, linha))
                    for linha in cursor.fetchall()
                ]

            if not dados["pesagens"]:
                raise MigrationError("O SQLite não contém fichas para importar.")

            try:
                for linha in dados["pesagens"]:
                    json.dumps(linha, ensure_ascii=False, allow_nan=False)
            except (TypeError, ValueError):
                raise MigrationError(
                    "Há uma ficha SQLite com valores que não podem ser "
                    "representados como JSON."
                ) from None

            usuarios = {linha["usuario"] for linha in dados["usuarios"]}
            if any(
                not isinstance(usuario, str) or not usuario.strip()
                for usuario in usuarios
            ):
                raise MigrationError(
                    "O SQLite contém um usuário sem identificador válido."
                )
            nomes_normalizados = [usuario.casefold() for usuario in usuarios]
            if len(nomes_normalizados) != len(set(nomes_normalizados)):
                raise MigrationError(
                    "O SQLite contém usuários que diferem apenas por "
                    "maiúsculas e minúsculas."
                )
            assinaturas_sem_usuario = {
                linha["usuario"]
                for linha in dados["assinaturas_usuarios"]
                if linha["usuario"] not in usuarios
            }
            if assinaturas_sem_usuario:
                raise MigrationError(
                    "Há assinaturas sem usuário correspondente no SQLite."
                )

            return dados
    except sqlite3.Error as exc:
        raise MigrationError(
            f"Não foi possível ler o SQLite ({type(exc).__name__})."
        ) from None


def url_postgresql():
    try:
        from sqlalchemy.engine import make_url
        from sqlalchemy.exc import ArgumentError
    except ImportError:
        raise MigrationError(
            "Instale as dependências do projeto com: pip install -r requirements.txt"
        ) from None

    valor = os.environ.get("SUPABASE_DATABASE_URL", "").strip()
    if not valor:
        raise MigrationError(
            "Defina SUPABASE_DATABASE_URL no ambiente usando a credencial "
            "atual do Supabase; não a coloque no código nem no chat."
        )

    try:
        url = make_url(valor)
    except (ArgumentError, ValueError):
        raise MigrationError("SUPABASE_DATABASE_URL não é uma URL válida.") from None

    if url.drivername == "postgresql":
        url = url.set(drivername="postgresql+psycopg")
    if url.drivername != "postgresql+psycopg":
        raise MigrationError(
            "SUPABASE_DATABASE_URL deve usar PostgreSQL com o driver psycopg."
        )
    return url


def verificar_tabelas_destino(conn):
    from sqlalchemy import text

    for tabela in TABLES:
        existe = conn.execute(
            text("SELECT to_regclass(:nome)"),
            {"nome": f"public.{tabela}"},
        ).scalar_one()
        if existe is None:
            raise MigrationError(
                f"A tabela public.{tabela} não existe. Inicialize o app primeiro."
            )


def verificar_conflitos_assinaturas(conn, dados):
    from sqlalchemy import text

    for assinatura in dados["assinaturas_usuarios"]:
        usuario = assinatura["usuario"]
        existentes = conn.execute(
            text(
                "SELECT usuario FROM public.usuarios "
                "WHERE lower(usuario) = lower(:usuario)"
            ),
            {"usuario": usuario},
        ).scalars().all()
        if existentes and usuario not in existentes:
            raise MigrationError(
                "Há um usuário no Supabase que difere apenas por "
                "maiúsculas/minúsculas do dono de uma assinatura SQLite; "
                "resolva esse conflito antes da importação."
            )


def importar(conn, dados):
    from sqlalchemy import text

    total_fichas = conn.execute(
        text("SELECT COUNT(*) FROM public.pesagens")
    ).scalar_one()
    if total_fichas:
        raise MigrationError(
            "O Supabase já contém fichas; a importação foi cancelada para "
            "evitar duplicação ou mistura de históricos."
        )

    for tabela in ("usuarios", "pesagens", *DEPENDENT_TABLES):
        linhas = dados[tabela]
        if not linhas:
            continue

        colunas = list(linhas[0])
        if any(set(linha) != set(colunas) for linha in linhas):
            raise MigrationError(
                f"As linhas da tabela {tabela} têm formatos diferentes."
            )

        if tabela == "pesagens":
            consulta = text(
                "INSERT INTO public.pesagens (dados) "
                "VALUES (CAST(:dados AS JSONB))"
            )
            parametros = [
                {
                    "dados": json.dumps(
                        linha, ensure_ascii=False, allow_nan=False
                    )
                }
                for linha in linhas
            ]
        else:
            nomes_colunas = ", ".join(
                f'"{coluna.replace(chr(34), chr(34) * 2)}"'
                for coluna in colunas
            )
            nomes_parametros = ", ".join(f":{coluna}" for coluna in colunas)
            consulta = text(
                f"INSERT INTO public.{tabela} ({nomes_colunas}) "
                f"VALUES ({nomes_parametros}) ON CONFLICT DO NOTHING"
            )
            parametros = linhas

        conn.execute(consulta, parametros)

    total_importado = conn.execute(
        text("SELECT COUNT(*) FROM public.pesagens")
    ).scalar_one()
    esperado = len(dados["pesagens"])
    if total_importado != esperado:
        raise MigrationError(
            f"Verificação pós-importação falhou: esperado {esperado} "
            f"fichas, encontrado {total_importado}."
        )
    return total_importado


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Valida ou importa o SQLite para tabelas PostgreSQL vazias, "
            "sem substituir contas e metadados existentes."
        )
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=Path("aeronaves.db"),
        help="Caminho do SQLite de origem (padrão: aeronaves.db).",
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Executa a importação; sem esta opção apenas valida a conexão.",
    )
    argumentos = parser.parse_args()

    try:
        if not argumentos.source.is_file():
            raise MigrationError(
                f"Arquivo SQLite não encontrado: {argumentos.source}"
            )
        dados = ler_banco_sqlite(argumentos.source)
        url = url_postgresql()
        try:
            from sqlalchemy import create_engine, text
            from sqlalchemy.exc import SQLAlchemyError
        except ImportError:
            raise MigrationError(
                "Instale as dependências do projeto com: "
                "pip install -r requirements.txt"
            ) from None

        engine = None
        try:
            engine = create_engine(url, pool_pre_ping=True)
            with engine.connect() as conn:
                verificar_tabelas_destino(conn)
                quantidade_destino = conn.execute(
                    text("SELECT COUNT(*) FROM public.pesagens")
                ).scalar_one()
                print(
                    "Conexão PostgreSQL validada; "
                    f"SQLite íntegro com {len(dados['pesagens'])} fichas; "
                    f"Supabase contém {quantidade_destino} fichas."
                )
                if not argumentos.apply:
                    print("Modo de validação: nenhuma alteração foi feita.")
                    return 0

            with engine.begin() as conn:
                verificar_tabelas_destino(conn)
                conn.execute(
                    text(
                        "LOCK TABLE public.pesagens "
                        "IN SHARE ROW EXCLUSIVE MODE"
                    )
                )
                verificar_conflitos_assinaturas(conn, dados)
                total = importar(conn, dados)
            print(f"Importação concluída e verificada: {total} fichas.")
            return 0
        except SQLAlchemyError as exc:
            raise MigrationError(
                "Falha PostgreSQL durante conexão ou importação "
                f"({type(exc).__name__}); nenhuma senha foi exibida."
            ) from None
        finally:
            if engine is not None:
                engine.dispose()
    except MigrationError as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
