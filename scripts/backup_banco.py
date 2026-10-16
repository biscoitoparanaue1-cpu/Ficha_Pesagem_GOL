"""Exporta todas as tabelas do banco para arquivos JSON, um por tabela.
Usado pelo workflow backup-banco.yml.

Uso: python scripts/backup_banco.py <pasta de saída>
A conexão vem da variável SUPABASE_DB_URL.
"""
import base64
import datetime
import decimal
import json
import os
import pathlib
import sys

import psycopg

TABELAS = [
    "pesagens",
    "fluxo_fichas",
    "campos_com_erro",
    "historico_fichas",
    "aeronaves",
    "opcoes_cadastradas",
    "usuarios",
    "assinaturas_usuarios",
]


def para_json(valor):
    if isinstance(valor, (bytes, bytearray, memoryview)):
        return {"base64": base64.b64encode(bytes(valor)).decode("ascii")}
    if isinstance(valor, (datetime.date, datetime.datetime, datetime.time)):
        return valor.isoformat()
    if isinstance(valor, decimal.Decimal):
        return float(valor)
    return valor


def main(pasta):
    url = os.environ["SUPABASE_DB_URL"].replace("postgresql+psycopg://", "postgresql://")
    saida = pathlib.Path(pasta)
    saida.mkdir(parents=True, exist_ok=True)
    with psycopg.connect(url, prepare_threshold=None, connect_timeout=30) as conn:
        for tabela in TABELAS:
            existe = conn.execute("SELECT to_regclass(%s)", (f"public.{tabela}",)).fetchone()[0]
            if not existe:
                print(f"{tabela}: não existe, pulando")
                continue
            cursor = conn.execute(f'SELECT * FROM "{tabela}"')
            colunas = [coluna.name for coluna in cursor.description]
            linhas = [
                {coluna: para_json(valor) for coluna, valor in zip(colunas, linha)}
                for linha in cursor
            ]
            (saida / f"{tabela}.json").write_text(
                json.dumps(linhas, ensure_ascii=False, default=str),
                encoding="utf-8",
            )
            print(f"{tabela}: {len(linhas)} linhas")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "backup"))
