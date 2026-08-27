"""Backup simples do banco (todas as tabelas em um unico JSON).

Uso:
    python backup.py

Le a mesma DATABASE_URL do .env / variavel de ambiente usada pelo app,
e salva um arquivo em backups/backup_AAAAMMDD_HHMMSS.json com todas as
linhas de todas as tabelas. Nao precisa de pg_dump instalado.
"""
import json
import os
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
TABELAS = [
    "config", "produtos", "clientes", "vendas", "itens_venda",
    "caixa", "contas_receber", "contas_pagar",
]


def _carregar_env_local():
    caminho = BASE_DIR / ".env"
    if not caminho.exists():
        return
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, valor = linha.split("=", 1)
        os.environ.setdefault(chave.strip(), valor.strip().strip('"').strip("'"))


def main():
    _carregar_env_local()
    import psycopg
    from psycopg.rows import dict_row

    url = os.environ.get("DATABASE_URL")
    if not url:
        raise SystemExit("DATABASE_URL nao configurada (crie webapp/.env).")

    pasta = BASE_DIR / "backups"
    pasta.mkdir(exist_ok=True)
    nome = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    destino = pasta / nome

    dump = {}
    with psycopg.connect(url) as conn:
        for tabela in TABELAS:
            with conn.cursor(row_factory=dict_row) as cur:
                cur.execute(f"SELECT * FROM {tabela} ORDER BY 1")
                linhas = cur.fetchall()
                dump[tabela] = [dict(r) for r in linhas]

    with open(destino, "w", encoding="utf-8") as f:
        json.dump(dump, f, ensure_ascii=False, indent=2, default=str)

    total = sum(len(v) for v in dump.values())
    print(f"Backup salvo em: {destino}")
    for tabela, linhas in dump.items():
        print(f"  {tabela}: {len(linhas)} linha(s)")
    print(f"Total: {total} linha(s)")


if __name__ == "__main__":
    main()
