"""
database.py — Gerenciamento da base local de EAN via SQLite.
"""

import csv
import hashlib
import os
import sqlite3
from datetime import date, datetime
from typing import Dict, Optional, Tuple

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "ean_database.db")
CAMEX_CSV_PATH = os.path.join(BASE_DIR, "data", "camex_ncm_database.csv")
TTD409_CSV_PATH = os.path.join(BASE_DIR, "data", "ttd409_exclusions.csv")


def _conn() -> sqlite3.Connection:
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db() -> None:
    """Cria a tabela de EANs se não existir."""
    with _conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS ean_base (
                sku         TEXT PRIMARY KEY,
                ean         TEXT NOT NULL,
                descricao   TEXT DEFAULT '',
                criado_em   TEXT,
                atualizado_em TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS app_meta (
                chave TEXT PRIMARY KEY,
                valor TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS camex_ncm_base (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ncm TEXT NOT NULL,
                ncm_formatado TEXT,
                match_tipo TEXT DEFAULT 'exato',
                prefix_len INTEGER DEFAULT 8,
                ex TEXT DEFAULT '',
                lista TEXT,
                categoria TEXT,
                descricao TEXT,
                aliquota TEXT,
                quota TEXT,
                unidade_quota TEXT,
                inicio_vigencia TEXT,
                fim_vigencia TEXT,
                ato_legal TEXT,
                portaria_secex TEXT,
                modelo_lpco TEXT,
                observacao TEXT,
                fonte_arquivo TEXT,
                fonte_url TEXT,
                fonte_atualizacao TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS ttd409_exclusoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item TEXT NOT NULL,
                ncm TEXT NOT NULL,
                ncm_formatado TEXT,
                match_tipo TEXT DEFAULT 'prefixo',
                prefix_len INTEGER DEFAULT 4,
                descricao_legal TEXT,
                observacao TEXT,
                fonte_url TEXT,
                fonte_atualizacao TEXT
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_camex_ncm ON camex_ncm_base (ncm)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_camex_lista ON camex_ncm_base (lista)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_ttd409_ncm ON ttd409_exclusoes (ncm)")
        _seed_camex_from_csv(conn)
        _seed_ttd409_from_csv(conn)
        conn.commit()


def _csv_sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _seed_camex_from_csv(conn: sqlite3.Connection) -> None:
    if not os.path.exists(CAMEX_CSV_PATH):
        return

    digest = _csv_sha256(CAMEX_CSV_PATH)
    row = conn.execute(
        "SELECT valor FROM app_meta WHERE chave = 'camex_csv_sha256'"
    ).fetchone()
    total = conn.execute("SELECT COUNT(*) FROM camex_ncm_base").fetchone()[0]
    if row and row[0] == digest and total > 0:
        return

    conn.execute("DELETE FROM camex_ncm_base")
    with open(CAMEX_CSV_PATH, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        registros = [
            (
                r.get("ncm", ""),
                r.get("ncm_formatado", ""),
                r.get("match_tipo", "exato"),
                int(r.get("prefix_len") or 8),
                r.get("ex", ""),
                r.get("lista", ""),
                r.get("categoria", ""),
                r.get("descricao", ""),
                r.get("aliquota", ""),
                r.get("quota", ""),
                r.get("unidade_quota", ""),
                r.get("inicio_vigencia", ""),
                r.get("fim_vigencia", ""),
                r.get("ato_legal", ""),
                r.get("portaria_secex", ""),
                r.get("modelo_lpco", ""),
                r.get("observacao", ""),
                r.get("fonte_arquivo", ""),
                r.get("fonte_url", ""),
                r.get("fonte_atualizacao", ""),
            )
            for r in reader
        ]

    conn.executemany(
        """
        INSERT INTO camex_ncm_base (
            ncm, ncm_formatado, match_tipo, prefix_len, ex, lista, categoria,
            descricao, aliquota, quota, unidade_quota, inicio_vigencia,
            fim_vigencia, ato_legal, portaria_secex, modelo_lpco, observacao,
            fonte_arquivo, fonte_url, fonte_atualizacao
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,
        registros,
    )
    conn.execute(
        "INSERT OR REPLACE INTO app_meta (chave, valor) VALUES ('camex_csv_sha256', ?)",
        (digest,),
    )


def _seed_ttd409_from_csv(conn: sqlite3.Connection) -> None:
    if not os.path.exists(TTD409_CSV_PATH):
        return

    digest = _csv_sha256(TTD409_CSV_PATH)
    row = conn.execute(
        "SELECT valor FROM app_meta WHERE chave = 'ttd409_csv_sha256'"
    ).fetchone()
    total = conn.execute("SELECT COUNT(*) FROM ttd409_exclusoes").fetchone()[0]
    if row and row[0] == digest and total > 0:
        return

    conn.execute("DELETE FROM ttd409_exclusoes")
    with open(TTD409_CSV_PATH, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        registros = [
            (
                r.get("item", ""),
                r.get("ncm", ""),
                r.get("ncm_formatado", ""),
                r.get("match_tipo", "prefixo"),
                int(r.get("prefix_len") or len(r.get("ncm", "")) or 4),
                r.get("descricao_legal", ""),
                r.get("observacao", ""),
                r.get("fonte_url", ""),
                r.get("fonte_atualizacao", ""),
            )
            for r in reader
        ]

    conn.executemany(
        """
        INSERT INTO ttd409_exclusoes (
            item, ncm, ncm_formatado, match_tipo, prefix_len,
            descricao_legal, observacao, fonte_url, fonte_atualizacao
        ) VALUES (?,?,?,?,?,?,?,?,?)
        """,
        registros,
    )
    conn.execute(
        "INSERT OR REPLACE INTO app_meta (chave, valor) VALUES ('ttd409_csv_sha256', ?)",
        (digest,),
    )


def _limpar_digitos_db(s: str) -> str:
    return "".join(ch for ch in str(s or "") if ch.isdigit())


def _normalizar_ncm(ncm: str) -> str:
    digits = _limpar_digitos_db(ncm)
    return digits[:8] if len(digits) >= 8 else digits


def _normalizar_ex(ex: str) -> str:
    digits = _limpar_digitos_db(ex)
    if not digits:
        return ""
    return digits[-3:].zfill(3)


def _vigente(registro: Dict, hoje: Optional[str] = None) -> bool:
    hoje = hoje or date.today().isoformat()
    inicio = (registro.get("inicio_vigencia") or "")[:10]
    fim = (registro.get("fim_vigencia") or "")[:10]
    if inicio and inicio > hoje:
        return False
    if fim and fim < hoje:
        return False
    return True


def buscar_ean(sku: str) -> Optional[str]:
    """Retorna o EAN para um SKU ou None se não encontrado."""
    with _conn() as conn:
        row = conn.execute(
            "SELECT ean FROM ean_base WHERE sku = ?", (sku,)
        ).fetchone()
    return row[0] if row else None


def get_all_eans() -> Dict[str, str]:
    """Retorna dicionário {sku: ean} com toda a base."""
    with _conn() as conn:
        rows = conn.execute("SELECT sku, ean FROM ean_base").fetchall()
    return {r[0]: r[1] for r in rows}


def upsert_eans(registros: Dict[str, Dict]) -> Tuple[int, int, int]:
    """
    Insere ou atualiza registros em lote.
    registros: {sku: {"ean": "...", "descricao": "..."}}
    Retorna (inseridos, atualizados, erros).
    """
    inseridos = atualizados = erros = 0
    now = datetime.now().isoformat()
    with _conn() as conn:
        for sku, dados in registros.items():
            try:
                ean = str(dados.get("ean", "")).strip()
                desc = str(dados.get("descricao", "")).strip()
                if not sku or not ean:
                    erros += 1
                    continue
                existente = conn.execute(
                    "SELECT sku FROM ean_base WHERE sku = ?", (sku,)
                ).fetchone()
                if existente:
                    conn.execute(
                        "UPDATE ean_base SET ean=?, descricao=?, atualizado_em=? WHERE sku=?",
                        (ean, desc, now, sku),
                    )
                    atualizados += 1
                else:
                    conn.execute(
                        "INSERT INTO ean_base (sku, ean, descricao, criado_em, atualizado_em) VALUES (?,?,?,?,?)",
                        (sku, ean, desc, now, now),
                    )
                    inseridos += 1
            except Exception:
                erros += 1
        conn.commit()
    return inseridos, atualizados, erros


def salvar_ean_manual(sku: str, ean: str, descricao: str = "") -> None:
    """Salva um único EAN manualmente (upsert)."""
    now = datetime.now().isoformat()
    with _conn() as conn:
        existente = conn.execute(
            "SELECT criado_em FROM ean_base WHERE sku = ?", (sku,)
        ).fetchone()
        criado = existente[0] if existente else now
        conn.execute(
            "INSERT OR REPLACE INTO ean_base (sku, ean, descricao, criado_em, atualizado_em) VALUES (?,?,?,?,?)",
            (sku, ean, descricao, criado, now),
        )
        conn.commit()


def get_db_stats() -> Dict:
    """Retorna estatísticas da base."""
    with _conn() as conn:
        total = conn.execute("SELECT COUNT(*) FROM ean_base").fetchone()[0]
        ultima = conn.execute(
            "SELECT MAX(atualizado_em) FROM ean_base"
        ).fetchone()[0]
    return {"total": total, "ultima_atualizacao": ultima or "—"}


def listar_todos() -> list:
    """Retorna lista de todos os registros para exibição."""
    with _conn() as conn:
        rows = conn.execute(
            "SELECT sku, ean, descricao, atualizado_em FROM ean_base ORDER BY atualizado_em DESC"
        ).fetchall()
    return [
        {"SKU": r[0], "EAN": r[1], "Descrição": r[2], "Atualizado em": r[3]}
        for r in rows
    ]


def buscar_camex_por_ncm(ncm: str, extipi: str = "", somente_vigentes: bool = True) -> list:
    """Retorna registros CAMEX/Gecex aplicaveis ao NCM informado."""
    ncm_norm = _normalizar_ncm(ncm)
    if len(ncm_norm) != 8:
        return []

    with _conn() as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT * FROM camex_ncm_base
            WHERE ncm = ?
               OR (match_tipo = 'prefixo' AND ? LIKE ncm || '%')
            """,
            (ncm_norm, ncm_norm),
        ).fetchall()

    registros = [dict(r) for r in rows]
    if somente_vigentes:
        hoje = date.today().isoformat()
        registros = [r for r in registros if _vigente(r, hoje)]

    ex_norm = _normalizar_ex(extipi)

    def sort_key(r: Dict) -> tuple:
        ex = r.get("ex") or ""
        if ex_norm:
            ex_rank = 0 if ex == ex_norm else (1 if not ex else 2)
        else:
            ex_rank = 0 if not ex else 1
        match_rank = 0 if r.get("match_tipo") == "exato" else 1
        return (match_rank, ex_rank, r.get("lista") or "", ex)

    return sorted(registros, key=sort_key)


def get_camex_stats() -> Dict:
    """Retorna estatisticas da base CAMEX local."""
    with _conn() as conn:
        conn.row_factory = sqlite3.Row
        rows = [dict(r) for r in conn.execute("SELECT * FROM camex_ncm_base").fetchall()]
    ativos = [r for r in rows if _vigente(r)]
    return {
        "total_registros": len(rows),
        "registros_vigentes": len(ativos),
        "ncm_vigentes": len({r["ncm"] for r in ativos if r.get("match_tipo") == "exato"}),
        "prefixos_vigentes": len({r["ncm"] for r in ativos if r.get("match_tipo") == "prefixo"}),
        "fontes": len({r.get("fonte_arquivo") for r in rows if r.get("fonte_arquivo")}),
        "ultima_atualizacao": max((r.get("fonte_atualizacao") or "" for r in rows), default=""),
    }


def buscar_ttd409_por_ncm(ncm: str) -> list:
    """Retorna mercadorias do Decreto SC 2.128/2009 que batem com o NCM."""
    ncm_norm = _normalizar_ncm(ncm)
    if len(ncm_norm) != 8:
        return []

    with _conn() as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT * FROM ttd409_exclusoes
            WHERE ? LIKE ncm || '%'
            ORDER BY prefix_len DESC, item
            """,
            (ncm_norm,),
        ).fetchall()
    return [dict(r) for r in rows]


def get_ttd409_stats() -> Dict:
    """Retorna estatisticas da base de exclusoes TTD409."""
    with _conn() as conn:
        total = conn.execute("SELECT COUNT(*) FROM ttd409_exclusoes").fetchone()[0]
        itens = conn.execute("SELECT COUNT(DISTINCT item) FROM ttd409_exclusoes").fetchone()[0]
        ultima = conn.execute("SELECT MAX(fonte_atualizacao) FROM ttd409_exclusoes").fetchone()[0]
    return {
        "total_registros": total,
        "itens_legais": itens,
        "ultima_atualizacao": ultima or "",
    }
