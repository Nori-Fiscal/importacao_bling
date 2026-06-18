"""
products_db.py — Banco de produtos para padronização de descrição e conferência de NCM.
"""
import os, sqlite3
from datetime import datetime
from typing import Dict, List, Optional, Tuple

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "produtos_base.db")

def _conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db_produtos():
    with _conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS produtos_base (
                sku_al TEXT PRIMARY KEY,
                sku_tonina TEXT NOT NULL DEFAULT '',
                descricao_padrao TEXT NOT NULL DEFAULT '',
                unidade TEXT NOT NULL DEFAULT '',
                ncm_al TEXT NOT NULL DEFAULT '',
                ncm_tonina TEXT NOT NULL DEFAULT '',
                marca TEXT NOT NULL DEFAULT 'LORBEN',
                criado_em TEXT,
                atualizado_em TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS divergencias_ncm (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sku_al TEXT NOT NULL,
                sku_tonina TEXT NOT NULL DEFAULT '',
                ncm_duimp TEXT NOT NULL DEFAULT '',
                ncm_tonina TEXT NOT NULL DEFAULT '',
                ncm_padrao_al TEXT NOT NULL DEFAULT '',
                arquivo_xml TEXT NOT NULL DEFAULT '',
                processado_em TEXT,
                resolvido INTEGER DEFAULT 0
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_div_sku ON divergencias_ncm(sku_al)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_div_res ON divergencias_ncm(resolvido)")
        conn.commit()

def _sku_ton(sku_al: str) -> str:
    s = sku_al.strip().upper()
    if s.startswith("AL "): return s[3:]
    if s.startswith("AL"): return s[2:]
    return s

def _ncm8(ncm: str) -> str:
    d = "".join(ch for ch in str(ncm or "") if ch.isdigit())
    return d[:8] if len(d) >= 8 else d

# --- IMPORT produtos_2026 ---
def importar_produtos_2026(rows: List[Dict]) -> Tuple[int, int, int]:
    ins = upd = err = 0
    now = datetime.now().isoformat()
    with _conn() as conn:
        for row in rows:
            try:
                sku_al = str(row.get("sku","")).strip().upper()
                if not sku_al: err += 1; continue
                desc = str(row.get("descricao","")).strip()
                und = str(row.get("unidade","")).strip()
                ncm = str(row.get("ncm","")).strip()
                sku_t = _sku_ton(sku_al)
                nome = desc
                for sep in [" - CÓD: ", " - CÓD.: ", " - COD: "]:
                    if sep in desc: nome = desc.split(sep)[0].strip(); break
                if not nome: nome = desc
                padrao = f"{nome} - CÓD: {sku_al} - MARCA: LORBEN"
                ex = conn.execute("SELECT 1 FROM produtos_base WHERE sku_al=?", (sku_al,)).fetchone()
                if ex:
                    conn.execute("UPDATE produtos_base SET sku_tonina=?, descricao_padrao=?, unidade=?, ncm_al=?, marca=?, atualizado_em=? WHERE sku_al=?", (sku_t, padrao, und, ncm, "LORBEN", now, sku_al))
                    upd += 1
                else:
                    conn.execute("INSERT INTO produtos_base (sku_al,sku_tonina,descricao_padrao,unidade,ncm_al,marca,criado_em,atualizado_em) VALUES (?,?,?,?,?,?,?,?)", (sku_al, sku_t, padrao, und, ncm, "LORBEN", now, now))
                    ins += 1
            except: err += 1
        conn.commit()
    return ins, upd, err

# --- IMPORT NCM Tonina ---
def importar_relprodutos01(rows: List[Dict]) -> Tuple[int, int, int]:
    atu = sem = err = 0
    now = datetime.now().isoformat()
    with _conn() as conn:
        for row in rows:
            try:
                cod = str(row.get("codigo","")).strip().upper()
                if not cod: err += 1; continue
                ncm = _ncm8(str(row.get("ncm","")).strip())
                sku_al = "AL" + cod
                ex = conn.execute("SELECT 1 FROM produtos_base WHERE sku_al=?", (sku_al,)).fetchone()
                if ex:
                    conn.execute("UPDATE produtos_base SET ncm_tonina=?, atualizado_em=? WHERE sku_al=?", (ncm, now, sku_al))
                    atu += 1
                else:
                    m = conn.execute("SELECT sku_al FROM produtos_base WHERE sku_tonina=? AND sku_tonina!=''", (cod,)).fetchone()
                    if m:
                        conn.execute("UPDATE produtos_base SET ncm_tonina=?, atualizado_em=? WHERE sku_al=?", (ncm, now, m[0]))
                        atu += 1
                    else: sem += 1
            except: err += 1
        conn.commit()
    return atu, sem, err

# --- CONSULTAS ---
def buscar_produto_por_sku(sku_al: str) -> Optional[Dict]:
    with _conn() as conn:
        conn.row_factory = sqlite3.Row
        r = conn.execute("SELECT * FROM produtos_base WHERE sku_al=?", (sku_al.strip().upper(),)).fetchone()
    return dict(r) if r else None

def conferir_ncm(sku_al: str, ncm_duimp: str) -> Dict:
    ncm_norm = _ncm8(ncm_duimp)
    r = {"sku_al": sku_al.strip().upper(), "sku_tonina":"", "ncm_duimp":ncm_norm, "ncm_padrao_al":"", "ncm_tonina":"", "ok":True, "sem_base":False, "alertas":[]}
    with _conn() as conn:
        row = conn.execute("SELECT ncm_al,ncm_tonina,sku_tonina FROM produtos_base WHERE sku_al=?", (sku_al.strip().upper(),)).fetchone()
    if not row:
        r["alertas"].append(f"SKU {sku_al} não encontrado na base.")
        r["ok"] = False
        r["sem_base"] = True
        return r
    ncm_al, ncm_ton, sku_ton = row[0] or "", row[1] or "", row[2] or ""
    r["ncm_padrao_al"] = ncm_al; r["ncm_tonina"] = ncm_ton; r["sku_tonina"] = sku_ton
    if ncm_al and _ncm8(ncm_al) != ncm_norm:
        r["alertas"].append(f"NCM do XML ({ncm_norm}) difere do padrão Albema ({ncm_al})."); r["ok"]=False
    if ncm_ton and _ncm8(ncm_ton) != ncm_norm:
        r["alertas"].append(f"NCM do XML ({ncm_norm}) difere do cadastro TONINA ({ncm_ton}, SKU Tonina: {sku_ton}). Altere na Tonina se a DUIMP estiver correta."); r["ok"]=False
    return r

def registrar_divergencia_ncm(sku_al:str, sku_tonina:str, ncm_duimp:str, ncm_tonina:str, ncm_padrao_al:str, arquivo_xml:str=""):
    now = datetime.now().isoformat()
    with _conn() as conn:
        conn.execute("INSERT INTO divergencias_ncm (sku_al,sku_tonina,ncm_duimp,ncm_tonina,ncm_padrao_al,arquivo_xml,processado_em) VALUES (?,?,?,?,?,?,?)", (sku_al,sku_tonina,ncm_duimp,ncm_tonina,ncm_padrao_al,arquivo_xml,now))
        conn.commit()

def listar_divergencias(apenas_pendentes=True):
    with _conn() as conn:
        conn.row_factory = sqlite3.Row
        q = "SELECT * FROM divergencias_ncm WHERE resolvido=0 ORDER BY processado_em DESC" if apenas_pendentes else "SELECT * FROM divergencias_ncm ORDER BY processado_em DESC LIMIT 200"
        return [dict(r) for r in conn.execute(q).fetchall()]

def marcar_divergencia_resolvida(id_div:int):
    with _conn() as conn:
        conn.execute("UPDATE divergencias_ncm SET resolvido=1 WHERE id=?", (id_div,)); conn.commit()

def listar_produtos_base():
    with _conn() as conn:
        conn.row_factory = sqlite3.Row
        return [dict(r) for r in conn.execute("SELECT sku_al,sku_tonina,descricao_padrao,unidade,ncm_al,ncm_tonina,marca FROM produtos_base ORDER BY sku_al").fetchall()]

def get_produtos_stats():
    with _conn() as conn:
        t = conn.execute("SELECT COUNT(*) FROM produtos_base").fetchone()[0]
        cd = conn.execute("SELECT COUNT(*) FROM produtos_base WHERE descricao_padrao!=''").fetchone()[0]
        na = conn.execute("SELECT COUNT(*) FROM produtos_base WHERE ncm_al!=''").fetchone()[0]
        nt = conn.execute("SELECT COUNT(*) FROM produtos_base WHERE ncm_tonina!=''").fetchone()[0]
        dp = conn.execute("SELECT COUNT(*) FROM divergencias_ncm WHERE resolvido=0").fetchone()[0]
    return {"total":t,"com_descricao_padrao":cd,"com_ncm_al":na,"com_ncm_tonina":nt,"divergencias_pendentes":dp}
