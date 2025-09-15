# app/app.py
# -*- coding: utf-8 -*-
# Cascata (Complexo > Unidade > Família > TAG > Programa) + Checklist (Etapa 1/2)
# Etapa 2 dinâmica por Categoria (Sim/Não/Não Aplica) com salvamento em QUESTOES.
# Calcula Índice de Vulnerabilidade (controle e ativo) e salva em:
# INDICES_CONTROLE e INDICE_ATIVO, além das abas já existentes.

from pathlib import Path
import re
import unicodedata
import hashlib
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(
    page_title="Diagnóstico - Checklist & Inspeção",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUT_FILE = DATA_DIR / "Inspecoes.xlsx"

ASSET_PREFERRED = [DATA_DIR / "Lista de Ativos.xlsx"]
DIAG_PREFERRED  = [DATA_DIR / "Diagnostico.xlsx", DATA_DIR / "Diagnóstico.xlsx"]
QUEST_PREFERRED = [DATA_DIR / "Perguntas - Categoria.xlsx", DATA_DIR / "Perguntas_Categoria.xlsx"]

# ------------------ helpers ------------------
def _strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s))
    return s.encode("ascii", "ignore").decode("ascii")

def _canon(s: str) -> str:
    s = _strip_accents(str(s)).lower()
    return re.sub(r"\s+", " ", s).strip()

def _safe_key(texto: str) -> str:
    s = _strip_accents(texto or "").lower()
    return "".join(ch if ch.isalnum() else "_" for ch in s)

def _mk_widget_key(*parts) -> str:
    raw = "|".join(str(p) for p in parts)
    base = "_".join((_safe_key(str(p))[:16] if not isinstance(p, int) else f"{p:03d}") for p in parts)
    return f"k_{base}_{hashlib.md5(raw.encode()).hexdigest()[:8]}"

def _find_file(preferred, pattern) -> Path | None:
    for p in preferred:
        if p.exists(): return p
    cands = list(DATA_DIR.glob(pattern))
    return cands[0] if cands else None

def _to_float(x, default=0.0):
    try:
        v = pd.to_numeric(x, errors="coerce")
        return float(v) if pd.notna(v) else float(default)
    except Exception:
        return float(default)

def _cols_top(spec, gap="small"):
    try:
        return st.columns(spec, gap=gap, vertical_alignment="top")
    except TypeError:
        return st.columns(spec, gap=gap)

# ------------------ renomes ------------------
def _rename_assets_cols(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for col in df.columns:
        cc = _canon(col)
        if "complexo" in cc: mapping[col] = "Complexo"
        elif "unidade" in cc: mapping[col] = "Unidade"
        elif "familia" in cc: mapping[col] = "Família"
        elif "ativo" in cc:   mapping[col] = "Ativos"
    return df.rename(columns=mapping)

def _rename_items_cols(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for col in df.columns:
        cc = _canon(col)
        if "programa" in cc: mapping[col] = "Programa"
        elif "familia" in cc: mapping[col] = "Família"
        elif ("cenario" in cc) and ("risco" in cc): mapping[col] = "Cenário de Risco"
        elif "verificar" in cc: mapping[col] = "Verificar"
        elif "componente" in cc: mapping[col] = "Componente"
        elif ("tipo" in cc) and ("controle" in cc): mapping[col] = "Tipo de Controle"
        elif "categoria" in cc: mapping[col] = "Categoria"
        elif "ameaca" in cc or "ameaça" in cc: mapping[col] = "Ameaça"
        elif "hierarquia" in cc and "controle" in cc: mapping[col] = "Hierarquia (Controle)"
        elif "hierarquia" in cc and "componente" in cc: mapping[col] = "Hierarquia (Componente)"
        elif cc == "hierarquia": mapping[col] = "Hierarquia"
        elif "peso" in cc: mapping[col] = "Peso"
    df = df.rename(columns=mapping)
    hier_cols = [c for c in df.columns if re.fullmatch(r"Hierarquia(\.\d+)?", str(c))]
    if len(hier_cols) == 1 and "Hierarquia (Controle)" not in df.columns and "Hierarquia (Componente)" not in df.columns:
        df = df.rename(columns={hier_cols[0]: "Hierarquia (Controle)"})
    elif len(hier_cols) >= 2:
        df = df.rename(columns={hier_cols[0]: "Hierarquia (Controle)", hier_cols[1]: "Hierarquia (Componente)"})
        for j, c in enumerate(hier_cols[2:], start=3):
            df = df.rename(columns={c: f"Hierarquia (Extra {j})"})
    return df

def _rename_questions_cols(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for col in df.columns:
        cc = _canon(col)
        if "categoria" in cc: mapping[col] = "Categoria"
        elif "pergunta" in cc: mapping[col] = "Pergunta"
        elif "peso" in cc: mapping[col] = "Peso_q"
    return df.rename(columns=mapping)

# ------------------ loaders ------------------
@st.cache_data(show_spinner=False)
def load_assets() -> pd.DataFrame | None:
    f = _find_file(ASSET_PREFERRED, "*Ativos*.xls*")
    if not f: return None
    xls = pd.ExcelFile(f)
    sheet = "D_Equipamentos" if "D_Equipamentos" in xls.sheet_names else xls.sheet_names[0]
    df = xls.parse(sheet)
    df = _rename_assets_cols(df)
    if not {"Família","Ativos"}.issubset(df.columns): return None
    keep = [c for c in ["Complexo","Unidade","Família","Ativos"] if c in df.columns]
    df = df[keep].dropna(how="all").copy().rename(columns={"Ativos":"Tag"})
    for c in df.columns: df[c] = df[c].astype(str).str.strip()
    if "Tag" in df.columns: df["Tag"] = df["Tag"].str.upper()
    for miss in ["Complexo","Unidade"]:
        if miss not in df.columns: df[miss] = "-"
    df = df.drop_duplicates()
    df["Familia_norm"] = df["Família"].apply(_canon)
    return df[["Complexo","Unidade","Família","Familia_norm","Tag"]]

@st.cache_data(show_spinner=False)
def load_items_cr() -> pd.DataFrame | None:
    f = _find_file(DIAG_PREFERRED, "Diagnost*.xls*")
    if not f: return None
    df = pd.read_excel(f, sheet_name=0)
    df = _rename_items_cols(df)

    # Converte numéricos
    for numcol in ["Peso", "Hierarquia (Controle)", "Hierarquia (Componente)"]:
        if numcol in df.columns:
            df[numcol] = (
                df[numcol].astype(str).str.replace(",", ".", regex=False)
                .apply(lambda x: pd.to_numeric(x, errors="coerce"))
            )

    # Fallback: se "Hierarquia (Componente)" estiver vazio/zero, usa "Peso"
    if "Hierarquia (Componente)" in df.columns:
        hc_series = pd.to_numeric(df["Hierarquia (Componente)"], errors="coerce")
        if hc_series.isna().all() or (hc_series.fillna(0) == 0).all():
            if "Peso" in df.columns:
                df["Hierarquia (Componente)"] = pd.to_numeric(df["Peso"], errors="coerce")
    elif "Peso" in df.columns:
        df["Hierarquia (Componente)"] = pd.to_numeric(df["Peso"], errors="coerce")

    for col in ["Programa","Família","Cenário de Risco","Verificar","Componente",
                "Tipo de Controle","Categoria","Ameaça",
                "Hierarquia (Controle)","Hierarquia (Componente)"]:
        if col in df.columns and not isinstance(df[col], pd.DataFrame):
            df[col] = df[col].astype(str).str.strip()
    if "Família" in df.columns: df["Familia_norm"] = df["Família"].apply(_canon)
    return df

@st.cache_data(show_spinner=False)
def load_questions() -> pd.DataFrame | None:
    f = _find_file(QUEST_PREFERRED, "*Perguntas*Categoria*.xls*")
    if not f: return None
    xls = pd.ExcelFile(f)
    sheet = "D_Categoria" if "D_Categoria" in xls.sheet_names else xls.sheet_names[0]
    df = xls.parse(sheet)
    df = _rename_questions_cols(df)
    if "Peso_q" in df.columns:
        df["Peso_q"] = (
            df["Peso_q"].astype(str).str.replace(",", ".", regex=False)
            .apply(lambda x: pd.to_numeric(x, errors="coerce"))
        )
    for c in ["Categoria","Pergunta"]:
        if c in df.columns: df[c] = df[c].astype(str).str.strip()
    df = df.dropna(subset=["Categoria","Pergunta"], how="any")
    if "Peso_q" not in df.columns: df["Peso_q"] = 1.0
    return df[["Categoria","Pergunta","Peso_q"]]

# ------------------ output helpers ------------------
EXPECTED_SHEETS = {
    "INSPECTIONS": ["inspection_id","timestamp","complexo","unidade","familia","tag","programa","inspetor","obs"],
    "CONTROLES":   ["inspection_id","controle","etapa1","justificativa"],
    "COMPONENTS":  [
        "inspection_id","controle","componente","categoria","tipo_controle","peso",
        "hierarquia_controle","hierarquia_componente",
        "resposta","justificativa"
    ],
    "QUESTOES":    ["inspection_id","controle","componente","categoria","pergunta","peso_q","resposta","justificativa"],
    "ACTIONS":     ["inspection_id","nivel","controle","componente","descricao","responsavel","prazo"],
    "INDICES_CONTROLE": ["inspection_id","controle","hierarquia_controle","nt_controle","nf_controle","iv_controle","etapa1"],
    "INDICE_ATIVO":     ["inspection_id","iv_ativo","soma_pesos_controle"]
}

def _load_existing_output(path: Path) -> dict:
    sheets = {name: pd.DataFrame(columns=cols) for name, cols in EXPECTED_SHEETS.items()}
    if path.exists():
        xls = pd.ExcelFile(path)
        for name, cols in EXPECTED_SHEETS.items():
            if name in xls.sheet_names:
                df = xls.parse(name)
                for c in cols:
                    if c not in df.columns: df[c] = ""
                sheets[name] = df[cols].copy()
    return sheets

def _save_output(path: Path, sheets: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

# ------------------ cálculo do IV (normalização + guards) ------------------
def _build_category_norms(quest_df: pd.DataFrame):
    """
    norms[cat][pergunta] = fração normalizada (soma por categoria = 1).
    Se não houver pesos válidos numa categoria, divide igualmente entre as perguntas.
    Se uma pergunta/categoria não aparecer em norms, o fallback será fração 1.0 (bloco único).
    """
    q = quest_df.copy()
    q["Peso_q"] = pd.to_numeric(q.get("Peso_q", 1.0), errors="coerce").fillna(0.0)

    norms = {}
    for cat, grp in q.groupby("Categoria"):
        total = grp["Peso_q"].sum()
        if total > 0:
            norms[str(cat)] = {str(r["Pergunta"]): float(r["Peso_q"] / total) for _, r in grp.iterrows()}
        else:
            n = max(len(grp), 1)
            frac = 1.0 / n
            norms[str(cat)] = {str(r["Pergunta"]): frac for _, r in grp.iterrows()}
    return norms

def _build_maps_for_control(block_df: pd.DataFrame):
    comp2hc = {}
    for _, r in block_df.iterrows():
        # tenta primeiro Hierarquia (Componente); se vazio/0 cai para Peso
        hc = r.get("Hierarquia (Componente)")
        hc_val = _to_float(hc, None)
        if hc_val is None or hc_val <= 0:
            hc_val = _to_float(r.get("Peso"), 0.0)
        if hc_val <= 0:
            hc_val = 1.0
        comp2hc[str(r.get("Componente", "-"))] = hc_val

    hctrl_raw = block_df["Hierarquia (Controle)"].dropna().tolist()
    hctrl = _to_float(hctrl_raw[0], 0.0) if hctrl_raw else 0.0
    return comp2hc, hctrl

def compute_iv_per_control(flt_df: pd.DataFrame, quest_df: pd.DataFrame,
                           controles_rows: list, quest_rows: list):
    """
    Para cada CC:
      - Nt = soma dos hc de cada componente (Hierarquia (Componente) ou Peso).
      - Perguntas por componente/categoria são normalizadas (somam 1*hc via frações).
      - Nf soma hc*fração das perguntas "Sim".
      - "Não Aplica" remove hc*fração do Nt.
      - IV_CC = (Nt - Nf) / Nt, com travas numéricas.
    IV_ativo = média ponderada por Hierarquia (Controle), também travado em [0,1].
    """
    EPS = 1e-12

    cat_norms = _build_category_norms(quest_df)
    e1_map = {r["controle"]: r["etapa1"] for r in controles_rows}

    resultados = []
    controles = list(flt_df["Verificar"].dropna().unique())

    soma_pesos_ctrl = 0.0
    soma_pesosxIV   = 0.0

    for ctrl in controles:
        block = flt_df[flt_df["Verificar"] == ctrl]
        comp2hc, hctrl = _build_maps_for_control(block)
        etapa1 = e1_map.get(ctrl, "Existe")

        # Denominador base: soma dos hc dos componentes
        Nt = 0.0
        for _, r in block.iterrows():
            comp = str(r.get("Componente", "-"))
            Nt  += comp2hc.get(comp, 0.0)

        if etapa1 == "Não Existe":
            Nf = 0.0
            IV = 1.0 if Nt > EPS else 0.0

        elif etapa1 == "Não Aplica":
            Nt = 0.0
            Nf = 0.0
            IV = 0.0

        else:
            Nf = 0.0

            rows_ctrl = [q for q in quest_rows if q["controle"] == ctrl]
            for q in rows_ctrl:
                comp = str(q["componente"])
                cat  = str(q.get("categoria", ""))
                perg = str(q.get("pergunta", ""))
                resp = q.get("resposta", "Sim")

                hc = comp2hc.get(comp, 0.0)
                frac = cat_norms.get(cat, {}).get(perg, 1.0)  # fallback bloco único
                peff = hc * float(frac)

                if resp == "Sim":
                    Nf += peff
                elif resp == "Não Aplica":
                    Nt -= peff  # retira do denominador

            # Guard-rails
            if Nt < 0: Nt = 0.0
            if Nf < 0: Nf = 0.0
            if Nf > Nt: Nf = Nt  # evita Nf > Nt por arredondamento

            if Nt <= EPS:
                IV = 0.0
            else:
                IV = (Nt - Nf) / Nt
                if abs(IV) < EPS:  # limpa -0.00%
                    IV = 0.0
                IV = max(0.0, min(1.0, IV))  # clamp em [0,1]

        resultados.append({
            "controle": ctrl,
            "hierarquia_controle": hctrl,
            "nt_controle": round(Nt, 6),
            "nf_controle": round(Nf, 6),
            "iv_controle": round(IV, 6),
            "etapa1": etapa1
        })
        soma_pesos_ctrl += hctrl
        soma_pesosxIV   += hctrl * IV

    iv_ativo = (soma_pesosxIV / soma_pesos_ctrl) if soma_pesos_ctrl > 0 else 0.0
    iv_ativo = max(0.0, min(1.0, iv_ativo))  # trava em [0,1]

    return resultados, iv_ativo, soma_pesos_ctrl

# ------------------ CSS ------------------
st.markdown("""
<style>
.block-container { padding-top: 1rem; }
.q-label { margin: .25rem 0 0 0; }
.q-actions { margin:.25rem 0 .6rem 0; padding:.5rem; background:#f7f9fc; border:1px solid #e9eef5; border-radius:8px; }
.small-note { color:#6b7280; font-size:0.9rem; }
</style>
""", unsafe_allow_html=True)

# ------------------ load ------------------
assets = load_assets()
items  = load_items_cr()
quest  = load_questions()

st.title("Diagnóstico - Checklist & Inspeção")
st.caption("Checklist com perguntas por categoria, ações e armazenamento estruturado.")

with st.sidebar.expander("Arquivos em /data", expanded=False):
    st.write("Dir:", str(DATA_DIR))
    st.write("Arquivos:", ", ".join(p.name for p in DATA_DIR.glob("*.xls*")) or "nenhum .xls")

missing = []
if assets is None: missing.append("Lista de Ativos.xlsx (aba D_Equipamentos) não encontrada ou colunas inválidas.")
if items  is None: missing.append("Diagnóstico.xlsx/Diagnostico.xlsx não encontrado.")
if quest  is None: missing.append("Perguntas - Categoria.xlsx não encontrado (aba D_Categoria).")
if missing:
    for m in missing: st.error(m)
    st.stop()

# ------------------ filtros ------------------
ALL="(Todos)"
st.sidebar.header("Filtro")

comp_opts = sorted(assets["Complexo"].dropna().unique())
comp_sel  = st.sidebar.selectbox("Complexo", [ALL]+comp_opts)
mask = pd.Series(True, index=assets.index)
if comp_sel != ALL: mask &= assets["Complexo"] == comp_sel

unid_opts = sorted(assets.loc[mask,"Unidade"].dropna().unique())
unid_sel  = st.sidebar.selectbox("Unidade", [ALL]+unid_opts)
if unid_sel != ALL: mask &= assets["Unidade"] == unid_sel

fam_opts  = sorted(assets.loc[mask,"Família"].dropna().unique())
if not fam_opts:
    st.error("Nenhuma família encontrada para os filtros atuais."); st.stop()
fam_sel   = st.sidebar.selectbox("Família", fam_opts, index=0)
mask     &= assets["Família"] == fam_sel
fam_sel_norm = _canon(fam_sel)

tag_opts = sorted(assets.loc[mask,"Tag"].dropna().unique())
if not tag_opts:
    st.error("Nenhum TAG encontrado para os filtros atuais."); st.stop()
tag_sel  = st.sidebar.selectbox("TAG (Ativo)", tag_opts)

prog_series = items.loc[items["Familia_norm"] == fam_sel_norm, "Programa"] if "Programa" in items.columns else pd.Series([], dtype=str)
prog_opts = sorted(pd.Series(prog_series).dropna().unique().tolist())
if not prog_opts:
    with st.expander("Ajuda (programas por família)", expanded=False):
        st.write(items.get("Família", pd.Series(dtype=str)).dropna().unique().tolist())
    st.error("Nenhum Programa encontrado para a família selecionada."); st.stop()
prog_sel = st.sidebar.selectbox("Programa", prog_opts)

st.markdown(f"**Complexo:** {comp_sel} · **Unidade:** {unid_sel} · **Família:** {fam_sel} · **TAG:** {tag_sel} · **Programa:** {prog_sel}")

st.sidebar.markdown("---")
inspetor  = st.sidebar.text_input("Inspetor", "")
obs_geral = st.sidebar.text_area("Observação geral", "")

# ------------------ catálogo + checklist ------------------
flt_mask = (items["Familia_norm"] == fam_sel_norm)
if "Programa" in items.columns: flt_mask &= (items["Programa"] == prog_sel)
flt = items[flt_mask].copy()

st.subheader("Itens do diagnóstico (catálogo)")
cols_show = [c for c in ["Verificar","Componente","Tipo de Controle","Categoria","Ameaça"] if c in flt.columns]
st.dataframe(flt[cols_show].reset_index(drop=True), use_container_width=True)

st.subheader("Checklist")

controles_rows = []
component_rows = []
quest_rows     = []
actions_rows   = []

if flt.empty:
    st.info("Nenhum item para este filtro.")
else:
    for g in list(flt["Verificar"].dropna().unique()):
        block = flt[flt["Verificar"] == g].reset_index(drop=True)
        with st.expander(g, expanded=False):
            # Etapa 1
            e1_key = _mk_widget_key("e1", g, 0)
            e1 = st.radio(f'Existe "{g}"?', ["Existe", "Não Existe", "Não Aplica"], index=0, key=e1_key)

            justificativa_e1 = ""
            if e1 == "Não Aplica":
                justificativa_e1 = st.text_area("Justificativa - Não Aplica (obrigatório)", key=_mk_widget_key("just", g, 0), height=80)
            if e1 == "Não Existe":
                st.warning("Ação obrigatória: controle inexistente.")
                a1, a2, a3 = _cols_top([2,1,1])
                adesc = a1.text_area("Descrição da ação", key=_mk_widget_key("acao_desc", g, 0), height=80)
                aresp = a2.text_input("Responsável", key=_mk_widget_key("acao_resp", g, 0))
                apraz = a3.date_input("Prazo", key=_mk_widget_key("acao_prazo", g, 0))
                actions_rows.append({"inspection_id":"__PENDING__","nivel":"controle","controle":g,"componente":"","descricao":adesc,"responsavel":aresp,"prazo":str(apraz) if apraz else ""})
            controles_rows.append({"inspection_id":"__PENDING__","controle":g,"etapa1":e1,"justificativa":justificativa_e1})

            # Etapa 2
            if e1 == "Existe":
                st.markdown("**Etapa 2 – Componentes do controle**")

                for i, row in block.iterrows():
                    comp  = str(row.get("Componente","-"))
                    cat   = str(row.get("Categoria","-"))
                    tctl  = str(row.get("Tipo de Controle","-"))
                    h_ctrl = str(row.get("Hierarquia (Controle)",""))
                    h_comp = str(row.get("Hierarquia (Componente)",""))
                    peso_raw = row.get("Peso", None)
                    try: peso_int = int(peso_raw) if pd.notna(peso_raw) else None
                    except Exception: peso_int = None

                    c1, c2, c3, c4 = _cols_top([2,1,1,2])
                    c1.markdown(f"**{comp}**")
                    c2.write(cat)
                    c3.write(tctl)
                    c4.markdown("**Perguntas (Categoria)**")

                    qcat = quest[quest["Categoria"] == cat]
                    if qcat.empty:
                        c4.info("Sem perguntas para esta categoria.")
                    else:
                        for j, qrow in qcat.reset_index(drop=True).iterrows():
                            qtxt  = str(qrow["Pergunta"])
                            qpeso = float(qrow["Peso_q"]) if pd.notna(qrow["Peso_q"]) else 1.0

                            ql, qr = _cols_top([3,1])
                            ql.markdown(f"<div class='q-label'>{qtxt}</div>", unsafe_allow_html=True)
                            qresp = qr.selectbox(
                                "Resposta", ["Sim","Não","Não Aplica"],
                                key=_mk_widget_key("q", g, comp, qtxt, i, j),
                                label_visibility="collapsed",
                            )

                            justificativa_q = ""
                            if qresp == "Não":
                                a1, a2, a3 = _cols_top([2,1,1])
                                a1.markdown("<div class='q-actions'>", unsafe_allow_html=True)
                                adesc3 = a1.text_area("Descrição da ação (questão)", key=_mk_widget_key("acao_desc_q", g, comp, qtxt, i, j), height=80, label_visibility="visible")
                                a1.markdown("</div>", unsafe_allow_html=True)
                                aresp3 = a2.text_input("Responsável", key=_mk_widget_key("acao_resp_q", g, comp, qtxt, i, j))
                                apraz3 = a3.date_input("Prazo", key=_mk_widget_key("acao_prazo_q", g, comp, qtxt, i, j))
                                actions_rows.append({"inspection_id":"__PENDING__","nivel":"questao","controle":g,"componente":comp,"descricao":adesc3,"responsavel":aresp3,"prazo":str(apraz3) if apraz3 else ""})
                            elif qresp == "Não Aplica":
                                justificativa_q = st.text_area("Justificativa (N/A)", key=_mk_widget_key("just_q", g, comp, qtxt, i, j), height=70)

                            quest_rows.append({
                                "inspection_id":"__PENDING__","controle":g,"componente":comp,"categoria":cat,
                                "pergunta":qtxt,"peso_q":qpeso,"resposta":qresp,"justificativa":justificativa_q
                            })

                    component_rows.append({
                        "inspection_id":"__PENDING__","controle":g,"componente":comp,"categoria":cat,"tipo_controle":tctl,
                        "peso":peso_int,"hierarquia_controle":h_ctrl,"hierarquia_componente":h_comp,"resposta":"","justificativa":""
                    })

                    st.divider()

# ------------------ salvar ------------------
def _validate(controles_rows, quest_rows, actions_rows):
    errs = []
    for r in controles_rows:
        if r["etapa1"] == "Não Aplica" and not r.get("justificativa"):
            errs.append(f"Controle '{r['controle']}' N/A sem justificativa.")
        if r["etapa1"] == "Não Existe":
            has = any(a for a in actions_rows if a["nivel"]=="controle" and a["controle"]==r["controle"] and a.get("descricao"))
            if not has: errs.append(f"Controle '{r['controle']}' marcado como 'Não Existe' sem ação.")
    for q in quest_rows:
        if q["resposta"] == "Não Aplica" and not q.get("justificativa"):
            errs.append(f"Questão '{q['pergunta']}' do componente '{q['componente']}' N/A sem justificativa.")
        if q["resposta"] == "Não":
            has = any(a for a in actions_rows if a["nivel"]=="questao" and a["controle"]==q["controle"] and a["componente"]==q["componente"] and a.get("descricao"))
            if not has: errs.append(f"Questão '{q['pergunta']}' do componente '{q['componente']}' com resposta 'Não' sem ação.")
    return errs

left, right = _cols_top([1,1])
with left:  salvar = st.button("Salvar inspeção")
with right: st.caption("Feche o Excel/preview antes de salvar.")

if salvar:
    errors = _validate(controles_rows, quest_rows, actions_rows)
    if not inspetor: errors.append("Preencha o nome do inspetor (sidebar).")
    if errors:
        st.error("Não foi possível salvar. Corrija os itens abaixo:")
        for e in errors: st.write("- ", e)
        st.stop()

    inspection_id = datetime.utcnow().strftime("%Y%m%d%H%M%S%f") + "_" + _safe_key(tag_sel)
    ts = datetime.utcnow().isoformat()

    ins_df = pd.DataFrame([{
        "inspection_id": inspection_id, "timestamp": ts,
        "complexo": comp_sel, "unidade": unid_sel, "familia": fam_sel,
        "tag": tag_sel, "programa": prog_sel, "inspetor": inspetor, "obs": obs_geral,
    }])

    for lst in (controles_rows, component_rows, quest_rows, actions_rows):
        for r in lst: r["inspection_id"] = inspection_id

    # Recalcula índices e prepara saída
    indices_ctrl, iv_ativo, soma_pesos_ctrl = compute_iv_per_control(flt, quest, controles_rows, quest_rows)
    idx_ctrl_df = pd.DataFrame(indices_ctrl)
    idx_ativo_df = pd.DataFrame([{
        "inspection_id": inspection_id,
        "iv_ativo": round(float(iv_ativo), 6),
        "soma_pesos_controle": round(float(soma_pesos_ctrl), 6)
    }])

    existing = _load_existing_output(OUT_FILE)
    existing["INSPECTIONS"] = pd.concat([existing["INSPECTIONS"], ins_df], ignore_index=True)
    if controles_rows: existing["CONTROLES"]  = pd.concat([existing["CONTROLES"],  pd.DataFrame(controles_rows)], ignore_index=True)
    if component_rows: existing["COMPONENTS"] = pd.concat([existing["COMPONENTS"], pd.DataFrame(component_rows)], ignore_index=True)
    if quest_rows:     existing["QUESTOES"]   = pd.concat([existing["QUESTOES"],   pd.DataFrame(quest_rows)], ignore_index=True)
    if actions_rows:   existing["ACTIONS"]    = pd.concat([existing["ACTIONS"],    pd.DataFrame(actions_rows)], ignore_index=True)
    if not idx_ctrl_df.empty:
        idx_ctrl_df["inspection_id"] = inspection_id
        existing["INDICES_CONTROLE"] = pd.concat([existing["INDICES_CONTROLE"], idx_ctrl_df], ignore_index=True)
    existing["INDICE_ATIVO"] = pd.concat([existing["INDICE_ATIVO"], idx_ativo_df], ignore_index=True)

    try:
        with st.spinner("Salvando..."):
            _save_output(OUT_FILE, existing)
        st.success(f"Inspeção salva em: {OUT_FILE}")
    except PermissionError:
        st.error("Não consegui salvar porque o arquivo está aberto. Feche o 'Inspecoes.xlsx' e tente novamente.")
    else:
        # RESULTADOS (só APÓS salvar) ------------------------
        st.markdown("### Resultados – Índice de Vulnerabilidade")
        if not idx_ctrl_df.empty:
            df_show = idx_ctrl_df.rename(columns={
                "controle":"Controle",
                "hierarquia_controle":"Hierarquia (Ctrl.)",
                "iv_controle":"Índice de Vulnerabilidade (CC)",
                "etapa1":"Etapa 1"
            })
            # clamp no output para evitar "-0,00%"
            df_show["Índice de Vulnerabilidade (CC)"] = (
                pd.to_numeric(df_show["Índice de Vulnerabilidade (CC)"], errors="coerce")
                .clip(lower=0, upper=1)
                .apply(lambda x: f"{float(x):.2%}")
            )
            df_show = df_show[["Controle","Hierarquia (Ctrl.)","Índice de Vulnerabilidade (CC)","Etapa 1"]]
            st.dataframe(df_show, use_container_width=True)

        # métrica do ativo (também clampada)
        iv_fmt = max(0.0, min(1.0, float(iv_ativo)))
        st.metric(label="Índice de Vulnerabilidade do Ativo", value=f"{iv_fmt:.2%}")
