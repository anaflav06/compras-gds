import streamlit as st
import pandas as pd
import json
import os
import base64
import requests
import hashlib
import hmac
from io import BytesIO
from datetime import datetime, date
from calendar import month_name

# Exportação Excel/PDF
import xlsxwriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.units import mm

# ==========================================================
# CONFIGURAÇÃO
# ==========================================================
st.set_page_config(
    page_title="Relatório de Compras Mensal",
    page_icon="🛒",
    layout="wide",
)

APP_TITLE = "RELATÓRIO DE COMPRAS MENSAL"
LOCAL_DB = "database_compras_mensal.json"
DEFAULT_GITHUB_DB_PATH = "database_compras_mensal.json"

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

CENTROS_CUSTO_PADRAO = ["CDSP2", "CPQ08", "SAO12", "ASSISTÊNCIA"]
FORMAS_PAGAMENTO_PADRAO = ["2260", "2688"]
SITES_PADRAO = ["MERCADO LIVRE"]
STATUS_DEVOLUCAO = [
    "Devolução solicitada",
    "Produto devolvido",
    "Aguardando reembolso",
    "Reembolso confirmado",
]

DATA_INICIO_APP = date(2026, 7, 1)

# Credenciais armazenadas como hashes derivados (as senhas não ficam em texto puro no código).
LOGIN_CREDENTIALS = {
    "julia": {
        "salt": "gds-compras-julia-2026",
        "hash": "0dea791a18ede08cf8a693ec0b90ef296c7aa8208b97cfee8f6ca284a1982034",
    },
    "jessica": {
        "salt": "gds-compras-jessica-2026",
        "hash": "5c184a8a6d44af369a48f42dacf4d00e978532bb2d6fe534657002373be0ab5b",
    },
}

def validar_login(usuario, senha):
    usuario = (usuario or "").strip().lower()
    registro = LOGIN_CREDENTIALS.get(usuario)
    if not registro:
        return False
    digest = hashlib.pbkdf2_hmac(
        "sha256", senha.encode("utf-8"), registro["salt"].encode("utf-8"), 200_000
    ).hex()
    return hmac.compare_digest(digest, registro["hash"])

def exigir_login():
    if st.session_state.get("logado"):
        return
    st.markdown(f'<div class="app-title">🛒 {APP_TITLE}</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Acesso restrito • Controle mensal de compras</div>', unsafe_allow_html=True)
    box1, box2, box3 = st.columns([1, 1.2, 1])
    with box2:
        st.subheader("🔐 Entrar")
        with st.form("login_form"):
            usuario = st.text_input("Usuário")
            senha = st.text_input("Senha", type="password")
            entrar = st.form_submit_button("Entrar", use_container_width=True, type="primary")
        if entrar:
            if validar_login(usuario, senha):
                st.session_state["logado"] = True
                st.session_state["usuario"] = usuario.strip().lower()
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")
    st.stop()

def meses_validos_para_ano(ano):
    if int(ano) == 2026:
        return list(range(7, 13))
    if int(ano) > 2026:
        return list(range(1, 13))
    return []

def anos_validos(df=None):
    hoje = date.today()
    anos = {2026, hoje.year}
    if df is not None and not df.empty and "data" in df.columns:
        anos.update(int(x) for x in df["data"].dt.year.dropna().unique())
    return sorted([a for a in anos if a >= 2026], reverse=True)

def nf_esta_pendente(c):
    return not str(c.get("numero_nf") or "").strip() and c.get("status_nf") != "Finalizada sem NF"

# ==========================================================
# ESTILO
# ==========================================================
st.markdown(
    """
    <style>
    .block-container {padding-top: 3.8rem !important; padding-bottom: 2rem;}
    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #E5E7EB;
        padding: 14px 16px;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,.04);
    }
    .app-title {
        font-size: 1.75rem;
        font-weight: 800;
        line-height: 1.25;
        margin-top: .4rem;
        margin-bottom: .35rem;
        overflow: visible;
    }
    .subtitle {color:#6B7280; margin-bottom:1rem;}
    .pending-box {
        border:1px solid #F3D49B;
        background:#FFFBEB;
        padding:10px 12px;
        border-radius:10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================================
# BANCO LOCAL + GITHUB
# ==========================================================
def _secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


def github_configured():
    return bool(_secret("GITHUB_TOKEN") and _secret("GITHUB_REPO"))


def github_params():
    return {
        "token": _secret("GITHUB_TOKEN"),
        "repo": _secret("GITHUB_REPO"),
        "branch": _secret("GITHUB_DATA_BRANCH", "main"),
        "path": _secret("GITHUB_COMPRAS_DB_PATH", DEFAULT_GITHUB_DB_PATH),
    }


def load_from_github():
    if not github_configured():
        return None
    cfg = github_params()
    url = f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['path']}"
    headers = {
        "Authorization": f"Bearer {cfg['token']}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    r = requests.get(url, headers=headers, params={"ref": cfg["branch"]}, timeout=15)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    payload = r.json()
    content = base64.b64decode(payload["content"]).decode("utf-8")
    data = json.loads(content)
    data["_github_sha"] = payload.get("sha")
    return data


def save_to_github(data):
    if not github_configured():
        return False, "GitHub não configurado"
    cfg = github_params()
    url = f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['path']}"
    headers = {
        "Authorization": f"Bearer {cfg['token']}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    sha = data.get("_github_sha")
    clean = {k: v for k, v in data.items() if k != "_github_sha"}
    body = {
        "message": f"Atualiza banco de compras {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        "content": base64.b64encode(
            json.dumps(clean, ensure_ascii=False, indent=2).encode("utf-8")
        ).decode("ascii"),
        "branch": cfg["branch"],
    }
    if sha:
        body["sha"] = sha
    r = requests.put(url, headers=headers, json=body, timeout=20)
    if r.status_code in (200, 201):
        data["_github_sha"] = r.json().get("content", {}).get("sha")
        return True, "Banco permanente atualizado no GitHub"
    return False, f"GitHub: {r.status_code} - {r.text[:180]}"


def load_local():
    if os.path.exists(LOCAL_DB):
        try:
            with open(LOCAL_DB, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return None


def save_local(data):
    clean = {k: v for k, v in data.items() if k != "_github_sha"}
    with open(LOCAL_DB, "w", encoding="utf-8") as f:
        json.dump(clean, f, ensure_ascii=False, indent=2)


def blank_db():
    return {
        "compras": [],
        "config": {
            "centros_custo": CENTROS_CUSTO_PADRAO,
            "formas_pagamento": FORMAS_PAGAMENTO_PADRAO,
            "sites": SITES_PADRAO,
        },
        "meta": {
            "criado_em": datetime.now().isoformat(timespec="seconds"),
            "versao": 1,
        },
    }


# Carga inicial baseada na planilha COMPRAS - GDSLOGISTICA - JULHO 2026.xlsx
SEED_JULHO_2026 = [
("MERCADO LIVRE","CPQ08","2026-07-02","000.091.043","FAROL DIANTEIRO","2260",2,206.10,206.10,"Reembolso confirmado"),
("MERCADO LIVRE","CPQ08","2026-07-02","000.205.644","EXTENSÃO ELETRO MIXER","2260",1,38.00,0,""),
("MERCADO LIVRE","CDSP2","2026-07-02","000.058.123","4 CABOS DE ALUMÍNIO","2260",1,68.06,0,""),
("MERCADO LIVRE","CDSP2","2026-07-02","000.010.966","ASSENTO SANITÁRIO OVAL","2260",1,39.92,0,""),
("MERCADO LIVRE","CDSP2","2026-07-02","000.010.965","ASSENTO SANITÁRIO OVAL","2260",1,119.76,0,""),
("MERCADO LIVRE","CDSP2","2026-07-02","000.047.661","ESCOVA SANITÁRIA VASSOURINHA","2260",1,79.50,0,""),
("MERCADO LIVRE","CPQ08","2026-07-03","000.232.198","BOTA SEGURANÇA COURO","2260",1,59.96,0,""),
("MERCADO LIVRE","CPQ08","2026-07-03","000.232.200","BOTA SEGURANÇA COURO","2260",1,59.96,0,""),
("MERCADO LIVRE","CPQ08","2026-07-06","000.062.226","KIT 6 GARRAFAS SQUEEZE","2260",2,33.90,0,""),
("MERCADO LIVRE","CPQ08","2026-07-06","000.232.264","BOTA SEGURANÇA COURO","2260",1,59.96,0,""),
("MERCADO LIVRE","CPQ08","2026-07-06","000.080.978","TRANSMISSOR RECEPTOR","2260",1,271.00,0,""),
("MERCADO LIVRE","SAO12","2026-07-07","000.012.616","CILINDRO PARA DCP -I2540DW","2260",3,82.54,0,""),
("MERCADO LIVRE","CPQ08","2026-07-07","000.047.941","ESCOVA SANITÁRIA VASSOURINHA","2260",1,59.65,0,""),
("MERCADO LIVRE","CPQ08","2026-07-07","000.118.848","CORDEL VELAME COM 10M","2260",1,29.09,0,""),
("MERCADO LIVRE","CPQ08","2026-07-08","1227397","CARRINHO PARA CPU","2260",1,30.99,0,""),
("MERCADO LIVRE","CPQ08","2026-07-08","000.232.350","BOTA SEGURANÇA COURO","2260",1,59.96,0,""),
("MERCADO LIVRE","CPQ08","2026-07-08","000.232.348","BOTA SEGURANÇA COURO","2260",1,59.96,0,""),
("MERCADO LIVRE","CPQ08","2026-07-10","000.042.446","BOTA SEGURANÇA COURO","2688",1,48.99,0,""),
("MERCADO LIVRE","CDSP2","2026-07-10","000.113.156","CANETA ESFEROGRAFICA BIC","2688",1,77.50,0,""),
("MERCADO LIVRE","CDSP2","2026-07-10","000.027.442","CABO VGA MACHO PARA MONITOR","2688",1,38.00,0,""),
("MERCADO LIVRE","CDSP2","2026-07-10","000.356.337","CARREGADOR TURBO 55W","2688",1,78.00,0,""),
("MERCADO LIVRE","CDSP2","2026-07-13","000.232.484","BOTA SEGURANÇA COURO","2688",1,59.96,0,""),
("MERCADO LIVRE","CDSP2","2026-07-14","000.232.496","BOTA SEGURANÇA COURO","2688",1,59.96,0,""),
("MERCADO LIVRE","CPQ08","2026-07-14","000.106.236","VASSOURA CERDAS MACIAS","2688",1,20.50,0,""),
("MERCADO LIVRE","CDSP2","2026-07-15","000.107.475","TONER COMPATIVEL MS310","2688",1,67.36,0,""),
("MERCADO LIVRE","CDSP2","2026-07-15","000.002.411","TONERCOMPATIVEL CE505A","2688",1,119.70,0,""),
("MERCADO LIVRE","SAO12","2026-07-16","000.089.783","ADAPTADOR USB WIFI","2688",1,59.90,0,""),
("MERCADO LIVRE","CPQ08","2026-07-17","000.232.617","BOTA SEGURANÇA COURO","2260",1,59.96,0,""),
("MERCADO LIVRE","CPQ08","2026-07-17","000.232.581","BOTA SEGURANÇA COURO","2260",1,59.96,0,""),
("MERCADO LIVRE","CPQ08","2026-07-20","000.232.642","BOTA SEGURANÇA COURO","2688",1,59.96,0,""),
("MERCADO LIVRE","CPQ08","2026-07-20","","BOTA SEGURANÇA COURO","2688",1,59.96,59.96,"Reembolso confirmado"),
("MERCADO LIVRE","CPQ08","2026-07-20","000.063.355","KIT 6 GARRAFAS SQUEEZE","2260",2,33.90,0,""),
("MERCADO LIVRE","CPQ08","2026-07-21","000.232.664","BOTA SEGURANÇA COURO","2688",1,59.96,0,""),
("MERCADO LIVRE","CDSP2","2026-07-21","139895420","APARELHO DE JANTAR","2260",6,218.32,218.32,"Reembolso confirmado"),
("MERCADO LIVRE","CDSP2","2026-07-24","","CARTUCHO TONER BYQUALY","2260",1,104.95,0,""),
("MERCADO LIVRE","SAO12","2026-07-24","141096856","APARELHO DE JANTAR","2260",6,218.51,0,""),
("MERCADO LIVRE","CDSP2","2026-07-27","000.934.839","CÁPSULA CAFÉ ESPRESSO","2260",1,21.31,0,""),
("MERCADO LIVRE","CDSP2","2026-07-27","23426373","CÁPSULA DE CAFÉ CERRADO E DIFUSOR","2260",1,51.21,0,""),
("MERCADO LIVRE","CDSP2","2026-07-27","000.065.543","CÁPSULA CAFÉ ESPRESSO","2260",1,26.23,0,""),
("MERCADO LIVRE","CDSP2","2026-07-27","","CALÇA LEGGING LEG","2260",1,100.68,0,""),
("MERCADO LIVRE","CDSP2","2026-07-28","000.238.263","CAT6 CABO DE REDE 5M","2260",1,45.12,0,""),
("MERCADO LIVRE","CDSP2","2026-07-28","000.233.320","BOTA SEGURANÇA COURO","2260",1,58.99,0,""),
("MERCADO LIVRE","CDSP2","2026-07-28","000.233.322","BOTA SEGURANÇA COURO","2260",1,59.99,0,""),
("MERCADO LIVRE","CDSP2","2026-07-29","000.007.784","10 ETIQUETAS TERMOCOLANTE","2260",1,23.22,0,""),
("MERCADO LIVRE","CDSP2","2026-07-29","000.015.702","BLUSA DE FRIO MOLEROM","2260",4,140.97,0,""),
("MERCADO LIVRE","CDSP2","2026-07-29","","MOUSEPAD EXBOM","2260",5,155.00,155.00,"Reembolso confirmado"),
("MERCADO LIVRE","CDSP2","2026-07-29","","MOUSEPAD EXBOM","2260",5,155.00,0,""),
("MERCADO LIVRE","CDSP2","2026-07-29","000.157.220","KIT 10 LÂMPADAS LED","2260",1,59.99,0,""),
("MERCADO LIVRE","CDSP2","2026-07-29","000.033.402","20 PILHAS ALCALINAS","2260",1,57.48,0,""),
("MERCADO LIVRE","CDSP2","2026-07-29","","KIT 10 RELÓGIO BRACELETE","2260",1,110.12,0,""),
("MERCADO LIVRE","CDSP2","2026-07-29","","KIT 30 SAQUINHO ORGANZA 13X18","2260",1,21.99,0,""),
("MERCADO LIVRE","CPQ08","2026-07-31","000.213.621","BOTINA DE SEGURANÇA","2260",1,59.90,59.90,"Reembolso confirmado"),
("MERCADO LIVRE","SAO12","2026-07-31","000.176.164","KIT 3 PALHETA SILICONE LIMPADOR","2260",1,37.80,0,""),
("MERCADO LIVRE","CDSP2","2026-07-31","3504618","LIMPADOR DE PARA BRISA","2260",1,78.99,0,""),
("MERCADO LIVRE","CDSP2","2026-07-31","000.313.601","KIT10 PORTA RETRATO 15X21","2260",1,75.51,0,""),
]



# Carga inicial baseada na planilha COMPRAS - GDSLOGISTICA - AGOSTO 2026.xlsx
SEED_AGOSTO_2026 = [
('MERCADO LIVRE', 'CPQ08', '2026-08-03', '000.100.441', 'KIT 50 CHAVEIROS', '2260', 1, 39.9, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-03', '1523023', 'KIT 2 ESPANADOR', '2260', 1, 33.93, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-03', '000.013.241', 'HEADSET MULTILASER', '2260', 1, 28.86, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-03', '000.015.846', 'BLUSA DE FRIO MOLETOM', '2260', 1, 328.93, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-03', '000.000.448', 'KIT 4 PLACAS BANHEIROS', '2260', 1, 35.79, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-03', '000.008.857', 'CABO ALONGADOR 2 METROS', '2260', 1, 37.99, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-03', '000.060.182', 'KIT 6 GARRAFAS SQUEEZE', '2260', 2, 33.9, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-03', '000.060.183', 'KIT 6 GARRAFAS SQUEEZE', '2260', 2, 33.9, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-04', '000.163.408', 'CASE CAPINHA CAPA', '2260', 1, 19.99, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-04', '000.214.469', 'BOTINA DE SEGURANÇA', '2260', 1, 69.9, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-04', '000.214.468', 'BOTINA DE SEGURANÇA', '2260', 1, 69.9, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-04', '000.015.635', 'FONE C MICROFONE', '2260', 1, 39.9, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-05', '000.018.319', 'SAQUINHO PLÁSTICO 25X35', '2260', 1, 19.9, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-05', '000.233.858', 'BOTA SEGURANÇA COURO', '2260', 1, 59.99, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-05', '000.001.220', 'ROUPEIRO LOCKER 4 PORTAS', '2260', 5, 543.2, 0.0, '', ''),
('MERCADO LIVRE', 'SAO12', '2026-08-05', '000.061.291', 'KIT PAR PALHETA LIMPADOR', '2260', 1, 28.49, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-05', '000.031.762', 'TECLADO USB', '2260', 1, 29.49, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-05', '000.083.232', 'KIT C/10 PASTAS SUSPENSA', '2260', 1, 52.8, 0.0, '', ''),
('MERCADO LIVRE', 'SAO12', '2026-08-05', '000.041.725', 'PLACA PARA BANHEIROS', '2260', 2, 39.0, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-05', '', 'CESTO RETANGULAR', '2260', 1, 32.17, 0.0, '', 'AGUARDANDO A NF'),
('MERCADO LIVRE', 'CPQ08', '2026-08-05', '000.207.932', 'DISPENSER SUPORTE DE PAPEL', '2260', 1, 107.0, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-05', '000.001.288', 'MOLETON  GOLA CARECA', '2260', 1, 67.86, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-05', '000.002.409', 'CAMISETA DRY FIT', '2260', 1, 34.0, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-06', '1163276', '12 PR LUVA EPI SEGURANÇA', '2260', 1, 38.69, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-06', '000.122.591', '1 FONTE DE ENERGIA 12V', '2260', 2, 47.22, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-06', '000.032.136', 'TECLADO USB EXBOM  BK 102', '2260', 1, 58.98, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-06', '000.020.556', 'KIT 4 RODINHAS PARA CADEIRA', '2260', 1, 38.58, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-06', '000.002.565', 'PASTA CATÁLOGO A4', '2260', 1, 28.9, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-07', '000.028.825', '50 FAIXA REFLETIVA LATERAL', '2260', 4, 107.57, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-07', '000.023.897', 'FAIXA REFLETIVA ADESIVA', '2260', 1, 56.9, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-12', '000.024.120', 'FAIXA REFLETIVA ADESIVA', '2260', 1, 59.9, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-12', '000.398.641', 'KIT 10 MOLDURA PORTA RETRATO', '2260', 3, 64.9, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-13', '000.024.909', 'SOPRADOR TÉRMICO', '2260', 3, 93.11, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-14', '000.011.572', 'PANO DE CHÃO GROSSO', '2260', 1, 59.0, 0.0, '', ''),
('MERCADO LIVRE', 'SAO12', '2026-08-14', '000.685.619', 'FONE DE OUVIDO HEADSET', '2260', 1, 40.0, 0.0, '', ''),
('MERCADO LIVRE', 'CPQ08', '2026-08-17', '1352765', 'SIFÃO SANFONADO AJUSTAVEL', '2260', 1, 19.0, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-18', '000.029.922', 'SOPRADOR TÉRMICO', '2260', 3, 93.11, 0.0, '', ''),
('MERCADO LIVRE', 'CDSP2', '2026-08-18', '000.007.961', 'MOTOROLA MOTO G06', '2260', 7, 765.0, 0.0, '', ''),
]


def _seed_compra_dict(row, seed_id, observacao_padrao):
    # Julho tem 10 campos; agosto tem 11 (o último preserva observação da planilha).
    if len(row) == 10:
        site, cc, data_compra, nf, item, pagamento, parcelas, total, devolucao, status = row
        obs_planilha = ""
    else:
        site, cc, data_compra, nf, item, pagamento, parcelas, total, devolucao, status, obs_planilha = row

    now = datetime.now().isoformat(timespec="seconds")
    observacao = observacao_padrao
    if obs_planilha:
        observacao += f" | {obs_planilha}"

    return {
        "id": seed_id,
        "site": site,
        "centro_custo": cc,
        "data_compra": data_compra,
        "numero_nf": str(nf or ""),
        "status_nf": "Recebida" if str(nf or "").strip() else "Pendente",
        "motivo_nf": "",
        "item": item,
        "forma_pagamento": str(pagamento),
        "parcelamento": int(parcelas),
        "total": float(total),
        "houve_devolucao": bool(devolucao),
        "valor_devolucao": float(devolucao or 0),
        "status_devolucao": status,
        "observacao": observacao,
        "criado_em": now,
        "atualizado_em": now,
        "usuario_criacao": "IMPORTAÇÃO INICIAL",
        "usuario_atualizacao": "IMPORTAÇÃO INICIAL",
    }


def apply_seed_if_empty(db):
    """
    Mantém as bases iniciais de julho e agosto/2026 sem duplicar registros.
    Cada compra importada tem um ID fixo por mês. Assim, uma atualização do app
    pode acrescentar a base de agosto mesmo se julho já estiver no banco.
    """
    db.setdefault("compras", [])
    existing_ids = {str(c.get("id", "")) for c in db["compras"]}
    added = 0

    # Julho
    for i, row in enumerate(SEED_JULHO_2026, 1):
        seed_id = f"JUL2026-{i:03d}"
        if seed_id not in existing_ids:
            db["compras"].append(
                _seed_compra_dict(row, seed_id, "Importado da planilha de julho/2026")
            )
            existing_ids.add(seed_id)
            added += 1

    # Agosto - base enviada em 24/08/2026
    for i, row in enumerate(SEED_AGOSTO_2026, 1):
        seed_id = f"AGO2026-{i:03d}"
        if seed_id not in existing_ids:
            db["compras"].append(
                _seed_compra_dict(row, seed_id, "Importado da planilha de agosto/2026")
            )
            existing_ids.add(seed_id)
            added += 1

    db.setdefault("meta", {})["seed_julho_2026"] = True
    db.setdefault("meta", {})["seed_agosto_2026"] = True
    db["meta"]["ultima_carga_inicial_em"] = datetime.now().isoformat(timespec="seconds")
    db["meta"]["ultima_carga_inicial_adicionados"] = added
    return db


def load_database():
    remote = None
    try:
        remote = load_from_github()
    except Exception as e:
        st.session_state["sync_warning"] = f"Não foi possível ler o banco do GitHub: {e}"

    db = remote or load_local() or blank_db()
    db.setdefault("compras", [])
    db.setdefault("config", {})
    db["config"].setdefault("centros_custo", CENTROS_CUSTO_PADRAO.copy())
    db["config"].setdefault("formas_pagamento", FORMAS_PAGAMENTO_PADRAO.copy())
    db["config"].setdefault("sites", SITES_PADRAO.copy())
    for c in db["compras"]:
        c.setdefault("status_nf", "")
        c.setdefault("motivo_nf", "")
        c.setdefault("usuario_criacao", c.get("usuario_criacao", ""))
        c.setdefault("usuario_atualizacao", c.get("usuario_atualizacao", ""))
    before_seed = len(db["compras"])
    db = apply_seed_if_empty(db)
    added_seed = len(db["compras"]) - before_seed
    save_local(db)
    if added_seed > 0 and github_configured():
        try:
            save_to_github(db)
        except Exception:
            pass
    return db


def persist(db, show=True):
    save_local(db)
    if github_configured():
        try:
            ok, msg = save_to_github(db)
            if show:
                if ok:
                    st.success("✅ Dados salvos no banco permanente.")
                else:
                    st.warning(f"Salvo localmente, mas houve falha na nuvem. {msg}")
            return ok
        except Exception as e:
            if show:
                st.warning(f"Salvo localmente, mas não foi possível atualizar o GitHub: {e}")
            return False
    if show:
        st.warning("⚠️ Dados salvos apenas neste computador. Para o Streamlit Cloud, configure o GitHub nos Secrets para garantir persistência permanente.")
    return True


exigir_login()

if "db_compras" not in st.session_state:
    st.session_state.db_compras = load_database()

db = st.session_state.db_compras

# ==========================================================
# FUNÇÕES DE DADOS
# ==========================================================
def brl(v):
    try:
        s = f"{float(v):,.2f}"
        return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def compra_liquida(c):
    total = float(c.get("total") or 0)
    devolucao = float(c.get("valor_devolucao") or 0)
    # Só reduz o gasto efetivo quando o reembolso foi confirmado.
    if c.get("status_devolucao") == "Reembolso confirmado":
        return max(0.0, total - devolucao)
    return total


def compras_df():
    rows = []
    for c in db.get("compras", []):
        d = dict(c)
        try:
            dt = pd.to_datetime(d.get("data_compra"))
        except Exception:
            dt = pd.NaT
        d["data"] = dt
        d["valor_liquido"] = compra_liquida(c)
        rows.append(d)
    return pd.DataFrame(rows)


def filter_month(df, ano, mes):
    if df.empty:
        return df.copy()
    return df[(df["data"].dt.year == ano) & (df["data"].dt.month == mes)].copy()


def add_option(config_key, value):
    value = (value or "").strip().upper()
    if not value:
        return
    arr = db["config"].setdefault(config_key, [])
    if value not in arr:
        arr.append(value)
        arr.sort()
        persist(db, show=False)


def next_id():
    stamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
    return f"CMP-{stamp}"

# ==========================================================
# SIDEBAR / MENU
# ==========================================================
st.sidebar.markdown(f"### 🛒 {APP_TITLE.title()}")
st.sidebar.caption(f"👤 Acesso: **{st.session_state.get('usuario','')}**")
if st.sidebar.button("🚪 Sair", use_container_width=True):
    st.session_state.pop("logado", None)
    st.session_state.pop("usuario", None)
    st.rerun()

menu = st.sidebar.radio(
    "Menu",
    ["📊 Dashboard", "🛒 Inserir compras", "📥 Exportar relatório"],
    label_visibility="collapsed",
)

if github_configured():
    st.sidebar.success("☁️ Banco permanente: GitHub")
else:
    st.sidebar.warning("💾 Banco local")

if st.sidebar.button("🔄 Recarregar banco", use_container_width=True):
    st.session_state.db_compras = load_database()
    st.rerun()

st.markdown(f'<div class="app-title">🛒 {APP_TITLE}</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Controle simples de compras, notas fiscais, devoluções e fechamento mensal.</div>', unsafe_allow_html=True)

if st.session_state.get("sync_warning"):
    st.warning(st.session_state.pop("sync_warning"))

# ==========================================================
# DASHBOARD
# ==========================================================
if menu == "📊 Dashboard":
    hoje = date.today()
    df = compras_df()
    c1, c2 = st.columns([1, 1])
    with c2:
        anos_disponiveis = anos_validos(df)
        ano = st.selectbox("Ano", anos_disponiveis, index=0)
    with c1:
        meses_disponiveis = meses_validos_para_ano(ano)
        mes_padrao = hoje.month if hoje.year == ano and hoje.month in meses_disponiveis else meses_disponiveis[0]
        mes = st.selectbox("Mês", meses_disponiveis, index=meses_disponiveis.index(mes_padrao), format_func=lambda x: MESES_PT[x])
    dm = filter_month(df, ano, mes)

    bruto = float(dm["total"].sum()) if not dm.empty else 0
    reembolso_confirmado = float(dm.loc[dm["status_devolucao"] == "Reembolso confirmado", "valor_devolucao"].sum()) if not dm.empty else 0
    liquido = float(dm["valor_liquido"].sum()) if not dm.empty else 0
    nf_pendentes = int((dm["numero_nf"].fillna("").astype(str).str.strip() == "").sum()) if not dm.empty else 0
    reemb_pend = float(dm.loc[
        dm["houve_devolucao"].fillna(False) & (dm["status_devolucao"] != "Reembolso confirmado"),
        "valor_devolucao"
    ].sum()) if not dm.empty else 0

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Compras brutas", brl(bruto))
    m2.metric("Reembolsos confirmados", brl(reembolso_confirmado))
    m3.metric("Gasto líquido", brl(liquido))
    m4.metric("NFs pendentes", nf_pendentes)
    m5.metric("Reembolso pendente", brl(reemb_pend))

    st.caption(f"Período selecionado: {MESES_PT[mes]} de {ano}. Quando for o mês atual, os valores representam o mês até o momento.")

    left, right = st.columns(2)
    with left:
        st.subheader("Por centro de custo")
        if dm.empty:
            st.info("Nenhuma compra neste período.")
        else:
            cc = dm.groupby("centro_custo", as_index=False)["valor_liquido"].sum().sort_values("valor_liquido", ascending=False)
            cc = cc.rename(columns={"centro_custo": "Centro de custo", "valor_liquido": "Gasto líquido"})
            st.bar_chart(cc.set_index("Centro de custo"))
            st.dataframe(cc.assign(**{"Gasto líquido": cc["Gasto líquido"].map(brl)}), hide_index=True, use_container_width=True)

    with right:
        st.subheader("Por cartão / forma de pagamento")
        if dm.empty:
            st.info("Nenhuma compra neste período.")
        else:
            fp = dm.groupby("forma_pagamento", as_index=False)["valor_liquido"].sum().sort_values("valor_liquido", ascending=False)
            fp = fp.rename(columns={"forma_pagamento": "Cartão / pagamento", "valor_liquido": "Gasto líquido"})
            st.bar_chart(fp.set_index("Cartão / pagamento"))
            st.dataframe(fp.assign(**{"Gasto líquido": fp["Gasto líquido"].map(brl)}), hide_index=True, use_container_width=True)

    st.subheader("Pendências do mês")
    p1, p2 = st.columns(2)
    with p1:
        st.markdown("**🧾 Notas fiscais pendentes**")
        pend_nf = dm[dm["numero_nf"].fillna("").astype(str).str.strip() == ""].copy() if not dm.empty else pd.DataFrame()
        if pend_nf.empty:
            st.success("Nenhuma NF pendente.")
        else:
            view = pend_nf[["data", "item", "centro_custo", "total"]].copy()
            view["data"] = view["data"].dt.strftime("%d/%m/%Y")
            view["total"] = view["total"].map(brl)
            view.columns = ["Data", "Item", "Centro de custo", "Total"]
            st.dataframe(view, hide_index=True, use_container_width=True)

    with p2:
        st.markdown("**↩️ Devoluções / reembolsos pendentes**")
        pend_dev = dm[dm["houve_devolucao"].fillna(False) & (dm["status_devolucao"] != "Reembolso confirmado")].copy() if not dm.empty else pd.DataFrame()
        if pend_dev.empty:
            st.success("Nenhum reembolso pendente.")
        else:
            view = pend_dev[["data", "item", "status_devolucao", "valor_devolucao"]].copy()
            view["data"] = view["data"].dt.strftime("%d/%m/%Y")
            view["valor_devolucao"] = view["valor_devolucao"].map(brl)
            view.columns = ["Data", "Item", "Status", "Valor"]
            st.dataframe(view, hide_index=True, use_container_width=True)

# ==========================================================
# INSERIR COMPRAS
# ==========================================================
elif menu == "🛒 Inserir compras":
    tab_nova, tab_pend_nf, tab_devolucoes, tab_historico = st.tabs([
        "➕ Nova compra", "🧾 NFs pendentes", "↩️ Devoluções", "📋 Compras lançadas"
    ])

    with tab_nova:
        st.subheader("Nova compra")
        with st.form("form_nova_compra", clear_on_submit=True):
            a, b, c = st.columns(3)
            with a:
                sites = db["config"].get("sites", SITES_PADRAO)
                site_sel = st.selectbox("Site da compra *", sites + ["OUTRO / CADASTRAR"])
                site_outro = st.text_input("Novo site", placeholder="Ex.: Amazon") if site_sel == "OUTRO / CADASTRAR" else ""
                centro = st.selectbox("Centro de custo *", db["config"].get("centros_custo", CENTROS_CUSTO_PADRAO))
            with b:
                data_compra = st.date_input("Data da compra *", value=max(date.today(), DATA_INICIO_APP), min_value=DATA_INICIO_APP, format="DD/MM/YYYY")
                numero_nf = st.text_input("Número da NF", placeholder="Pode deixar vazio e completar depois")
            with c:
                pagamentos = db["config"].get("formas_pagamento", FORMAS_PAGAMENTO_PADRAO)
                pag_sel = st.selectbox("Cartão / forma de pagamento *", pagamentos + ["OUTRO / CADASTRAR"])
                pag_outro = st.text_input("Novo cartão / pagamento", placeholder="Ex.: 1234 ou PIX") if pag_sel == "OUTRO / CADASTRAR" else ""
                parcelas = st.selectbox("Parcelamento *", list(range(1, 13)), format_func=lambda x: f"{x}x")

            itens_historicos = sorted({str(c.get("item", "")).strip().upper() for c in db.get("compras", []) if str(c.get("item", "")).strip()})
            sugestao_item = st.selectbox(
                "Sugestão de item já cadastrado (opcional)",
                ["DIGITAR NOVO ITEM"] + itens_historicos,
                index=0,
                help="Use esta lista apenas se quiser reaproveitar um item já cadastrado. Para qualquer item novo, deixe em 'DIGITAR NOVO ITEM' e escreva livremente no campo abaixo.",
            )
            valor_inicial_item = "" if sugestao_item == "DIGITAR NOVO ITEM" else sugestao_item
            item = st.text_input(
                "Item / descrição da compra *",
                value=valor_inicial_item,
                placeholder="Digite livremente o nome do item",
                help="Este campo aceita qualquer descrição nova. O texto digitado aqui é o que será salvo.",
            )
            total = st.number_input("Valor total da compra (R$) *", min_value=0.0, step=0.01, format="%.2f")

            houve = st.radio("Houve devolução?", ["Não", "Sim"], horizontal=True)
            valor_dev = 0.0
            status_dev = ""
            if houve == "Sim":
                d1, d2 = st.columns(2)
                with d1:
                    valor_dev = st.number_input("Valor da devolução / reembolso (R$) *", min_value=0.0, max_value=float(total) if total > 0 else None, step=0.01, format="%.2f")
                with d2:
                    status_dev = st.selectbox("Status da devolução *", STATUS_DEVOLUCAO)

            observacao = st.text_area("Observação", height=80)
            salvar = st.form_submit_button("💾 Salvar compra", use_container_width=True, type="primary")

        if salvar:
            site = (site_outro if site_sel == "OUTRO / CADASTRAR" else site_sel).strip().upper()
            pagamento = (pag_outro if pag_sel == "OUTRO / CADASTRAR" else pag_sel).strip().upper()
            erros = []
            if not site: erros.append("Informe o site da compra.")
            if not item.strip(): erros.append("Informe o item.")
            if not pagamento: erros.append("Informe o cartão / forma de pagamento.")
            if total <= 0: erros.append("Informe um valor de compra maior que zero.")
            if houve == "Sim" and valor_dev <= 0: erros.append("Informe o valor da devolução.")
            if houve == "Sim" and valor_dev > total: erros.append("A devolução não pode ser maior que o valor da compra.")

            if erros:
                for e in erros:
                    st.error(e)
            else:
                add_option("sites", site)
                add_option("formas_pagamento", pagamento)
                now = datetime.now().isoformat(timespec="seconds")
                compra = {
                    "id": next_id(),
                    "site": site,
                    "centro_custo": centro,
                    "data_compra": data_compra.isoformat(),
                    "numero_nf": numero_nf.strip(),
                    "status_nf": "Recebida" if numero_nf.strip() else "Pendente",
                    "motivo_nf": "",
                    "item": item.strip().upper(),
                    "forma_pagamento": pagamento,
                    "parcelamento": int(parcelas),
                    "total": float(total),
                    "houve_devolucao": houve == "Sim",
                    "valor_devolucao": float(valor_dev if houve == "Sim" else 0),
                    "status_devolucao": status_dev if houve == "Sim" else "",
                    "observacao": observacao.strip(),
                    "criado_em": now,
                    "atualizado_em": now,
                    "usuario_criacao": st.session_state.get("usuario", ""),
                    "usuario_atualizacao": st.session_state.get("usuario", ""),
                }
                db["compras"].append(compra)
                persist(db)
                st.info(f"Gasto líquido atual desta compra: {brl(compra_liquida(compra))}")

    with tab_pend_nf:
        st.subheader("Notas fiscais pendentes")
        hoje_nf = date.today()
        filtro_nf = st.radio("Exibir", ["Mês atual", "Todas as pendências"], horizontal=True, key="filtro_nf")
        pendentes = [c for c in db["compras"] if nf_esta_pendente(c)]
        if filtro_nf == "Mês atual":
            pendentes = [
                c for c in pendentes
                if pd.to_datetime(c.get("data_compra")).year == hoje_nf.year
                and pd.to_datetime(c.get("data_compra")).month == hoje_nf.month
            ]

        if not pendentes:
            st.success("Nenhuma nota fiscal pendente neste filtro.")
        else:
            pendentes = sorted(pendentes, key=lambda x: x.get("data_compra", ""), reverse=True)
            opcoes = {f"{pd.to_datetime(c['data_compra']).strftime('%d/%m/%Y')} • {c['item']} • {c['centro_custo']} • {brl(c['total'])}": c for c in pendentes}
            escolha = st.selectbox("Selecione a compra", list(opcoes.keys()))
            c = opcoes[escolha]
            st.caption(f"Cartão: {c['forma_pagamento']} | Site: {c['site']}")
            nf = st.text_input("Número da NF recebida", key=f"nf_{c['id']}")
            b_nf1, b_nf2 = st.columns(2)
            if b_nf1.button("✅ Salvar NF", type="primary", use_container_width=True):
                if not nf.strip():
                    st.error("Informe o número da NF.")
                else:
                    c["numero_nf"] = nf.strip()
                    c["status_nf"] = "Recebida"
                    c["motivo_nf"] = ""
                    c["atualizado_em"] = datetime.now().isoformat(timespec="seconds")
                    c["usuario_atualizacao"] = st.session_state.get("usuario", "")
                    persist(db)
                    st.rerun()

            with b_nf2.popover("❌ Finalizar sem NF", use_container_width=True):
                st.caption("Use quando a NF foi cobrada e não houve retorno. No relatório, o campo NF ficará como N/C.")
                motivo = st.selectbox("Motivo", ["Falta de retorno", "Fornecedor não enviou", "Outro"], key=f"motivo_nf_{c['id']}")
                detalhe = st.text_input("Detalhe / observação", key=f"detalhe_nf_{c['id']}") if motivo == "Outro" else ""
                if st.button("Confirmar finalização", key=f"finalizar_nf_{c['id']}", type="primary", use_container_width=True):
                    c["numero_nf"] = "N/C"
                    c["status_nf"] = "Finalizada sem NF"
                    c["motivo_nf"] = detalhe.strip() if motivo == "Outro" and detalhe.strip() else motivo
                    c["atualizado_em"] = datetime.now().isoformat(timespec="seconds")
                    c["usuario_atualizacao"] = st.session_state.get("usuario", "")
                    persist(db)
                    st.rerun()

    with tab_devolucoes:
        st.subheader("Acompanhamento de devoluções")

        # Permite registrar uma devolução depois que a compra já foi lançada.
        # Isso é útil para compras recentes, quando a necessidade de devolução ainda não era conhecida no cadastro inicial.
        with st.expander("➕ Registrar devolução de uma compra já lançada", expanded=False):
            compras_sem_dev = [c for c in db["compras"] if not c.get("houve_devolucao")]
            if not compras_sem_dev:
                st.info("Todas as compras cadastradas já possuem informação de devolução.")
            else:
                compras_sem_dev = sorted(compras_sem_dev, key=lambda x: x.get("data_compra", ""), reverse=True)
                opcoes_nova_dev = {
                    f"{pd.to_datetime(c['data_compra']).strftime('%d/%m/%Y')} • {c['item']} • {c['centro_custo']} • {brl(c['total'])}": c
                    for c in compras_sem_dev
                }
                escolha_nova_dev = st.selectbox("Selecione a compra", list(opcoes_nova_dev.keys()), key="nova_dev_compra")
                compra_dev = opcoes_nova_dev[escolha_nova_dev]
                nd1, nd2 = st.columns(2)
                with nd1:
                    valor_nova_dev = st.number_input(
                        "Valor da devolução / reembolso (R$)",
                        min_value=0.0,
                        max_value=float(compra_dev["total"]),
                        step=0.01,
                        format="%.2f",
                        key=f"valor_nova_dev_{compra_dev['id']}",
                    )
                with nd2:
                    status_nova_dev = st.selectbox(
                        "Status da devolução",
                        STATUS_DEVOLUCAO,
                        index=0,
                        key=f"status_nova_dev_{compra_dev['id']}",
                    )
                if st.button("💾 Registrar devolução", type="primary", use_container_width=True, key=f"registrar_dev_{compra_dev['id']}"):
                    if valor_nova_dev <= 0:
                        st.error("Informe o valor da devolução / reembolso.")
                    else:
                        compra_dev["houve_devolucao"] = True
                        compra_dev["valor_devolucao"] = float(valor_nova_dev)
                        compra_dev["status_devolucao"] = status_nova_dev
                        compra_dev["atualizado_em"] = datetime.now().isoformat(timespec="seconds")
                        compra_dev["usuario_atualizacao"] = st.session_state.get("usuario", "")
                        persist(db)
                        st.rerun()

        st.markdown("### Pendências em acompanhamento")
        # Reembolso confirmado é etapa concluída: permanece no histórico e nos cálculos,
        # mas não aparece mais na fila de acompanhamento.
        devs = [
            c for c in db["compras"]
            if c.get("houve_devolucao") and c.get("status_devolucao") != "Reembolso confirmado"
        ]
        if not devs:
            st.success("Nenhuma devolução pendente de acompanhamento.")
        else:
            devs = sorted(devs, key=lambda x: x.get("data_compra", ""), reverse=True)
            opcoes = {f"{pd.to_datetime(c['data_compra']).strftime('%d/%m/%Y')} • {c['item']} • {brl(c['valor_devolucao'])} • {c.get('status_devolucao','')}": c for c in devs}
            escolha = st.selectbox("Selecione a devolução pendente", list(opcoes.keys()))
            c = opcoes[escolha]
            col1, col2 = st.columns(2)
            with col1:
                novo_valor = st.number_input("Valor da devolução", min_value=0.0, max_value=float(c["total"]), value=float(c.get("valor_devolucao", 0)), step=0.01, format="%.2f")
            with col2:
                atual = c.get("status_devolucao") or STATUS_DEVOLUCAO[0]
                idx = STATUS_DEVOLUCAO.index(atual) if atual in STATUS_DEVOLUCAO else 0
                novo_status = st.selectbox("Status", STATUS_DEVOLUCAO, index=idx)
            st.info(f"Valor da compra: {brl(c['total'])} | Gasto líquido se confirmado: {brl(max(0, c['total']-novo_valor))}")
            if st.button("💾 Atualizar devolução", type="primary"):
                c["valor_devolucao"] = float(novo_valor)
                c["status_devolucao"] = novo_status
                c["atualizado_em"] = datetime.now().isoformat(timespec="seconds")
                c["usuario_atualizacao"] = st.session_state.get("usuario", "")
                persist(db)
                st.rerun()

    with tab_historico:
        st.subheader("Compras lançadas")
        df = compras_df()
        if df.empty:
            st.info("Nenhuma compra lançada.")
        else:
            f1, f2, f3 = st.columns(3)
            with f1:
                anos = [a for a in sorted(df["data"].dt.year.dropna().unique().astype(int), reverse=True) if a >= 2026]
                ano_f = st.selectbox("Ano", anos, key="hist_ano")
            with f2:
                meses_hist = meses_validos_para_ano(ano_f)
                mes_padrao_hist = date.today().month if date.today().year == ano_f and date.today().month in meses_hist else meses_hist[0]
                mes_f = st.selectbox("Mês", meses_hist, index=meses_hist.index(mes_padrao_hist), format_func=lambda x: MESES_PT[x], key="hist_mes")
            with f3:
                cc_f = st.selectbox("Centro de custo", ["TODOS"] + db["config"].get("centros_custo", CENTROS_CUSTO_PADRAO), key="hist_cc")
            d = filter_month(df, ano_f, mes_f)
            if cc_f != "TODOS":
                d = d[d["centro_custo"] == cc_f]
            if d.empty:
                st.info("Nenhuma compra para os filtros selecionados.")
            else:
                view = d.sort_values("data", ascending=False).copy()
                view["Data"] = view["data"].dt.strftime("%d/%m/%Y")
                view["Total"] = view["total"].map(brl)
                view["Devolução"] = view["valor_devolucao"].map(brl)
                view["Gasto líquido"] = view["valor_liquido"].map(brl)
                view["NF"] = view["numero_nf"].replace("", "PENDENTE")
                view["Status NF"] = view.get("status_nf", pd.Series(index=view.index, dtype=str)).replace("", "-")
                view["Status"] = view["status_devolucao"].replace("", "-")
                st.dataframe(view[["Data","site","centro_custo","NF","Status NF","item","forma_pagamento","parcelamento","Total","Devolução","Status","Gasto líquido"]].rename(columns={
                    "site":"Site", "centro_custo":"Centro de custo", "item":"Item", "forma_pagamento":"Pagamento", "parcelamento":"Parcelas"
                }), hide_index=True, use_container_width=True)

                ids = d.sort_values("data", ascending=False)["id"].tolist()
                labels = {c["id"]: f"{pd.to_datetime(c['data_compra']).strftime('%d/%m/%Y')} • {c['item']} • {brl(c['total'])}" for c in db["compras"] if c["id"] in ids}
                with st.expander("✏️ Editar ou excluir lançamento"):
                    eid = st.selectbox("Compra", ids, format_func=lambda x: labels[x])
                    c = next(x for x in db["compras"] if x["id"] == eid)

                    st.caption("Todos os dados abaixo podem ser corrigidos. Salvar altera o mesmo lançamento, sem criar duplicidade.")

                    e1, e2, e3 = st.columns(3)
                    with e1:
                        data_atual = pd.to_datetime(c.get("data_compra"), errors="coerce")
                        data_atual = data_atual.date() if pd.notna(data_atual) else date.today()
                        e_data = st.date_input(
                            "Data da compra",
                            value=max(data_atual, DATA_INICIO_APP),
                            min_value=DATA_INICIO_APP,
                            format="DD/MM/YYYY",
                            key=f"edit_data_{eid}",
                        )
                        e_site = st.text_input("Site / fornecedor", value=c.get("site", ""), key=f"edit_site_{eid}")
                        centros = db["config"].get("centros_custo", CENTROS_CUSTO_PADRAO)
                        e_cc = st.selectbox(
                            "Centro de custo",
                            centros,
                            index=centros.index(c.get("centro_custo")) if c.get("centro_custo") in centros else 0,
                            key=f"edit_cc_{eid}",
                        )
                    with e2:
                        e_nf = st.text_input("Número da NF", value=c.get("numero_nf", ""), key=f"edit_nf_{eid}")
                        e_item = st.text_input("Item / descrição", value=c.get("item", ""), key=f"edit_item_{eid}")
                        e_pag = st.text_input("Cartão / forma de pagamento", value=c.get("forma_pagamento", ""), key=f"edit_pag_{eid}")
                    with e3:
                        e_parc = st.selectbox(
                            "Parcelamento",
                            list(range(1, 13)),
                            index=max(0, min(11, int(c.get("parcelamento", 1) or 1) - 1)),
                            format_func=lambda x: f"{x}x",
                            key=f"edit_parc_{eid}",
                        )
                        e_total = st.number_input(
                            "Valor total da compra (R$)",
                            min_value=0.01,
                            value=float(c.get("total", 0) or 0),
                            step=0.01,
                            format="%.2f",
                            key=f"edit_total_{eid}",
                        )
                        e_motivo_nf = st.text_input(
                            "Motivo / observação da NF",
                            value=c.get("motivo_nf", ""),
                            placeholder="Ex.: Falta de retorno",
                            key=f"edit_motivo_nf_{eid}",
                        )

                    e_obs = st.text_area(
                        "Observação da compra",
                        value=c.get("observacao", ""),
                        height=90,
                        key=f"edit_obs_{eid}",
                    )

                    houve_atual = bool(c.get("houve_devolucao", False))
                    e_houve = st.radio(
                        "Houve devolução?",
                        ["Não", "Sim"],
                        index=1 if houve_atual else 0,
                        horizontal=True,
                        key=f"edit_houve_{eid}",
                    )

                    e_valor_dev = 0.0
                    e_status_dev = ""
                    if e_houve == "Sim":
                        d1, d2 = st.columns(2)
                        with d1:
                            valor_dev_atual = min(float(c.get("valor_devolucao", 0) or 0), float(e_total))
                            e_valor_dev = st.number_input(
                                "Valor da devolução / reembolso (R$)",
                                min_value=0.0,
                                max_value=float(e_total),
                                value=valor_dev_atual,
                                step=0.01,
                                format="%.2f",
                                key=f"edit_valor_dev_{eid}",
                            )
                        with d2:
                            status_atual = c.get("status_devolucao", "")
                            idx_status = STATUS_DEVOLUCAO.index(status_atual) if status_atual in STATUS_DEVOLUCAO else 0
                            e_status_dev = st.selectbox(
                                "Status da devolução",
                                STATUS_DEVOLUCAO,
                                index=idx_status,
                                key=f"edit_status_dev_{eid}",
                            )

                    b1, b2 = st.columns(2)
                    if b1.button("💾 Salvar alterações", use_container_width=True, type="primary"):
                        erros_edicao = []
                        if not e_site.strip():
                            erros_edicao.append("Informe o site / fornecedor.")
                        if not e_item.strip():
                            erros_edicao.append("Informe o item / descrição.")
                        if not e_pag.strip():
                            erros_edicao.append("Informe o cartão / forma de pagamento.")
                        if e_houve == "Sim" and e_valor_dev <= 0:
                            erros_edicao.append("Informe o valor da devolução.")
                        if e_houve == "Sim" and e_valor_dev > e_total:
                            erros_edicao.append("A devolução não pode ser maior que o valor da compra.")

                        if erros_edicao:
                            for erro in erros_edicao:
                                st.error(erro)
                        else:
                            nf_limpa = e_nf.strip()
                            if nf_limpa.upper() == "N/C":
                                status_nf_edit = "Finalizada sem NF"
                            elif nf_limpa:
                                status_nf_edit = "Recebida"
                            else:
                                status_nf_edit = "Pendente"

                            site_edit = e_site.strip().upper()
                            pag_edit = e_pag.strip().upper()
                            add_option("sites", site_edit)
                            add_option("formas_pagamento", pag_edit)

                            c.update({
                                "data_compra": e_data.isoformat(),
                                "site": site_edit,
                                "numero_nf": nf_limpa,
                                "status_nf": status_nf_edit,
                                "motivo_nf": e_motivo_nf.strip(),
                                "item": e_item.strip().upper(),
                                "total": float(e_total),
                                "centro_custo": e_cc,
                                "forma_pagamento": pag_edit,
                                "parcelamento": int(e_parc),
                                "houve_devolucao": e_houve == "Sim",
                                "valor_devolucao": float(e_valor_dev if e_houve == "Sim" else 0),
                                "status_devolucao": e_status_dev if e_houve == "Sim" else "",
                                "observacao": e_obs.strip(),
                                "atualizado_em": datetime.now().isoformat(timespec="seconds"),
                                "usuario_atualizacao": st.session_state.get("usuario", ""),
                            })
                            persist(db)
                            st.rerun()

                    confirmar_exclusao = b2.checkbox(
                        "Confirmar exclusão",
                        key=f"confirm_delete_{eid}",
                        help="Marque esta opção para liberar o botão de exclusão.",
                    )
                    if st.button(
                        "🗑️ Excluir compra definitivamente",
                        use_container_width=True,
                        disabled=not confirmar_exclusao,
                        key=f"delete_{eid}",
                    ):
                        db["compras"] = [x for x in db["compras"] if x["id"] != eid]
                        persist(db)
                        st.rerun()

# ==========================================================
# EXPORTAÇÃO
# ==========================================================
elif menu == "📥 Exportar relatório":
    st.subheader("Fechamento mensal")
    df = compras_df()
    hoje = date.today()
    e1, e2 = st.columns(2)
    with e2:
        anos = anos_validos(df)
        ano = st.selectbox("Ano do relatório", anos)
    with e1:
        meses_relatorio = meses_validos_para_ano(ano)
        mes_padrao_rel = hoje.month if hoje.year == ano and hoje.month in meses_relatorio else meses_relatorio[0]
        mes = st.selectbox("Mês do relatório", meses_relatorio, index=meses_relatorio.index(mes_padrao_rel), format_func=lambda x: MESES_PT[x])

    dm = filter_month(df, ano, mes)
    bruto = float(dm["total"].sum()) if not dm.empty else 0
    reembolso = float(dm.loc[dm["status_devolucao"] == "Reembolso confirmado", "valor_devolucao"].sum()) if not dm.empty else 0
    liquido = float(dm["valor_liquido"].sum()) if not dm.empty else 0
    nf_pend = int((dm["numero_nf"].fillna("").astype(str).str.strip()=="").sum()) if not dm.empty else 0
    reemb_pend = float(dm.loc[dm["houve_devolucao"].fillna(False) & (dm["status_devolucao"] != "Reembolso confirmado"), "valor_devolucao"].sum()) if not dm.empty else 0

    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Compras brutas", brl(bruto))
    m2.metric("Reembolsos", brl(reembolso))
    m3.metric("Gasto líquido", brl(liquido))
    m4.metric("NFs pendentes", nf_pend)
    m5.metric("Reembolso pendente", brl(reemb_pend))

    def gerar_excel(dm, ano, mes):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            workbook = writer.book
            ws = workbook.add_worksheet("Relatório")
            writer.sheets["Relatório"] = ws

            fmt_title = workbook.add_format({"bold":True,"font_size":16,"align":"center","valign":"vcenter"})
            fmt_header = workbook.add_format({"bold":True,"bg_color":"#1F4E78","font_color":"#FFFFFF","border":1,"align":"center","valign":"vcenter","text_wrap":True})
            fmt_text = workbook.add_format({"border":1,"valign":"top"})
            fmt_date = workbook.add_format({"border":1,"num_format":"dd/mm/yyyy","align":"center"})
            fmt_money = workbook.add_format({"border":1,"num_format":"R$ #,##0.00"})
            fmt_center = workbook.add_format({"border":1,"align":"center"})
            fmt_total = workbook.add_format({"bold":True,"top":1,"num_format":"R$ #,##0.00"})

            ws.merge_range("A1:L1", f"COMPRAS - GDS LOGÍSTICA - {MESES_PT[mes].upper()} {ano}", fmt_title)
            headers = ["SITE DA COMPRA","CENTRO DE CUSTO","DATA DA COMPRA","NUMERO NF","ITEM","FORMA DE PAGAMENTO","PARCELAMENTO","TOTAL","DEVOLUÇÃO","HOUVE DEVOLUÇÃO?","STATUS DEVOLUÇÃO","VALOR LÍQUIDO"]
            for col,h in enumerate(headers): ws.write(2,col,h,fmt_header)

            row = 3
            for _,r in dm.sort_values("data").iterrows():
                ws.write(row,0,r.get("site",""),fmt_text)
                ws.write(row,1,r.get("centro_custo",""),fmt_center)
                ws.write_datetime(row,2,r["data"].to_pydatetime(),fmt_date)
                ws.write(row,3,r.get("numero_nf","") or "",fmt_text)
                ws.write(row,4,r.get("item",""),fmt_text)
                ws.write(row,5,r.get("forma_pagamento",""),fmt_center)
                ws.write(row,6,int(r.get("parcelamento",1)),fmt_center)
                ws.write_number(row,7,float(r.get("total",0)),fmt_money)
                ws.write_number(row,8,float(r.get("valor_devolucao",0)),fmt_money)
                ws.write(row,9,"SIM" if r.get("houve_devolucao") else "NÃO",fmt_center)
                ws.write(row,10,r.get("status_devolucao","") or "",fmt_text)
                ws.write_number(row,11,float(r.get("valor_liquido",0)),fmt_money)
                row += 1
            ws.write(row+1,6,"VALOR BRUTO",fmt_total); ws.write_number(row+1,7,bruto,fmt_total)
            ws.write(row+2,6,"REEMBOLSOS CONFIRMADOS",fmt_total); ws.write_number(row+2,7,reembolso,fmt_total)
            ws.write(row+3,6,"VALOR LÍQUIDO",fmt_total); ws.write_number(row+3,7,liquido,fmt_total)
            ws.set_column("A:A",18); ws.set_column("B:B",16); ws.set_column("C:C",14); ws.set_column("D:D",16)
            ws.set_column("E:E",35); ws.set_column("F:G",18); ws.set_column("H:I",15); ws.set_column("J:K",21); ws.set_column("L:L",16)
            ws.freeze_panes(3,0)
            ws.set_landscape(); ws.fit_to_pages(1,0); ws.set_margins(0.25,0.25,0.5,0.5)
        return output.getvalue()

    def gerar_pdf(dm, ano, mes):
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=16*mm, leftMargin=16*mm, topMargin=14*mm, bottomMargin=14*mm)
        styles = getSampleStyleSheet()
        title = ParagraphStyle("TitleX", parent=styles["Title"], alignment=TA_CENTER, fontSize=19, leading=23, spaceAfter=4)
        subtitle = ParagraphStyle("SubX", parent=styles["Normal"], alignment=TA_CENTER, fontSize=10, textColor=colors.HexColor("#5F6368"), spaceAfter=14)
        sec = ParagraphStyle("Sec", parent=styles["Heading2"], fontSize=12, leading=15, spaceBefore=8, spaceAfter=7, textColor=colors.HexColor("#1F4E78"))
        small = ParagraphStyle("Small", parent=styles["Normal"], fontSize=8.5, leading=11)

        story = [
            Paragraph("GDS LOGÍSTICA", title),
            Paragraph("RELATÓRIO MENSAL DE COMPRAS", title),
            Paragraph(f"{MESES_PT[mes].upper()} / {ano}", title),
            Paragraph(f"Emitido em {datetime.now().strftime('%d/%m/%Y às %H:%M')}", subtitle),
        ]

        resumo = [
            ["INDICADOR", "VALOR"],
            ["Compras realizadas", str(len(dm))],
            ["Valor bruto comprado", brl(bruto)],
            ["Reembolsos confirmados", brl(reembolso)],
            ["Gasto líquido do mês", brl(liquido)],
            ["NFs pendentes", str(nf_pend)],
            ["Reembolsos pendentes", brl(reemb_pend)],
        ]
        t = Table(resumo, colWidths=[120*mm, 50*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),
            ("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTNAME",(0,3),(-1,3),"Helvetica-Bold"),
            ("ALIGN",(1,0),(1,-1),"RIGHT"),
            ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#D1D5DB")),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#F8FAFC")]),
            ("BOTTOMPADDING",(0,0),(-1,-1),7), ("TOPPADDING",(0,0),(-1,-1),7),
        ]))
        story += [Paragraph("Resumo do mês", sec), t, Spacer(1, 8*mm)]

        if not dm.empty:
            cc = dm.groupby("centro_custo")["valor_liquido"].sum().sort_values(ascending=False)
            cc_data = [["CENTRO DE CUSTO", "GASTO LÍQUIDO"]] + [[str(k), brl(v)] for k,v in cc.items()]
            tc = Table(cc_data, colWidths=[100*mm,70*mm])
            tc.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E8EEF5")),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("ALIGN",(1,0),(1,-1),"RIGHT"),
                ("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#D1D5DB")),
                ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
            ]))
            story += [Paragraph("Gastos por centro de custo", sec), tc, Spacer(1,5*mm)]

            fp = dm.groupby("forma_pagamento")["valor_liquido"].sum().sort_values(ascending=False)
            fp_data = [["CARTÃO / PAGAMENTO", "GASTO LÍQUIDO"]] + [[str(k), brl(v)] for k,v in fp.items()]
            tf = Table(fp_data, colWidths=[100*mm,70*mm])
            tf.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E8EEF5")),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("ALIGN",(1,0),(1,-1),"RIGHT"),
                ("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#D1D5DB")),
                ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
            ]))
            story += [Paragraph("Gastos por cartão / forma de pagamento", sec), tf, Spacer(1,5*mm)]

        pendencias = []
        if nf_pend:
            pendencias.append(f"• {nf_pend} compra(s) aguardando Nota Fiscal.")
        pend_dev_count = int((dm["houve_devolucao"].fillna(False) & (dm["status_devolucao"] != "Reembolso confirmado")).sum()) if not dm.empty else 0
        if pend_dev_count:
            pendencias.append(f"• {pend_dev_count} devolução(ões) ainda não concluída(s), totalizando {brl(reemb_pend)}.")
        if not pendencias:
            pendencias.append("• Nenhuma pendência registrada no fechamento deste período.")
        story += [Paragraph("Pendências no fechamento", sec), Paragraph("<br/>".join(pendencias), small)]

        doc.build(story)
        return buffer.getvalue()

    if dm.empty:
        st.info("Nenhuma compra cadastrada para este mês. Os arquivos podem ser gerados após existirem lançamentos.")
    else:
        st.markdown("### Arquivos do fechamento")
        excel = gerar_excel(dm, ano, mes)
        pdf = gerar_pdf(dm, ano, mes)
        b1,b2 = st.columns(2)
        with b1:
            st.download_button(
                "📊 Baixar relatório Excel",
                data=excel,
                file_name=f"COMPRAS - GDSLOGISTICA - {MESES_PT[mes].upper()} {ano}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
        with b2:
            st.download_button(
                "📄 Baixar capa PDF",
                data=pdf,
                file_name=f"CAPA - RELATORIO DE COMPRAS - {MESES_PT[mes].upper()} {ano}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        st.markdown("### Prévia das compras")
        view = dm.sort_values("data").copy()
        view["Data"] = view["data"].dt.strftime("%d/%m/%Y")
        view["Total"] = view["total"].map(brl)
        view["Devolução"] = view["valor_devolucao"].map(brl)
        view["Gasto líquido"] = view["valor_liquido"].map(brl)
        view["NF"] = view["numero_nf"].replace("", "PENDENTE")
        st.dataframe(view[["Data","centro_custo","NF","item","forma_pagamento","Total","Devolução","status_devolucao","Gasto líquido"]].rename(columns={
            "centro_custo":"Centro de custo","item":"Item","forma_pagamento":"Pagamento","status_devolucao":"Status devolução"
        }), hide_index=True, use_container_width=True)

st.divider()
st.caption("GDS Logística • Controle mensal de compras")
