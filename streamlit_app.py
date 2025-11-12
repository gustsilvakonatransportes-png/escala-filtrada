# streamlit_app.py
import streamlit as st
from io import BytesIO
from datetime import datetime
import pandas as pd
import re
from openpyxl import load_workbook

st.set_page_config(page_title="ESCALA FILTRADA", layout="wide")
st.title("🚛 Escala Filtrada — Extrator automático de blocos - José Cristiano")
st.markdown("Envie a planilha com os blocos; escolha o turno; baixe a planilha padronizada.")

turno = st.selectbox("Escolha o turno:", ["Noturno", "Diurno"])
uploaded_file = st.file_uploader("Enviar arquivo .xlsx (escala)", type=["xlsx"])

# padrões
FROTA_RE = re.compile(r"\bT\d{2,4}\b", re.IGNORECASE)
PLACA_RE = re.compile(r"[A-Z0-9]{5,8}", re.IGNORECASE)
ROTA_RE = re.compile(r"\b\d{4,5}\b")           # rota = 4-5 dígitos (ajuste se precisar)
LARGADA_KEY_RE = re.compile(r"LARGAD|LARGA|LARGADA", re.IGNORECASE)
MOTORISTA_KEY = re.compile(r"motorista", re.IGNORECASE)
AJ1_KEY = re.compile(r"ajudante\s*1|aj1", re.IGNORECASE)
AJ2_KEY = re.compile(r"ajudante\s*2|aj2", re.IGNORECASE)

def cell_text(cell):
    return "" if cell is None else str(cell).strip()

def find_all_cells(ws):
    """Retorna lista de (r,c,text) para todas as células não vazias"""
    cells = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = cell_text(ws.cell(row=r, column=c).value)
            if v:
                cells.append((r, c, v))
    return cells

def find_frota_positions(ws):
    """Encontra células com frota (T###) e tenta pegar placa da direita"""
    positions = []
    for r, c, v in find_all_cells(ws):
        if FROTA_RE.search(v):
            placa = ""
            # procurar placa nas 4 células seguintes na mesma linha
            for cc in range(c+1, min(c+4, ws.max_column) + 1):
                txt = cell_text(ws.cell(row=r, column=cc).value)
                if txt and any(ch.isalpha() for ch in txt) and any(ch.isdigit() for ch in txt):
                    placa = txt
                    break
            positions.append((r, c, v, placa))
    return positions

def nearest_token_in_block(ws, start_row, end_row, start_col, end_col, pattern):
    """Retorna primeiro token que casa com pattern no bloco (varre em leitura)."""
    for r in range(start_row, end_row+1):
        for c in range(start_col, end_col+1):
            txt = cell_text(ws.cell(row=r, column=c).value)
            if txt:
                m = pattern.search(txt)
                if m:
                    token = m.group(0)
                    return token, r, c
    return None, None, None

def extract_names_from_titles(ws, start_row, end_row, start_col, end_col):
    """Procura por células com 'MOTORISTA','AJUDANTE 1','AJUDANTE 2','LARGADA' e pega valor abaixo (ou próximo)."""
    motorista = ajud1 = ajud2 = largada = ""
    for r in range(start_row, end_row+1):
        for c in range(start_col, end_col+1):
            v = cell_text(ws.cell(row=r, column=c).value)
            if not v:
                continue
            if MOTORISTA_KEY.search(v) and not motorista:
                # procura imediatamente abaixo (até 4 linhas)
                for rr in range(r+1, min(r+4, end_row)+1):
                    cand = cell_text(ws.cell(row=rr, column=c).value)
                    if cand:
                        motorista = cand
                        break
            if AJ1_KEY.search(v) and not ajud1:
                for rr in range(r+1, min(r+4, end_row)+1):
                    cand = cell_text(ws.cell(row=rr, column=c).value)
                    if cand:
                        ajud1 = cand
                        break
            if AJ2_KEY.search(v) and not ajud2:
                for rr in range(r+1, min(r+4, end_row)+1):
                    cand = cell_text(ws.cell(row=rr, column=c).value)
                    if cand:
                        ajud2 = cand
                        break
            if LARGADA_KEY_RE.search(v) and not largada:
                # pode ter hora na mesma célula
                m = re.search(r"(\d{1,2}[:h]\d{2})", v)
                if m:
                    largada = m.group(0).replace("h", ":")
                else:
                    for rr in range(r+1, min(r+4, end_row)+1):
                        cand = cell_text(ws.cell(row=rr, column=c).value)
                        if cand:
                            m2 = re.search(r"(\d{1,2}[:h]\d{2})", cand)
                            if m2:
                                largada = m2.group(0).replace("h", ":")
                            else:
                                largada = cand
                            break
    return motorista, ajud1, ajud2, largada

def extract_block(ws, start_row, end_row):
    """Extrai dados de um bloco definido por linhas (colunas = full width)."""
    # procura frota na linha start_row (ou acima)
    frota = placa = rota = motorista = ajud1 = ajud2 = largada = ""
    max_col = ws.max_column
    # 1) frota/placa: buscar na linha start_row e nas 2 linhas acima (caso deslocado)
    for rr in range(max(1, start_row-2), start_row+1):
        for c in range(1, max_col+1):
            v = cell_text(ws.cell(row=rr, column=c).value)
            if FROTA_RE.search(v):
                frota = v
                # placa à direita
                for cc in range(c+1, min(c+4, max_col)+1):
                    p = cell_text(ws.cell(row=rr, column=cc).value)
                    if p and any(ch.isalpha() for ch in p) and any(ch.isdigit() for ch in p):
                        placa = p
                        break
                break
        if frota:
            break

    # 2) rota: buscar primeiro token 4-5 dígitos dentro do bloco (evitar CPFs de 11 dígitos)
    token, rr_found, cc_found = nearest_token_in_block(ws, start_row, end_row, 1, max_col, ROTA_RE)
    if token and len(token) <= 5:
        rota = token.zfill(5)

    # 3) nomes: achar titulos MOTORISTA/AJUDANTE dentro do bloco
    motorista, ajud1, ajud2, largada = extract_names_from_titles(ws, start_row, end_row, 1, max_col)

    # 4) heurística: se não encontrou motorista via título, procurar linhas com nomes (maiúsculas longas)
    if not motorista:
        candidates = []
        for r in range(start_row, min(end_row, start_row + 12) + 1):
            row_text = " ".join([cell_text(ws.cell(row=r, column=c).value) for c in range(1, max_col+1)])
            if row_text and len(row_text) > 6 and row_text.upper() == row_text and any(ch.isalpha() for ch in row_text):
                candidates.append(row_text.strip())
        if candidates:
            motorista = motorista or (candidates[0] if len(candidates) >= 1 else "")
            ajud1 = ajud1 or (candidates[1] if len(candidates) >= 2 else "")
            ajud2 = ajud2 or (candidates[2] if len(candidates) >= 3 else "")

    return {
        "Frota": (frota or "").strip(),
        "Placa": (placa or "").strip(),
        "Rota": (rota or "").strip(),
        "Motorista": (motorista or "").strip(),
        "Ajudante 1": (ajud1 or "").strip(),
        "Ajudante 2": (ajud2 or "").strip(),
        "Largada": (largada or "").strip()
    }

def parse_workbook_bytes(file_bytes):
    wb = load_workbook(filename=BytesIO(file_bytes), data_only=True)
    ws = wb.active

    # 1) encontrar possíveis linhas de frota T###
    frota_positions = find_frota_positions(ws)
    blocks = []
    if frota_positions:
        # usa cada posição de frota como início de bloco até linha anterior da próxima frota
        frota_positions = sorted(frota_positions, key=lambda x: x[0])
        for i, pos in enumerate(frota_positions):
            start_row = pos[0]
            if i + 1 < len(frota_positions):
                end_row = frota_positions[i+1][0] - 1
            else:
                end_row = ws.max_row
            blocks.append((start_row, end_row))
    else:
        # fallback: dividir em blocos por grandes áreas vazias (simples)
        blocks.append((1, ws.max_row))

    results = []
    for start_row, end_row in blocks:
        data = extract_block(ws, start_row, end_row)
        # só adiciona se tiver ao menos rota, motorista ou placa/frota
        if any(data[k] for k in ("Frota", "Placa", "Rota", "Motorista")):
            results.append(data)

    df = pd.DataFrame(results, columns=["Frota", "Placa", "Rota", "Motorista", "Ajudante 1", "Ajudante 2", "Largada"])
    return df

# ---- main ----
if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        df_out = parse_workbook_bytes(file_bytes)

        if df_out.empty:
            st.error("Não foi possível detectar blocos. Verifique o arquivo.")
        else:
            df_out["Turno"] = turno
            df_out = df_out[["Frota", "Placa", "Rota", "Motorista", "Ajudante 1", "Ajudante 2", "Turno", "Largada"]]
            # limpa "nan" literais convertidos
            df_out = df_out.replace({"nan": ""})

            st.success(f"✅ {len(df_out)} blocos encontrados!")
            st.dataframe(df_out)

            # exporta para excel em memória e oferece download
            buf = BytesIO()
            data_hoje = datetime.now().strftime("%d-%m-%Y")
            nome_arquivo = f"ESCALA FILTRADA_{data_hoje}.xlsx"
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df_out.to_excel(writer, index=False, sheet_name="Escala Filtrada")
            st.download_button("📥 Baixar ESCALA FILTRADA", buf.getvalue(), file_name=nome_arquivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"❌ Erro ao processar a planilha: {e}")
