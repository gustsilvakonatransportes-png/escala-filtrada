import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Escala Filtrada — Konatransportes", page_icon="🚛", layout="wide")

st.title("🚛 Escala Filtrada — Konatransportes")
st.markdown("Extraia automaticamente os blocos de motoristas, ajudantes e rotas a partir da planilha!")

uploaded_file = st.file_uploader("📤 Envie a planilha (.xlsx ou .xls)", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)
    data = []

    for i in range(len(df)):
        row = df.iloc[i].astype(str).fillna("")

        # Detecta linhas que contêm os dados relevantes
        if any("LARGADA" in cell.upper() for cell in row):
            try:
                frota = str(df.iloc[i-1, 0]).strip()
                placa = str(df.iloc[i-1, 1]).strip()
                rota = str(df.iloc[i-1, 2]).strip()
                motorista = str(df.iloc[i, 2]).strip()
                ajud1 = str(df.iloc[i, 3]).strip()
                ajud2 = str(df.iloc[i, 4]).strip()
                largada = " ".join(re.findall(r"LARGADA\s+ÀS\s+\d{1,2}:\d{2}", " ".join(row), re.IGNORECASE))

                data.append({
                    "🚛 Frota": frota,
                    "🔢 Placa": placa,
                    "🗺️ Rota": rota,
                    "👨‍✈️ Motorista": motorista,
                    "🤝 Ajudante 1": ajud1,
                    "🤝 Ajudante 2": ajud2,
                    "⏰ Horário de Largada": largada
                })
            except Exception as e:
                pass

    if data:
        st.success("✅ Blocos extraídos com sucesso!")
        df_out = pd.DataFrame(data)
        st.dataframe(df_out, use_container_width=True)

        csv = df_out.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 Baixar CSV", csv, "escala_filtrada.csv", "text/csv")
    else:
        st.warning("⚠️ Nenhum bloco identificado. Verifique se o formato da planilha segue o padrão esperado.")
else:
    st.info("Envie o arquivo Excel acima para começar.")
