import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Gerador de Texto de Empresas", layout="centered")
st.title("Gerador de Texto a partir de Excel")

uploaded_file = st.file_uploader("Faça upload do arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    texto_final = ""

    xls = pd.ExcelFile(uploaded_file)
    abas_validas = []

    for aba in xls.sheet_names:
        df = xls.parse(aba)
        colunas = [col.strip().lower() for col in df.columns.astype(str)]
        if any("itemid" in c for c in colunas) and any("apelido" in c for c in colunas):
            abas_validas.append(aba)

    if not abas_validas:
        st.error("Nenhuma aba contém as colunas necessárias (ItemId e APELIDO).")
    else:
        for aba in abas_validas:
            df = xls.parse(aba)
            df.columns = [col.strip().lower() for col in df.columns.astype(str)]

            df.rename(columns={
                "itemid": "itemid",
                "apelido": "apelido",
                "tipo documento": "tipo",
                "documento": "documento",
                "endereco": "endereco",
                "sobrenome": "sobrenome",
                "nome": "nome",
                "email": "email",
                "telefone_1": "telefone_1"
            }, inplace=True)

            empresas = df.groupby("apelido")

            for nome_apelido, grupo in empresas:
                if pd.isna(nome_apelido):
                    continue

                mlbs = grupo["itemid"].dropna().astype(str).tolist()
                if not mlbs:
                    continue  # pula se não tiver itemid válido

                texto = f"{', '.join(mlbs)}\n"
                texto += f"Nome: {nome_apelido}\n"

                # Nome/Razão social = nome + sobrenome
                nome_valor = grupo["nome"].dropna().astype(str).iloc[0] if "nome" in grupo.columns and not grupo["nome"].dropna().empty else ""
                sobrenome_valor = grupo["sobrenome"].dropna().astype(str).iloc[0] if "sobrenome" in grupo.columns and not grupo["sobrenome"].dropna().empty else ""
                if nome_valor or sobrenome_valor:
                    nome_completo = f"{nome_valor} {sobrenome_valor}".strip()
                    texto += f"Nome/Razão social: {nome_completo}\n"

                # Documento (CNPJ/CPF)
                if "tipo" in grupo.columns and "documento" in grupo.columns:
                    tipo = grupo["tipo"].dropna().astype(str).str.lower().iloc[0] if not grupo["tipo"].dropna().empty else None
                    documento = grupo["documento"].dropna().astype(str).iloc[0] if not grupo["documento"].dropna().empty else None
                    if tipo and documento:
                        tipo_str = "CNPJ" if "cnpj" in tipo else ("CPF" if "cpf" in tipo else tipo.upper())
                        texto += f"{tipo_str}: {documento}\n"

                if "email" in grupo.columns:
                    email = grupo["email"].dropna().astype(str).iloc[0] if not grupo["email"].dropna().empty else None
                    if email:
                        texto += f"E-mail: {email}\n"

                if "endereco" in grupo.columns:
                    end = grupo["endereco"].dropna().astype(str).iloc[0] if not grupo["endereco"].dropna().empty else None
                    if end:
                        texto += f"Endereço: {end}\n"

                if "telefone_1" in grupo.columns:
                    tel = grupo["telefone_1"].dropna().astype(str).iloc[0] if not grupo["telefone_1"].dropna().empty else None
                    if tel:
                        texto += f"Telefone: {tel}\n"

                texto += "\n"
                texto_final += texto

        st.text_area("Texto gerado:", value=texto_final, height=400)

        buffer = io.StringIO(texto_final)
        st.download_button(
            label="Baixar texto como .txt",
            data=buffer.getvalue(),
            file_name="empresas_formatado.txt",
            mime="text/plain"
        )
