import streamlit as st
import pandas as pd
import zipfile
import os
import xml.etree.ElementTree as ET
import tempfile
from datetime import datetime
from io import BytesIO   # para exportar Excel


# ===============================
# Configuração visual do app (corporativo)
# ===============================
st.set_page_config(
    page_title="Leitor de XMLs",
    page_icon="�",
    layout="wide"
)

# Logo e título
col_logo, col_title = st.columns([1, 8])
with col_logo:
    st.image("icon-xml-excel.svg", width=120)
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
with col_title:
    st.markdown("""
        <h1 style='margin-bottom:0; color:#1F2937; font-size:2.5rem;'>Leitor de XMLs</h1>
        <span style='color:#4B5563; font-size:1.2rem;'>Conversão de NFe/NFCe e CTe para Excel</span>
    """, unsafe_allow_html=True)

st.markdown("---")

# Instruções
with st.expander("ℹ️ Como usar", expanded=True):
    st.markdown("""
    1. Faça upload de um ou mais arquivos ZIP contendo XMLs de NFe ou CTe.<br>
    2. Utilize os filtros na barra lateral para verificar os resultados.<br>
    3. Baixe a planilha Excel pronta para análise.<br>
    <br>
    <span style='color:#6B7280;'>Atenção: Apenas arquivos XML válidos serão processados.</span>
    """, unsafe_allow_html=True)

# ===============================
# Função para extrair XMLs de um ZIP
# ===============================
def extrair_xmls_de_zip(zip_path, extract_path):
    xml_files = []
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)
        for root, _, files in os.walk(extract_path):
            for file in files:
                if file.endswith('.xml'):
                    xml_files.append(os.path.join(root, file))
    return xml_files

# ===============================
# Processar NFe por item
# ===============================
def processar_nfe_por_item(xml_path, ns):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        emit = root.find('.//ns:emit', ns)
        ide = root.find('.//ns:ide', ns)
        total = root.find('.//ns:total', ns)
        det_list = root.findall('.//ns:det', ns)

        if emit is None or ide is None or total is None:
            return []

        chave_acesso_tag = root.find('.//ns:infProt/ns:chNFe', ns)
        chave_acesso = chave_acesso_tag.text if chave_acesso_tag is not None else ""

        status_tag = root.find('.//ns:infProt/ns:cStat', ns)
        status = status_tag.text if status_tag is not None else ""

        emitente = emit.find('ns:xNome', ns).text if emit.find('ns:xNome', ns) is not None else ""
        cnpj_emitente = emit.find('ns:CNPJ', ns).text if emit.find('ns:CNPJ', ns) is not None else ""
        uf_emitente = emit.find('ns:enderEmit/ns:UF', ns).text if emit.find('ns:enderEmit/ns:UF', ns) is not None else ""
        numero_nfe = ide.find('ns:nNF', ns).text if ide.find('ns:nNF', ns) is not None else ""
        data_emissao = ide.find('ns:dhEmi', ns).text if ide.find('ns:dhEmi', ns) is not None else ""

        dados = []
        for det in det_list:
            prod = det.find('ns:prod', ns)
            imposto = det.find('ns:imposto', ns)
            if prod is None or imposto is None:
                continue

            icms = imposto.find('.//ns:ICMS', ns)
            icms_valor = icms.find('.//ns:vICMS', ns)
            icms_aliquota = icms.find('.//ns:pICMS', ns)
            icms_cst = icms.find('.//ns:CST', ns)
            icms_desonerado = icms.find('.//ns:vICMSDeson', ns)

            ipi_valor = imposto.find('.//ns:IPI/ns:IPITrib/ns:vIPI', ns)
            pis_valor = imposto.find('.//ns:PIS/ns:PISAliq/ns:vPIS', ns)
            cofins_valor = imposto.find('.//ns:COFINS/ns:COFINSAliq/ns:vCOFINS', ns)
            icms_st_valor = imposto.find('.//ns:ICMS/*/ns:vICMSST', ns)

            cbenef = prod.find('ns:cBenef', ns)
            cfop = prod.find('ns:CFOP', ns)

            frete = root.find('.//ns:transp/ns:vFrete', ns)
            seguro = root.find('.//ns:transp/ns:vSeg', ns)

            dados.append({
                "Número NFe": numero_nfe,
                "Data de Emissão": data_emissao,
                "CNPJ Emitente": cnpj_emitente,
                "Emitente": emitente,
                "UF Emitente": uf_emitente,
                "Valor da Nota": total.find('ns:ICMSTot/ns:vNF', ns).text if total.find('ns:ICMSTot/ns:vNF', ns) is not None else "",
                "ICMS": icms_valor.text if icms_valor is not None else "",
                "Alíquota ICMS": icms_aliquota.text if icms_aliquota is not None else "",
                "IPI": ipi_valor.text if ipi_valor is not None else "",
                "PIS": pis_valor.text if pis_valor is not None else "",
                "COFINS": cofins_valor.text if cofins_valor is not None else "",
                "ICMS ST": icms_st_valor.text if icms_st_valor is not None else "",
                "Frete": frete.text if frete is not None else "",
                "Seguro": seguro.text if seguro is not None else "",
                "Chave de Acesso": chave_acesso,
                "cBenef": cbenef.text if cbenef is not None else "",
                "ICMS Desonerado": icms_desonerado.text if icms_desonerado is not None else "",
                "CFOP": cfop.text if cfop is not None else "",
                "CST ICMS": icms_cst.text if icms_cst is not None else "",
                "Status da NFe": status
            })
        return dados
    except ET.ParseError:
        st.error(f"Erro ao analisar o arquivo XML: {os.path.basename(xml_path)}")
        return []

# ===============================
# Processar NFe por cabeçalho
# ===============================
def processar_nfe_por_cabecalho(xml_path, ns):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        # Tenta pegar o CFOP do primeiro item (det)
        det = root.find('.//ns:det', ns)
        cfop = det.find('ns:prod/ns:CFOP', ns).text if det is not None and det.find('ns:prod/ns:CFOP', ns) is not None else ""
        # ...restante do código permanece igual...
        emit = root.find('.//ns:emit', ns)
        ide = root.find('.//ns:ide', ns)
        total = root.find('.//ns:total', ns)
        if emit is None or ide is None or total is None:
            return []

        chave_acesso_tag = root.find('.//ns:infProt/ns:chNFe', ns)
        chave_acesso = chave_acesso_tag.text if chave_acesso_tag is not None else ""

        status_tag = root.find('.//ns:infProt/ns:cStat', ns)
        status = status_tag.text if status_tag is not None else ""

        emitente = emit.find('ns:xNome', ns).text if emit.find('ns:xNome', ns) is not None else ""
        cnpj_emitente = emit.find('ns:CNPJ', ns).text if emit.find('ns:CNPJ', ns) is not None else ""
        uf_emitente = emit.find('ns:enderEmit/ns:UF', ns).text if emit.find('ns:enderEmit/ns:UF', ns) is not None else ""
        numero_nfe = ide.find('ns:nNF', ns).text if ide.find('ns:nNF', ns) is not None else ""
        data_emissao = ide.find('ns:dhEmi', ns).text if ide.find('ns:dhEmi', ns) is not None else ""

        # Identificadores extras
    # cNF e cDV removidos conforme solicitado
        # CST/CSOSN (do primeiro item)
        cst_csosn = ""
        if det is not None:
            # Simples Nacional: busca CSOSN
            csosn = det.find('.//ns:CSOSN', ns)
            if csosn is not None and csosn.text:
                cst_csosn = csosn.text
            else:
                # Regime normal: busca CST
                cst = det.find('.//ns:CST', ns)
                if cst is not None and cst.text:
                    cst_csosn = cst.text
        modelo = ide.find('ns:mod', ns).text if ide.find('ns:mod', ns) is not None else ""
        serie = ide.find('ns:serie', ns).text if ide.find('ns:serie', ns) is not None else ""
        versao = root.attrib.get('versao', "")
        cUF = ide.find('ns:cUF', ns).text if ide.find('ns:cUF', ns) is not None else ""

        frete = root.find('.//ns:transp/ns:vFrete', ns)
        seguro = root.find('.//ns:transp/ns:vSeg', ns)

        return [{
            "Chave de Acesso": chave_acesso,
            "Número NFe": numero_nfe,
            "Série": serie,
            "Modelo": modelo,
            # "UF (cUF)": cUF,  # removido
            # "Versão": versao,  # removido
            "Data de Emissão": data_emissao,
            "CNPJ Emitente": cnpj_emitente,
            "Emitente": emitente,
            "UF Emitente": uf_emitente,
            "Valor da Nota": total.find('ns:ICMSTot/ns:vNF', ns).text if total.find('ns:ICMSTot/ns:vNF', ns) is not None else "",
            "CST/CSOSN": cst_csosn,
            "ICMS": total.find('ns:ICMSTot/ns:vICMS', ns).text if total.find('ns:ICMSTot/ns:vICMS', ns) is not None else "",
            "IPI": total.find('ns:ICMSTot/ns:vIPI', ns).text if total.find('ns:ICMSTot/ns:vIPI', ns) is not None else "",
            "PIS": total.find('ns:ICMSTot/ns:vPIS', ns).text if total.find('ns:ICMSTot/ns:vPIS', ns) is not None else "",
            "COFINS": total.find('ns:ICMSTot/ns:vCOFINS', ns).text if total.find('ns:ICMSTot/ns:vCOFINS', ns) is not None else "",
            "ICMS ST": total.find('ns:ICMSTot/ns:vST', ns).text if total.find('ns:ICMSTot/ns:vST', ns) is not None else "",
            "Frete": frete.text if frete is not None else "",
            "Seguro": seguro.text if seguro is not None else "",
            "ICMS Desonerado": total.find('ns:ICMSTot/ns:vICMSDeson', ns).text if total.find('ns:ICMSTot/ns:vICMSDeson', ns) is not None else "",
            "CFOP": cfop,
            "Status da NFe": status
        }]
    except ET.ParseError:
        st.error(f"Erro ao analisar o arquivo XML: {os.path.basename(xml_path)}")
        return []

        emit = root.find('.//ns:emit', ns)
        ide = root.find('.//ns:ide', ns)
        total = root.find('.//ns:total', ns)
        if emit is None or ide is None or total is None:
            return []

        chave_acesso_tag = root.find('.//ns:infProt/ns:chNFe', ns)
        chave_acesso = chave_acesso_tag.text if chave_acesso_tag is not None else ""

        status_tag = root.find('.//ns:infProt/ns:cStat', ns)
        status = status_tag.text if status_tag is not None else ""

        emitente = emit.find('ns:xNome', ns).text if emit.find('ns:xNome', ns) is not None else ""
        cnpj_emitente = emit.find('ns:CNPJ', ns).text if emit.find('ns:CNPJ', ns) is not None else ""
        uf_emitente = emit.find('ns:enderEmit/ns:UF', ns).text if emit.find('ns:enderEmit/ns:UF', ns) is not None else ""
        numero_nfe = ide.find('ns:nNF', ns).text if ide.find('ns:nNF', ns) is not None else ""
        data_emissao = ide.find('ns:dhEmi', ns).text if ide.find('ns:dhEmi', ns) is not None else ""

        # Identificadores extras
        cNF = ide.find('ns:cNF', ns).text if ide.find('ns:cNF', ns) is not None else ""
        cDV = ide.find('ns:cDV', ns).text if ide.find('ns:cDV', ns) is not None else ""
        modelo = ide.find('ns:mod', ns).text if ide.find('ns:mod', ns) is not None else ""
        serie = ide.find('ns:serie', ns).text if ide.find('ns:serie', ns) is not None else ""
        versao = root.attrib.get('versao', "")
        cUF = ide.find('ns:cUF', ns).text if ide.find('ns:cUF', ns) is not None else ""

        frete = root.find('.//ns:transp/ns:vFrete', ns)
        seguro = root.find('.//ns:transp/ns:vSeg', ns)

        return [{
            "Chave de Acesso": chave_acesso,
            "Número NFe": numero_nfe,
            "Série": serie,
            "Modelo": modelo,
            "UF (cUF)": cUF,
            "Versão": versao,
            "Data de Emissão": data_emissao,
            "CNPJ Emitente": cnpj_emitente,
            "Emitente": emitente,
            "UF Emitente": uf_emitente,
            "Valor da Nota": total.find('ns:ICMSTot/ns:vNF', ns).text if total.find('ns:ICMSTot/ns:vNF', ns) is not None else "",
            "CST/CSOSN": cst_csosn,
            "ICMS": total.find('ns:ICMSTot/ns:vICMS', ns).text if total.find('ns:ICMSTot/ns:vICMS', ns) is not None else "",
            "IPI": total.find('ns:ICMSTot/ns:vIPI', ns).text if total.find('ns:ICMSTot/ns:vIPI', ns) is not None else "",
            "PIS": total.find('ns:ICMSTot/ns:vPIS', ns).text if total.find('ns:ICMSTot/ns:vPIS', ns) is not None else "",
            "COFINS": total.find('ns:ICMSTot/ns:vCOFINS', ns).text if total.find('ns:ICMSTot/ns:vCOFINS', ns) is not None else "",
            "ICMS ST": total.find('ns:ICMSTot/ns:vST', ns).text if total.find('ns:ICMSTot/ns:vST', ns) is not None else "",
            "Frete": frete.text if frete is not None else "",
            "Seguro": seguro.text if seguro is not None else "",
            "ICMS Desonerado": total.find('ns:ICMSTot/ns:vICMSDeson', ns).text if total.find('ns:ICMSTot/ns:vICMSDeson', ns) is not None else "",
            "CFOP": cfop,
            "Status da NFe": status
        }]
    except ET.ParseError:
        st.error(f"Erro ao analisar o arquivo XML: {os.path.basename(xml_path)}")
        return []

# ===============================
# Processar CTe
# ===============================
def processar_cte(xml_path, ns):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        ide = root.find('.//ns:ide', ns)
        emit = root.find('.//ns:emit', ns)
        valor_total = root.find('.//ns:vTPrest', ns)
        icms = root.find('.//ns:ICMS00', ns)
        chave_acesso_tag = root.find('.//ns:infProt/ns:chCTe', ns)

        if ide is None or emit is None or valor_total is None or chave_acesso_tag is None:
            return []

        chave_acesso = chave_acesso_tag.text if chave_acesso_tag is not None else ""

        return [{
            "Número CTe": ide.find('ns:nCT', ns).text if ide.find('ns:nCT', ns) is not None else "",
            "Data de Emissão": ide.find('ns:dhEmi', ns).text if ide.find('ns:dhEmi', ns) is not None else "",
            "CNPJ Emitente": emit.find('ns:CNPJ', ns).text if emit.find('ns:CNPJ', ns) is not None else "",
            "Emitente": emit.find('ns:xNome', ns).text if emit.find('ns:xNome', ns) is not None else "",
            "UF Emitente": emit.find('ns:enderEmit/ns:UF', ns).text if emit.find('ns:enderEmit/ns:UF', ns) is not None else "",
            "Valor Total": valor_total.text if valor_total is not None else "",
            "ICMS": icms.find('ns:vICMS', ns).text if icms is not None and icms.find('ns:vICMS', ns) is not None else "",
            "Chave de Acesso": chave_acesso
        }]
    except ET.ParseError:
        st.error(f"Erro ao analisar o arquivo XML: {os.path.basename(xml_path)}")
        return []

# ===============================
# Interface Streamlit
# ===============================
def main():
    st.title("XML to EXCEL")

    tipo_doc = st.radio("Tipo de Documento:", ["NFe", "CTe"])
    layout = "Cabeçalho"  # Fixo, não mostra mais opção

    uploaded_files = st.file_uploader(
        "Selecione um ou mais arquivos ZIP com os XMLs",
        type="zip",
        accept_multiple_files=True,
        help="Apenas arquivos ZIP contendo XMLs de NFe ou CTe.",
        label_visibility="visible"
    )

    if uploaded_files:
        with st.spinner("Processando arquivos..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                xml_files = []

                for uploaded_file in uploaded_files:
                    zip_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(zip_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    arquivos_extraidos = extrair_xmls_de_zip(zip_path, temp_dir)
                    xml_files.extend(arquivos_extraidos)

                if not xml_files:
                    st.warning("Nenhum arquivo XML encontrado nos ZIPs.")
                else:
                    st.markdown(f"<div style='background-color:#E8EEF5; color:#1F2937; border-radius:8px; padding:0.7em 1em; margin-bottom:1em; font-size:1.1em;'><b>{len(xml_files)}</b> arquivo(s) XML encontrado(s)</div>", unsafe_allow_html=True)

                    progress_bar = st.progress(0)
                    dados_totais = []

                    for i, xml_file in enumerate(xml_files):
                        progress_bar.progress((i + 1) / len(xml_files))

                        if tipo_doc == "NFe":
                            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
                            dados_totais.extend(processar_nfe_por_cabecalho(xml_file, ns))
                        else:
                            ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}
                            dados_totais.extend(processar_cte(xml_file, ns))

                    if dados_totais:
                        df = pd.DataFrame(dados_totais)
                        if 'Data de Emissão' in df.columns:
                            df['Data de Emissão'] = pd.to_datetime(df['Data de Emissão'], errors='coerce', utc=True).dt.date

                        with st.sidebar:
                            st.markdown("<b>Filtros</b>", unsafe_allow_html=True)
                            cfop_options = sorted(df['CFOP'].unique().tolist()) if 'CFOP' in df.columns else []
                            selected_cfops = st.multiselect("Filtrar por CFOP:", cfop_options)

                            if 'Data de Emissão' in df.columns and not df['Data de Emissão'].isna().all():
                                min_date = df['Data de Emissão'].min()
                                max_date = df['Data de Emissão'].max()
                            else:
                                min_date = max_date = datetime.now().date()

                            start_date = st.date_input('Data de início', min_date)
                            end_date = st.date_input('Data final', max_date)

                        df_filtered = df.copy()

                        if selected_cfops:
                            df_filtered = df_filtered[df_filtered['CFOP'].isin(selected_cfops)]

                        if 'Data de Emissão' in df_filtered.columns:
                            df_filtered = df_filtered[(df_filtered['Data de Emissão'] >= start_date) & (df_filtered['Data de Emissão'] <= end_date)]

                        st.markdown("""
                            <div style='background-color:#F3F4F6; border-radius:10px; padding:1.5rem 1rem 1rem 1rem; margin-bottom:1.5rem;'>
                                <h3 style='color:#1F2937; margin-bottom:0.5rem;'>Notas Extraídas</h3>
                                <div style='font-size:0.95rem; color:#6B7280; margin-bottom:1rem;'>Veja abaixo a tabela com os dados extraídos dos XMLs.</div>
                        """, unsafe_allow_html=True)
                        st.dataframe(df_filtered, use_container_width=True)
                        st.markdown("</div>", unsafe_allow_html=True)

                        # Exporta para Excel
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            df_filtered.to_excel(writer, index=False, sheet_name='Notas')
                        output.seek(0)

                        # Define nome do arquivo com nome do emitente (se houver)
                        nome_emitente = ""
                        if 'Emitente' in df_filtered.columns and not df_filtered['Emitente'].isna().all():
                            nome_emitente = df_filtered['Emitente'].iloc[0]
                            if isinstance(nome_emitente, str):
                                nome_emitente = nome_emitente.strip().replace(' ', '_').replace('/', '_')
                        file_name = f"notas_{nome_emitente}.xlsx" if nome_emitente else "notas.xlsx"

                        st.download_button(
                            label="📥 Baixar Planilha Excel (.xlsx)",
                            data=output,
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            help="Baixe a planilha pronta para análise corporativa."
                        )
                    else:
                        st.warning("Nenhum dado válido foi extraído dos arquivos XML.")

    st.markdown("---")
    st.markdown("<div style='text-align:right; color:#6B7280; font-size:0.95rem;'>Desenvolvido por Beatriz Lourenço</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
