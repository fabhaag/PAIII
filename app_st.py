# --------------------------------------------------
# Rodar com o COMANDO: py -m streamlit run app_st.py
# --------------------------------------------------

import streamlit as st
import io
import zipfile
import numpy as np
import pandas as pd
import warnings
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image

# Configuração inicial da página Streamlit
st.set_page_config(page_title="Gerador de Certificados CMM - Schulz", layout="wide")

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

STRICT_NS = "http://openxmlformats.org"
TRANS_NS = "http://openxmlformats.org"

def repair_xlsx_stream(in_bytes):
    """Corrige problemas de namespace em arquivos Excel gerados por máquinas."""
    in_buf = io.BytesIO(in_bytes)
    out_buf = io.BytesIO()
    with zipfile.ZipFile(in_buf, "r") as z_in, zipfile.ZipFile(out_buf, "w") as z_out:
        for item in z_in.infolist():
            data = z_in.read(item.filename)
            if item.filename.endswith(".xml"):
                data_str = data.decode("utf-8")
                data_str = data_str.replace(STRICT_NS, TRANS_NS)
                data = data_str.encode("utf-8")
            z_out.writestr(item, data)
    out_buf.seek(0)
    return out_buf

# Dicionários de opções para os menus dropdown
OPCOES_CLIENTE = ["Metrologia", "Usinagem", "Pintura", "Compressores", "Lab. Fundição", "Q. Fundição", "Fusão", "Macharia", "Modelação", "Moldagem", "Gestão Ambiental", "Acabamento", "Manutenção", "SHMT", "IRA", "U. Compressores", "Eng. Processo", "Fabrica II", "Poka Yoke", "Logística", "Lab. China", "Lab. Sujidade", "Presset/Ferrament.", "Usinagem II", "Comp. Seg. do Trab. e Manut.", "Controle Dim. I", "Controle Dim. II", "Pintura Usi. II", "Schulz RVA", "Fund.", "Unid. 3"]
OPCOES_METROLOGISTA = ["José Alcides Rosa Jr", "Cleocir de Souza"]
OPCOES_MMC = ["500.008", "500.009", "500.011", "500.012", "500.015", "500.016", "500.018", "500.022"]
OPCOES_FORNECEDOR = ["CDM", "Usifer", "Torque", "Valente"]

# Função para inicializar valores no estado da sessão
def init_session_state():
    chaves = ['Teste', 'Dispositivo de Controle', 'Data', 'Ordem', 'Temperatura', 'Desenho', 'Hora', 'Numero peça', 'Cliente', 'Metrologista', 'MMC', 'Fornecedor', 'df_final', 'calibracao_sucesso', 'val_k', 'val_u', 'processado']
    for chave in chaves:
        if chave not in st.session_state:
            st.session_state[chave] = "" if chave not in ['df_final', 'calibracao_sucesso', 'val_k', 'val_u', 'processado'] else None
    if not st.session_state['Cliente']: st.session_state['Cliente'] = "Usinagem"
    if not st.session_state['MMC']: st.session_state['MMC'] = "500.011"

init_session_state()

def processar_planilha(uploaded_file):
    try:
        stream_corrigido = repair_xlsx_stream(uploaded_file.getvalue())
        
        # Extrai metadados do topo
        df_topo = pd.read_excel(stream_corrigido, sheet_name="Report", nrows=12, header=None, engine="openpyxl")
        
        # Mapeamento para atualizar a tela
        st.session_state['Dispositivo de Controle'] = str(df_topo.iloc[4, 1]) if pd.notna(df_topo.iloc[4, 1]) else ""
        st.session_state['Data'] = str(df_topo.iloc[4, 3]).split()[0] if pd.notna(df_topo.iloc[4, 3]) else ""
        st.session_state['Ordem'] = str(df_topo.iloc[4, 5]) if pd.notna(df_topo.iloc[4, 5]) else ""
        st.session_state['Desenho'] = str(df_topo.iloc[7, 1]) if pd.notna(df_topo.iloc[7, 1]) else ""
        st.session_state['Hora'] = str(df_topo.iloc[7, 3]) if pd.notna(df_topo.iloc[7, 3]) else ""
        st.session_state['Numero peça'] = str(df_topo.iloc[7, 5]) if pd.notna(df_topo.iloc[7, 5]) else ""
        
        metrologista = str(df_topo.iloc[10, 1]) if pd.notna(df_topo.iloc[10, 1]) else ""
        if metrologista in OPCOES_METROLOGISTA: st.session_state['Metrologista'] = metrologista
        
        mmc = str(df_topo.iloc[10, 3]) if pd.notna(df_topo.iloc[10, 3]) else ""
        if mmc in OPCOES_MMC: st.session_state['MMC'] = mmc
        
        fornecedor = str(df_topo.iloc[10, 5]) if pd.notna(df_topo.iloc[10, 5]) else ""
        if fornecedor in OPCOES_FORNECEDOR: st.session_state['Fornecedor'] = fornecedor

        # Extração e processamento da tabela
        stream_corrigido.seek(0)
        df_report = pd.read_excel(stream_corrigido, sheet_name="Report", header=12, engine="openpyxl")
        df_report = df_report.loc[:, ~df_report.columns.str.contains('^Unnamed')]
        df_report = df_report.dropna(subset=["Characteristic"])

        colunas_num = ["Actual", "Nominal", "Upper Tol", "Lower Tol"]
        for col in colunas_num:
            if col in df_report.columns:
                df_report[col] = pd.to_numeric(df_report[col], errors='coerce')

        df_report['Medida_Num'] = df_report.groupby('Characteristic').cumcount() + 1
        df_medidas = df_report.pivot(index='Characteristic', columns='Medida_Num', values='Actual').reset_index()
        
        if 1 in df_medidas.columns: df_medidas.rename(columns={1: 'Medida 1'}, inplace=True)
        else: df_medidas['Medida 1'] = np.nan
        if 2 in df_medidas.columns: df_medidas.rename(columns={2: 'Medida 2'}, inplace=True)
        else: df_medidas['Medida 2'] = np.nan

        df_limites = df_report.drop_duplicates(subset=['Characteristic'])[['Characteristic', 'Nominal', 'Upper Tol', 'Lower Tol']]

        df_final = pd.merge(df_limites, df_medidas, on='Characteristic')
        df_final['Média'] = df_final[['Medida 1', 'Medida 2']].mean(axis=1)
        df_final['Desvio padrão'] = df_final[['Medida 1', 'Medida 2']].std(axis=1, ddof=0)

        # Lógica de Avaliação
        calibracao_sucesso = True
        status_aprovado = []

        for i, row in df_final.iterrows():
            nominal, upper_tol, lower_tol, media = row['Nominal'], row['Upper Tol'], row['Lower Tol'], row['Média']
            aprovado = True

            if pd.notnull(nominal) and pd.notnull(media) and pd.notnull(upper_tol) and pd.notnull(lower_tol):
                limite_superior = nominal + upper_tol
                limite_inferior = nominal - lower_tol if lower_tol > 0 else nominal + lower_tol 
                if not (limite_inferior <= media <= limite_superior):
                    aprovado = False
                    calibracao_sucesso = False
            
            status_aprovado.append(aprovado)
        
        df_final['Aprovado'] = status_aprovado

        # Incerteza Fixa/Padrão
        st.session_state['val_k'] = 2.23
        st.session_state['val_u'] = 0.002
        st.session_state['df_final'] = df_final
        st.session_state['calibracao_sucesso'] = calibracao_sucesso
        st.session_state['processado'] = True
        
        st.success("Planilha lida e metadados importados com sucesso!")

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar a planilha: {str(e)}")

def gerar_pdf(dados_tela, df_final, val_k, val_u):
    dispositivo = dados_tela.get("Dispositivo de Controle", "N_A")
    data_raw = dados_tela.get("Data", "N_A")
    
    # Tratamento da data para converter YYYY-MM-DD em DD/MM/YYYY e fixar o ano
    if "-" in data_raw and len(data_raw.split("-")[0]) == 4:
        partes = data_raw.split("-")
        ano = partes[0]
        data_formatada = f"{partes[2]}/{partes[1]}/{partes[0]}"
    elif "/" in data_raw and len(data_raw.split("/")[-1]) == 4:
        partes = data_raw.split("/")
        ano = partes[-1]
        data_formatada = data_raw
    else:
        ano = datetime.datetime.now().strftime("%Y")
        data_formatada = data_raw.replace("-", "/")

    ordem = dados_tela.get("Ordem", "00000")
    cliente = dados_tela.get("Cliente", "")
    metrologista = dados_tela.get("Metrologista", "")
    fornecedor = dados_tela.get("Fornecedor", "Não consta")
    mmc = dados_tela.get("MMC", "")
    
    nome_arquivo = f"Certificado de Calibração {dispositivo} {data_formatada.replace('/', '-')}.pdf"
    
    # Gerando o PDF em memória (BytesIO)
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle(name='TitleStyle', parent=styles['Heading1'], alignment=1)
    style_normal = styles['Normal']
    style_bold = ParagraphStyle(name='BoldStyle', parent=styles['Normal'], fontName='Helvetica-Bold')
    
    elementos = []
    
    # Inserção da imagem superior centralizada (ajustada para largura útil da página A4 de 18cm)
    try:
        # IMPORTANTE: Substitua 'nome_da_sua_imagem.png' pelo caminho correto do arquivo
        img = Image('Imagens/BarraSuperior.jpg', width=18*cm, height=1.82*cm)
        img.hAlign = 'CENTER'
        elementos.append(img)
        elementos.append(Spacer(1, 12))
    except Exception:
        pass # Caso a imagem não seja encontrada, o gerador não será interrompido
    
    elementos.append(Paragraph(f"<b>Certificado de Calibração nº {ordem}/{ano}</b>   Data: {data_formatada}", style_title))
    elementos.append(Paragraph(f"<b>Solicitante:</b> {cliente}", style_normal))
    
    elementos.append(Spacer(1, 12))
    
    elementos.append(Paragraph("<b>1. Descrição do instrumento calibrado</b>", style_bold))
    elementos.append(Paragraph("Denominação: Dispositivo de controle", style_normal))
    elementos.append(Paragraph(f"Código: {dispositivo}", style_normal))
    elementos.append(Paragraph(f"Fabricante: {fornecedor}", style_normal))
    
    elementos.append(Spacer(1, 12))
    
    elementos.append(Paragraph("<b>2. Padrão utilizado</b>", style_bold))
    elementos.append(Paragraph(f"Máquina de medição por coordenadas {mmc} - Carl Zeiss", style_normal))
    elementos.append(Paragraph(f"Certificado de Calibração n° 241120502390 - Carl Zeiss; de 20/11/2024 - Próxima Calibração: 31/08/2026", style_normal))
    elementos.append(Spacer(1, 12))
    
    elementos.append(Paragraph("<b>3. Procedimento de calibração</b>", style_bold))
    elementos.append(Paragraph("Nr. do procedimento / revisão: Calibração de dispositivo de controle - ITL-0026 (revisão 2).", style_normal))
    elementos.append(Paragraph("O instrumento foi medido na máquina de medição por coordenadas através de elementos geométricos definidos, sendo o resultado apresentado a média de duas medições realizadas.", style_normal))
    elementos.append(Paragraph("Os resultados estão referenciados a temperatura de 20° Celsius.", style_normal))
    elementos.append(Spacer(1, 12))
    
    elementos.append(Paragraph("<b>4. Resultados (mm)</b>", style_bold))
    elementos.append(Spacer(1, 6))
    
    dados_tabela = [["Parâmetro avaliado", "Valor nominal", "Tol. +", "Tol. -", "Valor obtido", "Erro"]]
    
    for _, row in df_final.iterrows():
        parametro = row['Characteristic']
        nominal = row['Nominal']
        media_obtida = row['Média']
        
        try:
            erro_calculado = abs(float(media_obtida) - float(nominal))
            erro = f"{erro_calculado:.3f}".replace('.', ',')
        except (ValueError, TypeError):
            erro = "-"
            
        dados_tabela.append([
            str(parametro), 
            f"{nominal:.3f}".replace('.', ',') if pd.notna(nominal) else "-", 
            f"{row['Upper Tol']:.3f}".replace('.', ',') if pd.notna(row['Upper Tol']) else "-", 
            f"{row['Lower Tol']:.3f}".replace('.', ',') if pd.notna(row['Lower Tol']) else "-", 
            f"{media_obtida:.3f}".replace('.', ',') if pd.notna(media_obtida) else "-", 
            erro
        ])
        
    tabela = Table(dados_tabela, colWidths=[5.5*cm, 2.5*cm, 2*cm, 2*cm, 2.5*cm, 2*cm])
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.whitesmoke),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (0,0), (0,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
    ]))
    elementos.append(tabela)
    elementos.append(Spacer(1, 12))
    
    elementos.append(Paragraph(f"<b>5. Incerteza de medição:</b> {val_u} mm, k = {val_k}", style_bold))
    elementos.append(Spacer(1, 12))
    
    elementos.append(Paragraph("<b>6. Observações</b>", style_bold))
    elementos.append(Paragraph("*Os parâmetros avaliados foram obtidos da Ficha de Auto Controle do produto que o dispositivo controla.", style_normal))
    elementos.append(Paragraph(f"Técnico executante: {metrologista}", style_normal))
    elementos.append(Spacer(1, 24))

    elementos.append(Paragraph("_______________________________", style_normal))
    elementos.append(Paragraph("Signatário Autorizado", style_normal))
    elementos.append(Paragraph("João Carlos Gomes de Oliveira", style_normal))
    
    doc.build(elementos)
    pdf_buffer.seek(0)
    return pdf_buffer, nome_arquivo

# ==========================================
# INTERFACE DO STREAMLIT
# ==========================================

st.title("Automação de Certificados CMM - Schultz")
st.markdown("---")

# Uploader de Arquivo
uploaded_file = st.file_uploader("Carregar Planilha de Medições (Excel)", type=["xlsx", "xls"])

# Lógica para resetar a tela ao remover o arquivo ou carregar um novo
if uploaded_file is None:
    # Se o arquivo foi removido (clicou no X), mas o estado consta como processado
    if st.session_state.get('processado'):
        # Limpa todas as variáveis da sessão
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        init_session_state() # Reinicializa com os valores padrão
        st.rerun() # Atualiza a página (F5) limpando a tela
else:
    # Verifica se o arquivo carregado é diferente do que está processado na memória
    if st.session_state.get('nome_arquivo_atual') != uploaded_file.name:
        # Limpa o estado anterior para receber o novo arquivo
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        init_session_state()
        
        # Salva o nome do novo arquivo e o processa
        st.session_state['nome_arquivo_atual'] = uploaded_file.name
        processar_planilha(uploaded_file)
        st.rerun() # Atualiza a página (F5) com os novos dados

st.markdown("### Metadados do Certificado")

# Controles em Grade (Grid)
col1, col2, col3, col4 = st.columns(4)

with col1:
    teste = st.text_input("Teste", value=st.session_state['Teste'])
    temp = st.text_input("Temperatura (°C)", value=st.session_state['Temperatura'])
    cliente = st.selectbox("Cliente", options=OPCOES_CLIENTE, index=OPCOES_CLIENTE.index(st.session_state['Cliente']) if st.session_state['Cliente'] in OPCOES_CLIENTE else 1)

with col2:
    dispositivo = st.text_input("Dispositivo de Controle", value=st.session_state['Dispositivo de Controle'])
    desenho = st.text_input("Desenho", value=st.session_state['Desenho'])
    metro = st.selectbox("Metrologista", options=OPCOES_METROLOGISTA, index=OPCOES_METROLOGISTA.index(st.session_state['Metrologista']) if st.session_state['Metrologista'] in OPCOES_METROLOGISTA else 0)

with col3:
    data_med = st.text_input("Data (dd/mm/aaaa)", value=st.session_state['Data'])
    hora = st.text_input("Hora (HH:MM)", value=st.session_state['Hora'])
    mmc_sel = st.selectbox("MMC", options=OPCOES_MMC, index=OPCOES_MMC.index(st.session_state['MMC']) if st.session_state['MMC'] in OPCOES_MMC else 2)

with col4:
    ordem = st.text_input("Ordem", value=st.session_state['Ordem'])
    peca = st.text_input("Numero peça", value=st.session_state['Numero peça'])
    fornecedor = st.selectbox("Fornecedor", options=OPCOES_FORNECEDOR, index=OPCOES_FORNECEDOR.index(st.session_state['Fornecedor']) if st.session_state['Fornecedor'] in OPCOES_FORNECEDOR else 0)

st.markdown("---")

# Tabela de Resultados
if st.session_state.get('processado'):
    st.markdown("### PARÂMETROS LIDOS E CONSOLIDADOS")
    
    df_view = st.session_state['df_final'].copy()
    
    # Formatação de Estilo Pandas (Verde para Aprovado, Vermelho para Reprovado)
    def color_rows(row):
        color = '#c8e6c9' if row['Aprovado'] else '#ffcdd2'
        return [f'background-color: {color}; color: black' for _ in row]

    # Arredondando dados na visualização para ficar legível e aplicando estilo
    format_dict = {col: "{:.3f}" for col in ["Nominal", "Upper Tol", "Lower Tol", "Medida 1", "Medida 2", "Média", "Desvio padrão"]}
    st.dataframe(
        df_view.style.apply(color_rows, axis=1).format(format_dict, na_rep="-"),
        use_container_width=True, 
        hide_index=True
    )
    
    # --- CÓDIGO DA LEGENDA A ADICIONAR ---
    st.markdown(
        """
        <div style="display: flex; gap: 20px; padding-top: 10px; padding-bottom: 15px;">
            <div style="display: flex; align-items: center;">
                <div style="width: 15px; height: 15px; background-color: #c8e6c9; border: 1px solid #757575; margin-right: 8px;"></div>
                <span style="font-size: 15px; font-weight: 500;">DENTRO DA TOLERANCIA</span>
            </div>
            <div style="display: flex; align-items: center;">
                <div style="width: 15px; height: 15px; background-color: #ffcdd2; border: 1px solid #757575; margin-right: 8px;"></div>
                <span style="font-size: 15px; font-weight: 500;">FORA DA TOLERANCIA</span>
            </div>
        </div>
        """, 
        unsafe_allow_html=True
    )
    # ------------------------------------
    
    # Exibição da Incerteza e Status Final
    st.markdown(f"<h4 style='text-align: center; color: #0052cc;'>Fator de abrangência (k): {st.session_state['val_k']} &nbsp;|&nbsp; Incerteza (U): {st.session_state['val_u']:.4f} mm</h4>", unsafe_allow_html=True)

    if st.session_state['calibracao_sucesso']:
        st.markdown("<h3 style='text-align: center; color: #155724;'>APROVADO</h3>", unsafe_allow_html=True)
    else:
        st.markdown("<h3 style='text-align: center; color: #721c24;'>REPROVADO</h3>", unsafe_allow_html=True)

    # Coleta de dados para geração do PDF
    dados_tela = {
        "Dispositivo de Controle": dispositivo,
        "Data": data_med,
        "Ordem": ordem,
        "Cliente": cliente,
        "Metrologista": metro,
        "Fornecedor": fornecedor,
        "MMC": mmc_sel
    }

    # Botão para Download do PDF
    pdf_buffer, nome_arquivo = gerar_pdf(dados_tela, df_view, st.session_state['val_k'], st.session_state['val_u'])
    
    col_btn1, col_btn2, col_btn3 = st.columns([1,2,1])
    with col_btn2:
        st.download_button(
            label="📄 Baixar Certificado em PDF",
            data=pdf_buffer,
            file_name=nome_arquivo,
            mime="application/pdf",
            use_container_width=True
        )