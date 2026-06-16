import streamlit as st
import pandas as pd
import re
import io

# --- CONFIGURAÇÃO DA PÁGINA E ESTILIZAÇÃO VISUAL PREMIUM ---
st.set_page_config(page_title="Gerador de Planilha", page_icon="✨", layout="centered")

css_personalizado = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;700&display=swap');

    /* Fundo da tela inteira */
    .stApp {
        background-color: #0d0d0d;
        font-family: 'Montserrat', sans-serif;
    }
    
    /* Fontes globais */
    html, body, [class*="css"] {
        font-family: 'Montserrat', sans-serif !important;
    }

    /* Título com Gradiente Dourado */
    .titulo-premium {
        background: linear-gradient(45deg, #BF953F, #FCF6BA, #B38728, #FBF5B7, #AA771C);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 700;
        font-size: 2.8rem;
        text-align: center;
        margin-bottom: 10px;
    }
    
    .subtitulo {
        text-align: center;
        color: #A0A0A0;
        font-weight: 300;
        font-size: 1.1rem;
        margin-bottom: 30px;
    }

    /* Textos em geral */
    p, span, label, div, li {
        color: #E5E4E2 !important;
    }

    /* Área de Upload de Arquivos (Glassmorphism) */
    [data-testid="stFileUploadDropzone"] {
        background: rgba(26, 26, 26, 0.6);
        border: 2px dashed rgba(212, 175, 55, 0.4);
        border-radius: 15px;
        backdrop-filter: blur(5px);
        transition: all 0.3s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #FFDF00;
        background: rgba(36, 36, 36, 0.9);
        box-shadow: 0px 0px 20px rgba(212, 175, 55, 0.15);
    }

    /* Botão Principal de Processar */
    div.stButton > button {
        background: linear-gradient(135deg, #1a1a1a 0%, #0d0d0d 100%) !important;
        color: #D4AF37 !important;
        font-weight: 600 !important;
        font-size: 16px !important;
        letter-spacing: 1px;
        border: 1px solid #D4AF37 !important;
        border-radius: 10px !important;
        padding: 20px !important;
        transition: all 0.4s ease;
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #D4AF37 0%, #AA771C 100%) !important;
        color: #0d0d0d !important;
        box-shadow: 0px 10px 20px rgba(212, 175, 55, 0.4) !important;
        border: 1px solid transparent !important;
        transform: translateY(-2px);
    }

    /* Botão de Download */
    div.stDownloadButton > button {
        background: linear-gradient(135deg, #D4AF37 0%, #AA771C 100%) !important;
        color: #0d0d0d !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 20px !important;
        box-shadow: 0px 5px 15px rgba(212, 175, 55, 0.3) !important;
        transition: all 0.3s ease;
    }
    div.stDownloadButton > button:hover {
        background: linear-gradient(135deg, #FFDF00 0%, #D4AF37 100%) !important;
        color: #000000 !important;
        transform: translateY(-3px);
        box-shadow: 0px 8px 20px rgba(212, 175, 55, 0.5) !important;
    }
    
    /* Expander e Caixas de Alerta/Sucesso forçadas para Preto e Dourado */
    [data-testid="stExpander"], [data-testid="stAlert"] {
        background: rgba(26, 26, 26, 0.8) !important;
        border: 1px solid #D4AF37 !important;
        border-radius: 10px !important;
    }
    [data-testid="stAlert"] p, [data-testid="stAlert"] span, [data-testid="stAlert"] div {
        color: #FFDF00 !important;
        font-weight: 600;
    }
    [data-testid="stAlert"] svg {
        fill: #FFDF00 !important;
    }
    
    /* Separador Dourado */
    hr {
        border-top: 1px solid rgba(212, 175, 55, 0.2);
    }
</style>
"""
st.markdown(css_personalizado, unsafe_allow_html=True)

# --- FUNÇÕES DE PROCESSAMENTO (NÚCLEO INTACTO) ---
def ler_aba_inteligente(xls, nome_aba):
    df_bruto = pd.read_excel(xls, sheet_name=nome_aba, header=None)
    linha_cabecalho = 0
    for idx, row in df_bruto.iterrows():
        valores_linha = [str(val).lower().strip() for val in row.values]
        if any('matrícula' in s or 'matricula' in s for s in valores_linha):
            linha_cabecalho = idx
            break
            
    df_limpo = pd.read_excel(xls, sheet_name=nome_aba, skiprows=linha_cabecalho)
    df_limpo.columns = [str(c).strip() for c in df_limpo.columns]
    return df_limpo

def buscar_coluna_valor(colunas, palavras_chave, ignorar=[]):
    for col in colunas:
        col_texto = str(col).lower()
        if any(ign.lower() in col_texto for ign in ignorar):
            continue
        for palavra in palavras_chave:
            if palavra.lower() in col_texto:
                return col
    return colunas[-1]

def processar_planilhas(arquivo_carregado):
    xls = pd.ExcelFile(arquivo_carregado)
    
    mapa_abas = {}
    for sheet in xls.sheet_names:
        s_upper = sheet.upper()
        if 'ATIVO' in s_upper or 'PRINCIPAL' in s_upper: mapa_abas['ativos'] = sheet
        elif 'CESS' in s_upper: mapa_abas['cessao'] = sheet
        elif 'DEMITIDO' in s_upper: mapa_abas['demitidos'] = sheet
        elif 'WEB' in s_upper: mapa_abas['emp_web'] = sheet
        elif 'EXPER' in s_upper: mapa_abas['experiencia'] = sheet
        elif 'PENS' in s_upper: mapa_abas['pensao'] = sheet
        elif 'ADIPAR' in s_upper: mapa_abas['adipar'] = sheet
        elif 'FGTS' in s_upper: mapa_abas['fgts'] = sheet
        elif 'CONSIGNADO' in s_upper: mapa_abas['consignado'] = sheet

    bases_principais = []
    if 'ativos' in mapa_abas: bases_principais.append(ler_aba_inteligente(xls, mapa_abas['ativos']))
    if 'cessao' in mapa_abas: bases_principais.append(ler_aba_inteligente(xls, mapa_abas['cessao']))
    if 'demitidos' in mapa_abas: bases_principais.append(ler_aba_inteligente(xls, mapa_abas['demitidos']))
    
    if not bases_principais:
        raise ValueError("Nenhuma aba de funcionários (Principal, Cessão ou Demitidos) foi encontrada.")
        
    df_funcionarios = pd.concat(bases_principais, ignore_index=True)
    df_funcionarios = df_funcionarios.drop_duplicates(subset=['Matrícula'])
    df_funcionarios['Matrícula'] = df_funcionarios['Matrícula'].astype(str)
    
    df_res = df_funcionarios.copy()

    if 'experiencia' in mapa_abas:
        df_experiencia = ler_aba_inteligente(xls, mapa_abas['experiencia'])
        df_experiencia['Matrícula'] = df_experiencia['Matrícula'].astype(str)
        df_exp_clean = df_experiencia[['Matrícula', 'Data Fim Experiência', 'Data Fim Experiência/Renovação']].drop_duplicates(subset=['Matrícula'])
        df_res = pd.merge(df_res, df_exp_clean, on='Matrícula', how='left')
    
    if 'fgts' in mapa_abas:
        df_fgts = ler_aba_inteligente(xls, mapa_abas['fgts'])
        df_fgts['Matrícula'] = df_fgts['Matrícula'].astype(str)
        col_valor_fgts = buscar_coluna_valor(df_fgts.columns, ['valor', 'soma', 'fgts'])
        df_fgts_clean = df_fgts[['Matrícula', col_valor_fgts]].rename(columns={col_valor_fgts: 'FGTS Folha'})
        df_res = pd.merge(df_res, df_fgts_clean, on='Matrícula', how='left')

    if 'adipar' in mapa_abas:
        df_adipar = ler_aba_inteligente(xls, mapa_abas['adipar'])
        df_adipar['Matrícula'] = df_adipar['Matrícula'].astype(str)
        col_valor_adipar = buscar_coluna_valor(df_adipar.columns, ['valor', 'soma', 'adipar'])
        df_adipar_clean = df_adipar[['Matrícula', col_valor_adipar]].rename(columns={col_valor_adipar: 'Adipar'})
        df_res = pd.merge(df_res, df_adipar_clean, on='Matrícula', how='left')

    if 'consignado' in mapa_abas:
        df_consignado = ler_aba_inteligente(xls, mapa_abas['consignado'])
        col_mat_consignado = 'matricula' if 'matricula' in df_consignado.columns else 'Matrícula'
        df_consignado[col_mat_consignado] = df_consignado[col_mat_consignado].astype(str)
        col_valor_consignado = buscar_coluna_valor(df_consignado.columns, ['valor', 'soma', 'parcela'])
        df_consignado_clean = df_consignado[[col_mat_consignado, col_valor_consignado]].rename(columns={col_mat_consignado: 'Matrícula', col_valor_consignado: 'eConsignado'})
        df_res = pd.merge(df_res, df_consignado_clean, on='Matrícula', how='left')

    if 'pensao' in mapa_abas:
        df_pensao = ler_aba_inteligente(xls, mapa_abas['pensao'])
        df_pensao['Matrícula'] = df_pensao['Matrícula'].astype(str)
        col_valor_pensao = buscar_coluna_valor(df_pensao.columns, ['base de', 'cálculo', 'calculo', 'valor'])
        df_pensao_clean = df_pensao[['Matrícula', col_valor_pensao]].rename(columns={col_valor_pensao: 'Pensão'}).drop_duplicates(subset=['Matrícula'])
        df_res = pd.merge(df_res, df_pensao_clean, on='Matrícula', how='left')

    if 'emp_web' in mapa_abas:
        df_emp_web = ler_aba_inteligente(xls, mapa_abas['emp_web'])
        col_razao_web = buscar_coluna_valor(df_emp_web.columns, ['razão', 'razao', 'nome', 'social'])
        col_status_web = buscar_coluna_valor(df_emp_web.columns, ['status', 'situação', 'liberado'])
        
        df_res['Nome_Merge'] = df_res.get('Nome Estabelecimento', pd.Series(dtype='str')).astype(str).str.strip().str.upper()
        df_emp_web_clean = df_emp_web[[col_razao_web, col_status_web]].copy()
        df_emp_web_clean['Nome_Merge'] = df_emp_web_clean[col_razao_web].astype(str).str.strip().str.upper()
        df_emp_web_clean = df_emp_web_clean.drop_duplicates(subset=['Nome_Merge'])
        df_emp_web_clean = df_emp_web_clean.rename(columns={col_status_web: 'Liberado no Empregador Web?'})
        
        df_res = pd.merge(df_res, df_emp_web_clean[['Nome_Merge', 'Liberado no Empregador Web?']], on='Nome_Merge', how='left')

    def extrair_data_cessao(texto):
        if pd.isna(texto): return None
        match = re.search(r'\d{2}/\d{2}/\d{4}', str(texto))
        return match.group(0) if match else None
    
    df_res['Cessão_Data'] = df_res.get('Situação do Empregado', pd.Series(dtype='str')).apply(extrair_data_cessao)

    def mapear_codigo_empresa(nome):
        nome_str = str(nome).lower()
        if 'shpx' in nome_str: return '002'
        elif 'shps' in nome_str: return '001'
        elif 'shpp' in nome_str: return '004'
        else: return None

    df_res['Codigo_Chefia_Calculado'] = df_res.get('Nome Estabelecimento', pd.Series(dtype='str')).apply(mapear_codigo_empresa)

    col_sind_cnpj = buscar_coluna_valor(df_res.columns, ['sindicato cnpj', 'cnpj'])
    df_res['CNPJ_Limpo'] = df_res.get(col_sind_cnpj, pd.Series(dtype='str')).astype(str).str.replace(r'\D', '', regex=True)

    dados_finais = {
        'Nome_1': "", 
        'Processo': "",
        'Lote': "",
        'Status': "",
        'Leva': "",
        'DOCS': "",
        'ENVIOS DE DOCS': "",
        'OBS': "",
        'Estab. Nome': df_res.get('Nome Estabelecimento', pd.Series(dtype='str')),
        'Estab. Codigo': df_res.get('Estab.Código', pd.Series(dtype='str')),
        'CNPJ': df_res['CNPJ_Limpo'],
        'Liberado no Empregador Web?': df_res.get('Liberado no Empregador Web?', pd.Series(dtype='str')).fillna('NÃO'),
        'Empresa': df_res.get('Cliente.Nome', pd.Series(dtype='str')),
        'Id Global': df_res.get('Id Global', pd.Series(dtype='str')),
        'Matrícula': df_res.get('Matrícula', pd.Series(dtype='str')),
        'Nome_2': df_res.get('Nome', pd.Series(dtype='str')), 
        'CPF': df_res.get('CPF', pd.Series(dtype='str')),
        'PIS': df_res.get('PIS', pd.Series(dtype='str')),
        'Salário': df_res.get('Salário', pd.Series(dtype='str')),
        'Tipo Funcionário': df_res.get('Tipo Funcionário', pd.Series(dtype='str')),
        '1ª Experiência': df_res.get('Data Fim Experiência', pd.Series(dtype='str')),
        '2ª Experiência': df_res.get('Data Fim Experiência/Renovação', pd.Series(dtype='str')),
        'Adm': df_res.get('Data da Admissão', pd.Series(dtype='str')),
        'Cessão': df_res.get('Cessão_Data', pd.Series(dtype='str')),
        'Competência': "",
        'Data De Deslig': df_res.get('Data de Desligamento', pd.Series(dtype='str')),
        'Data de Crédito': "",
        'Prazo': "",
        'Código Motivo De Rescisão': "",
        'Motivo De Rescisão': df_res.get('Motivo Rescisão', pd.Series(dtype='str')),
        'Aviso Prévio Infos Extras': "",
        'Aviso Prévio Trabalhado (Sim Ou Não)': "",
        'Extrato conectividade + cessão': "",
        'Folha': "",
        'Total': "",
        'Validação Capinha': "",
        'Código Empresa Nova Chefia P/Subordinados': df_res.get('Codigo_Chefia_Calculado', pd.Series(dtype='str')),
        'Matrícula Nova Chefia P/Subordinados': df_res.get('Chefia Imediata - Matrícula', pd.Series(dtype='str')),
        'Data Inicio Do Aviso Prévio': "",
        'Número de Dias do Aviso Prévio': "",
        'Desconta Aviso Prévio (Sim Ou Não)': "",
        'eConsignado': df_res.get('eConsignado', pd.Series(dtype='str')).fillna(0),
        'Manutenção do Ponto': "",
        'Pagamentos Descontos & Outros': "",
        'Desc.\nCopart Care plus \n': "",
        'Desc. Copart Unimed': "",
        'Adipar': df_res.get('Adipar', pd.Series(dtype='str')).fillna(0),
        'Pensão': df_res.get('Pensão', pd.Series(dtype='str')).fillna('Não Possui'),
        'Telefone': "",
        'E-Mail': "",
        'Enviar Mensagem no WhatsApp': "",
        'Status de Envio mensagem no WhatsApp': "",
        'CONF MARCAÇÕES AHGORA': "",
        'CONF ABONO AHGORA': "",
        'DESLIGAMENTO AHGORA': "",
        'VALIDAÇÃO PONTO': "",
        'IMPORT PONTO': "",
        'ANÁLISE PONTO': "",
        'PONTO LIBERADO\nDIA E HORA': "",
        'BENEFÍCIOS': "",
        'HOMOLOGAÇÃO DATA:': "",
        'HOMOLOGAÇÃO STATUS:': ""
    }

    layout_final = pd.DataFrame(dados_finais)

    layout_final.columns = [
        'Nome', 'Processo', 'Lote', 'Status', 'Leva', 'DOCS', 'ENVIOS DE DOCS', 'OBS', 
        'Estab. Nome', 'Estab. Codigo', 'CNPJ', 'Liberado no Empregador Web?', 'Empresa', 
        'Id Global', 'Matrícula', 'Nome', 'CPF', 'PIS', 'Salário', 'Tipo Funcionário', 
        '1ª Experiência', '2ª Experiência', 'Adm', 'Cessão', 'Competência', 'Data De Deslig', 
        'Data de Crédito', 'Prazo', 'Código Motivo De Rescisão', 'Motivo De Rescisão', 
        'Aviso Prévio Infos Extras', 'Aviso Prévio Trabalhado (Sim Ou Não)', 
        'Extrato conectividade + cessão', 'Folha', 'Total', 'Validação Capinha', 
        'Código Empresa Nova Chefia P/Subordinados', 'Matrícula Nova Chefia P/Subordinados', 
        'Data Inicio Do Aviso Prévio', 'Número de Dias do Aviso Prévio', 
        'Desconta Aviso Prévio (Sim Ou Não)', 'eConsignado', 'Manutenção do Ponto', 
        'Pagamentos Descontos & Outros', 'Desc.\nCopart Care plus \n', 'Desc. Copart Unimed', 
        'Adipar', 'Pensão', 'Telefone', 'E-Mail', 'Enviar Mensagem no WhatsApp', 
        'Status de Envio mensagem no WhatsApp', 'CONF MARCAÇÕES AHGORA', 'CONF ABONO AHGORA', 
        'DESLIGAMENTO AHGORA', 'VALIDAÇÃO PONTO', 'IMPORT PONTO', 'ANÁLISE PONTO', 
        'PONTO LIBERADO\nDIA E HORA', 'BENEFÍCIOS', 'HOMOLOGAÇÃO DATA:', 'HOMOLOGAÇÃO STATUS:'
    ]
    
    layout_final = layout_final.fillna("")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        layout_final.to_excel(writer, index=False, sheet_name="Planilha1")
    return output.getvalue()

# --- CORPO DA INTERFACE ---

st.markdown('<h1 class="titulo-premium">✨ Gerador de Planilha</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitulo">Consolidador inteligente de relatórios e layouts.</p>', unsafe_allow_html=True)

with st.expander("📌 Como utilizar a ferramenta", expanded=False):
    st.write("""
    1. Exporte a planilha gerada contendo todas as abas (Ativos, Cessão, Demitidos, etc).
    2. Faça o upload do arquivo com a extensão `.xlsx` na área abaixo.
    3. Clique em **Processar Dados** e aguarde o sistema cruzar todas as chaves (Matrículas).
    4. Faça o download do layout final padronizado com as **62 colunas** necessárias.
    """)

st.divider()

arquivo_carregado = st.file_uploader("📂 Arraste ou selecione o arquivo Excel (.xlsx)", type=["xlsx"])

if arquivo_carregado is not None:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 PROCESSAR DADOS E GERAR LAYOUT", use_container_width=True):
        
        with st.status("⚙️ Analisando abas e cruzando informações...", expanded=True) as status:
            try:
                st.write("Lendo abas principais...")
                st.write("Mapeando cruzamentos de FGTS, Adipar, Consignado e Pensão...")
                dados_excel = processar_planilhas(arquivo_carregado)
                st.write("Formatando 62 colunas do layout de saída...")
                status.update(label="Processamento concluído com sucesso!", state="complete", expanded=False)
                
                st.success("Tudo pronto! Seu layout de 62 colunas foi gerado com perfeição.")
                
                st.download_button(
                    label="📥 BAIXAR PLANILHA CONSOLIDADA",
                    data=dados_excel,
                    file_name="Gerador_Planilha_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                status.update(label="Erro durante o processamento.", state="error", expanded=True)
                st.error(f"Detalhes do erro: {str(e)}")