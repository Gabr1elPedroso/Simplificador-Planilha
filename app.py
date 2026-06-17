import streamlit as st
import pandas as pd
import re
import io

# --- FUNÇÃO DE VERIFICAÇÃO DE LOGIN ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if not st.session_state.password_correct:
        st.markdown("<h1 style='text-align: center; color: #D4AF37;'>🔐 Acesso Restrito</h1>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            user = st.text_input("Utilizador")
            password = st.text_input("Palavra-passe", type="password")
            
            if st.button("Entrar", use_container_width=True):
                if "passwords" in st.secrets and user in st.secrets["passwords"] and st.secrets["passwords"][user] == password:
                    st.session_state.password_correct = True
                    st.rerun()
                else:
                    st.error("Utilizador ou palavra-passe incorretos.")
        return False
    return True

# --- CONFIGURAÇÃO DA PÁGINA E ESTILIZAÇÃO VISUAL PREMIUM ---
st.set_page_config(page_title="Gerador de Planilha", page_icon="✨", layout="centered")

css_personalizado = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;700&display=swap');

    .stApp { background-color: #0d0d0d; font-family: 'Montserrat', sans-serif; }
    html, body, [class*="css"] { font-family: 'Montserrat', sans-serif !important; }

    .titulo-premium {
        background: linear-gradient(45deg, #BF953F, #FCF6BA, #B38728, #FBF5B7, #AA771C);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 700; font-size: 2.8rem; text-align: center; margin-bottom: 10px;
    }
    .subtitulo { text-align: center; color: #A0A0A0; font-weight: 300; font-size: 1.1rem; margin-bottom: 30px; }

    p, span, label, div, li { color: #E5E4E2 !important; }

    [data-testid="stFileUploadDropzone"] {
        background: rgba(26, 26, 26, 0.6); border: 2px dashed rgba(212, 175, 55, 0.4);
        border-radius: 15px; backdrop-filter: blur(5px); transition: all 0.3s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #FFDF00; background: rgba(36, 36, 36, 0.9);
        box-shadow: 0px 0px 20px rgba(212, 175, 55, 0.15);
    }

    /* Estilização da caixa de filtro (TextArea) */
    [data-testid="stTextArea"] textarea {
        background-color: #1a1a1a !important;
        color: #FFDF00 !important;
        border: 1px solid #D4AF37 !important;
        border-radius: 10px !important;
        font-size: 15px !important;
        padding: 15px !important;
    }
    [data-testid="stTextArea"] textarea:focus {
        border-color: #FFDF00 !important;
        box-shadow: 0px 0px 10px rgba(212, 175, 55, 0.4) !important;
        outline: none !important;
    }

    div.stButton > button {
        background: linear-gradient(135deg, #1a1a1a 0%, #0d0d0d 100%) !important;
        color: #D4AF37 !important; font-weight: 600 !important; font-size: 16px !important;
        letter-spacing: 1px; border: 1px solid #D4AF37 !important; border-radius: 10px !important;
        padding: 20px !important; transition: all 0.4s ease;
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #D4AF37 0%, #AA771C 100%) !important;
        color: #0d0d0d !important; box-shadow: 0px 10px 20px rgba(212, 175, 55, 0.4) !important;
        border: 1px solid transparent !important; transform: translateY(-2px);
    }

    div.stDownloadButton > button {
        background: linear-gradient(135deg, #D4AF37 0%, #AA771C 100%) !important;
        color: #0d0d0d !important; font-weight: 700 !important; font-size: 16px !important;
        border: none !important; border-radius: 10px !important; padding: 20px !important;
        box-shadow: 0px 5px 15px rgba(212, 175, 55, 0.3) !important; transition: all 0.3s ease;
    }
    div.stDownloadButton > button:hover {
        background: linear-gradient(135deg, #FFDF00 0%, #D4AF37 100%) !important;
        color: #000000 !important; transform: translateY(-3px);
        box-shadow: 0px 8px 20px rgba(212, 175, 55, 0.5) !important;
    }
    
    [data-testid="stExpander"], [data-testid="stAlert"] {
        background: rgba(26, 26, 26, 0.8) !important; border: 1px solid #D4AF37 !important; border-radius: 10px !important;
    }
    [data-testid="stAlert"] p, [data-testid="stAlert"] span, [data-testid="stAlert"] div {
        color: #FFDF00 !important; font-weight: 600;
    }
    [data-testid="stAlert"] svg { fill: #FFDF00 !important; }
    hr { border-top: 1px dashed rgba(212, 175, 55, 0.3); }
</style>
"""
st.markdown(css_personalizado, unsafe_allow_html=True)

# BLOQUEIO DE ACESSO
if not check_password():
    st.stop()

# --- FUNÇÕES DE PROCESSAMENTO ---
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

def processar_planilhas(arquivo_carregado, lista_matriculas=None):
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
    df_funcionarios['Matrícula'] = df_funcionarios['Matrícula'].astype(str).str.strip()
    
    df_res = df_funcionarios.copy()

    if 'experiencia' in mapa_abas:
        df_experiencia = ler_aba_inteligente(xls, mapa_abas['experiencia'])
        df_experiencia['Matrícula'] = df_experiencia['Matrícula'].astype(str).str.strip()
        df_exp_clean = df_experiencia[['Matrícula', 'Data Fim Experiência', 'Data Fim Experiência/Renovação']].drop_duplicates(subset=['Matrícula'])
        df_res = pd.merge(df_res, df_exp_clean, on='Matrícula', how='left')
    
    if 'fgts' in mapa_abas:
        df_fgts = ler_aba_inteligente(xls, mapa_abas['fgts'])
        df_fgts['Matrícula'] = df_fgts['Matrícula'].astype(str).str.strip()
        col_valor_fgts = buscar_coluna_valor(df_fgts.columns, ['valor', 'soma', 'fgts'])
        df_fgts_clean = df_fgts[['Matrícula', col_valor_fgts]].rename(columns={col_valor_fgts: 'FGTS Folha'})
        df_res = pd.merge(df_res, df_fgts_clean, on='Matrícula', how='left')

    if 'adipar' in mapa_abas:
        df_adipar = ler_aba_inteligente(xls, mapa_abas['adipar'])
        df_adipar['Matrícula'] = df_adipar['Matrícula'].astype(str).str.strip()
        col_valor_adipar = buscar_coluna_valor(df_adipar.columns, ['valor', 'soma', 'adipar'])
        df_adipar_clean = df_adipar[['Matrícula', col_valor_adipar]].rename(columns={col_valor_adipar: 'Adipar'})
        df_res = pd.merge(df_res, df_adipar_clean, on='Matrícula', how='left')

    if 'consignado' in mapa_abas:
        df_consignado = ler_aba_inteligente(xls, mapa_abas['consignado'])
        col_mat_consignado = 'matricula' if 'matricula' in df_consignado.columns else 'Matrícula'
        df_consignado[col_mat_consignado] = df_consignado[col_mat_consignado].astype(str).str.strip()
        col_valor_consignado = buscar_coluna_valor(df_consignado.columns, ['valor', 'soma', 'parcela'])
        df_consignado_clean = df_consignado[[col_mat_consignado, col_valor_consignado]].rename(columns={col_mat_consignado: 'Matrícula', col_valor_consignado: 'eConsignado'})
        df_res = pd.merge(df_res, df_consignado_clean, on='Matrícula', how='left')

    if 'pensao' in mapa_abas:
        df_pensao = ler_aba_inteligente(xls, mapa_abas['pensao'])
        df_pensao['Matrícula'] = df_pensao['Matrícula'].astype(str).str.strip()
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
    
    # --- APLICAÇÃO DO FILTRO DE MATRÍCULAS ---
    if lista_matriculas:
        layout_final['Matrícula'] = layout_final['Matrícula'].astype(str).str.strip()
        layout_final = layout_final[layout_final['Matrícula'].isin(lista_matriculas)]
        
        if layout_final.empty:
            raise ValueError("Nenhum funcionário encontrado com as matrículas informadas.")
            
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
    1. Faça o upload do ficheiro Excel com a base bruta.
    2. (Opcional) Cole a lista de Matrículas no campo de filtro abaixo para exportar apenas essas pessoas.
    3. Clique em **Processar Dados** para gerar o layout unificado com as 62 colunas.
    """)

st.divider()

# Upload do Arquivo SEMPRE visível
arquivo_carregado = st.file_uploader("📂 Arraste ou selecione o ficheiro Excel (.xlsx)", type=["xlsx"])

# Linha divisória e Filtro SEMPRE visíveis
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("<h3 style='color: #D4AF37; font-size: 1.5rem; text-align: center; margin-bottom: 5px;'>🔍 Filtro Específico (Opcional)</h3>", unsafe_allow_html=True)

# A caixa de texto (sem label do Streamlit, pois já temos o título em HTML acima)
matriculas_input = st.text_area(
    "Filtro oculto", 
    label_visibility="collapsed",
    placeholder="Cole a lista de matrículas aqui (uma por linha ou separadas por vírgula).\n\n⚠️ Deixe em branco se desejar processar TODOS os funcionários.",
    height=130
)

# O botão de processar só aparece (ou só executa) se houver um ficheiro
if arquivo_carregado is not None:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 PROCESSAR DADOS E GERAR LAYOUT", use_container_width=True):
        
        lista_matriculas = []
        if matriculas_input.strip():
            raw_list = re.split(r'[,\n;\s]+', matriculas_input)
            lista_matriculas = [m.strip() for m in raw_list if m.strip()]
            
        with st.status("⚙️ A analisar abas e a cruzar informações...", expanded=True) as status:
            try:
                if lista_matriculas:
                    st.write(f"A aplicar filtro para {len(lista_matriculas)} matrícula(s)...")
                else:
                    st.write("A ler a base completa de funcionários...")
                    
                st.write("A mapear os cruzamentos de FGTS, Adipar, Consignado e Pensão...")
                
                dados_excel = processar_planilhas(arquivo_carregado, lista_matriculas)
                
                st.write("A formatar as 62 colunas do layout de saída...")
                status.update(label="Processamento concluído com sucesso!", state="complete", expanded=False)
                
                if lista_matriculas:
                    st.success("Layout filtrado gerado na perfeição!")
                else:
                    st.success("Tudo pronto! O layout completo de 62 colunas foi gerado na perfeição.")
                
                st.download_button(
                    label="📥 BAIXAR PLANILHA CONSOLIDADA",
                    data=dados_excel,
                    file_name="Gerador_Planilha_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except ValueError as ve:
                status.update(label="Aviso", state="error", expanded=True)
                st.warning(str(ve))
            except Exception as e:
                status.update(label="Erro durante o processamento.", state="error", expanded=True)
                st.error(f"Detalhes do erro: {str(e)}")