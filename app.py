import streamlit as st
import pandas as pd
import io
import requests
import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import firebase_admin
from firebase_admin import credentials, auth
from datetime import datetime, timezone, timedelta

# =========================================================================
# 1. CONFIGURAÇÕES TÉCNICAS DA PÁGINA & INICIALIZAÇÃO SEGURA DO FIREBASE
# =========================================================================

st.set_page_config(
    page_title="Gerador de Fichas Jurídicas - VCB Senado",
    page_icon="logo_bibliokhan.ico",
    layout="wide"
)

# --- ADICIONAR A LOGO NA BARRA LATERAL ---
st.sidebar.image("logo_bibliokhan.png", use_container_width=True)

# --- BARRA LATERAL (Tudo encostado na esquerda) ---
with st.sidebar:
    st.title("**BiblioKhan**")
    st.write("**Inteligência e Automação para Bibliotecas**")
    st.write("bibliokhancontato@gmail.com")
    st.markdown("---")

if not firebase_admin._apps:
    try:
        firebase_secrets = dict(st.secrets["firebase"])
        firebase_secrets["private_key"] = firebase_secrets["private_key"].replace("\\n", "\n")
        
        cred = credentials.Certificate(firebase_secrets)
        firebase_admin.initialize_app(cred)
    except Exception as e:
        st.error(f"❌ Erro crítico nas credenciais do Firebase: {str(e)}")

# =========================================================================
# 🌟 RECARGA AUTOMÁTICA EM BACKEND
# =========================================================================
def tratar_url_google_sheets(url):
    """
    Transforma o link padrão do Google Sheets em exportação direta de CSV
    e adiciona um parâmetro de tempo para forçar o Sheets a quebrar o cache.
    """
    url = url.strip()
    
    # Remove parâmetros extras do final do link se houver
    if "?" in url and not "docs.google.com" in url:
        url = url.split("?")[0]
        
    if "/edit" in url:
        url = url.split("/edit")[0] + "/export?format=csv"
    elif "/pubhtml" in url:
        url = url.split("/pubhtml")[0] + "/pub?output=csv"
    elif not url.endswith("/export?format=csv") and "docs.google.com" in url:
        if url.endswith("/"):
            url = url + "export?format=csv"
        else:
            url = url + "/export?format=csv"
            
    # 🔥 TRUQUE DO CACHE: Adiciona a hora atual em segundos no link.
    # Isso força o Google a gerar um CSV idêntico ao que está na tela agora.
    import time
    nocache_param = f"&nocache={int(time.time())}"
    url += nocache_param
    
    return url

def carregar_creditos_planilha(url_planilha):
    """
    Lê a planilha usando o link tratado e retorna o DataFrame.
    """
    try:
        url_tratada = tratar_url_google_sheets(url_planilha)
        df = pd.read_csv(url_tratada)
        return df
    except Exception as e:
        st.error(f"Erro ao acessar os dados da planilha: {e}")
        return None

def atualizar_saldo_usuario(email_usuario):
    try:
        url_planilha = st.secrets["URL_PLANILHA"]
        df = carregar_creditos_planilha(url_planilha)
        
        if df is not None:
            # --- DIAGNÓSTICO 1: Mostrar o que o Pandas leu ---
            st.warning(f"🔍 Colunas encontradas na planilha: {list(df.columns)}")
            
            # Normalizar os nomes das colunas para evitar erros de maiúsculas/minúsculas
            df.columns = df.columns.str.strip().str.lower()
            
            if 'token' in df.columns and 'creditos' in df.columns:
                # Limpar espaços e converter para maiúsculas para comparar com segurança
                df['token'] = df['token'].astype(str).str.strip().str.upper()
                email_chave = email_usuario.strip().upper()
                
                # --- DIAGNÓSTICO 2: Mostrar lista de e-mails cadastrados ---
                lista_emails_planilha = df['token'].tolist()
                st.write(f"📧 Tentando procurar por: `{email_chave}`")
                st.write(f"📋 Lista de e-mails lidos na planilha: {lista_emails_planilha}")
                
                if email_chave in df['token'].values:
                    # Captura o saldo garantindo que é um número inteiro
                    saldo = int(df.loc[df['token'] == email_chave, 'creditos'].values[0])
                    st.session_state["creditos_ativos"] = saldo
                    st.success(f"✅ Utilizador encontrado! Saldo atualizado para: {saldo}")
                else:
                    st.error("❌ O e-mail de login NÃO foi encontrado na coluna 'token' da planilha.")
                    st.session_state["creditos_ativos"] = 0
            else:
                st.error("❌ Erro crítico: A planilha precisa de ter as colunas com os nomes exatos: 'token' e 'creditos'.")
                st.session_state["creditos_ativos"] = 0
        else:
            st.error("❌ Erro crítico: O Pandas não conseguiu ler nenhum dado da URL fornecida.")
            st.session_state["creditos_ativos"] = 0
            
    except Exception as e:
        st.error(f"❌ Erro na sincronização de saldo: {e}")
        st.session_state["creditos_ativos"] = 0

# =========================================================================
# 2. SISTEMA DE AUTENTICAÇÃO E CONTROLE DE SESSÃO COMERCIAL
# =========================================================================

def verificar_login_firebase(email, senha):
    try:
        user = auth.get_user_by_email(email)
        st.session_state["logado"] = True
        st.session_state["usuario_atual"] = user.email
        atualizar_saldo_usuario(user.email)
        return True
    except Exception as e:
        st.error("❌ Acesso negado: E-mail não cadastrado ou credenciais inválidas.")
        return False

if "logado" not in st.session_state:
    st.session_state["logado"] = False

if "creditos_ativos" not in st.session_state:
    st.session_state["creditos_ativos"] = 0

# Exibe o saldo na barra lateral caso o usuário esteja logado
if st.session_state["logado"]:
    with st.sidebar:
        if st.session_state["creditos_ativos"] > 0:
            st.success(f"💳 Saldo: {st.session_state['creditos_ativos']} fichas")
        else:
            st.error("💳 Sem créditos ativos")

# =========================================================================
# 3. INTERFACE DE LOGIN OU FLUXO DO APLICATIVO PROTEGIDO
# =========================================================================

if not st.session_state["logado"]:
    st.markdown("# 🔒 Área do Cliente")
    st.markdown("### Faça o login para acessar o Gerador de Fichas Jurídicas.")
    
    with st.form("login_form"):
        email_input = st.text_input("E-mail de Usuário").strip()
        senha_input = st.text_input("Senha de Acesso", type="password").strip()
        botao_entrar = st.form_submit_button("Entrar no Sistema")
        
        if botao_entrar:
            if email_input and senha_input:
                verificar_login_firebase(email_input, senha_input)
                if st.session_state["logado"]:
                    st.rerun()
            else:
                st.warning("⚠️ Por favor, preencha o e-mail e a senha.")

    st.markdown("---")
    with st.expander("🔑 Esqueceu sua senha ou quer trocar a senha provisória?"):
        st.markdown("""
        Como medida de segurança, a alteração de credenciais é validada diretamente pela administração.
        
        Para redefinir sua senha, entre em contato diretamente com o suporte técnico através do e-mail informado na lateral do sistema ou pelo canal de atendimento onde adquiriu o produto. Um link oficial de redefinição será enviado para o seu e-mail cadastrado.
        """)

else:
    # --- CONTEÚDO DO APLICATIVO COMERCIAL ---
    st.markdown("""
        <style>
        textarea {
            font-family: 'Courier New', Courier, monospace !important;
        }
        .stTabs [data-baseweb="tab-list"] { gap: 24px; }
        .stTabs [data-baseweb="tab"] { 
            height: 50px; 
            white-space: pre-wrap; 
            background-color: #f0f2f6; 
            border-radius: 5px 5px 0px 0px; 
            gap: 1px; 
            padding-top: 10px; 
            padding-bottom: 10px; 
        }
        .stTabs [aria-selected="true"] { background-color: #B19FFB !important; color: black !important; font-weight: bold; }
        </style>
        """, unsafe_allow_html=True)

    if "lote_fichas" not in st.session_state:
        st.session_state.lote_fichas = []

    if "assuntos_selecionados" not in st.session_state:
        st.session_state.assuntos_selecionados = []

    def buscar_vcb_senado(termo_busca):
        url_api = "https://adm.senado.leg.br/vcb/vocab/services.php"
        params = {"task": "search", "arg": termo_busca, "output": "json"}
        try:
            resposta = requests.get(url_api, params=params, timeout=8, verify=False)
            if resposta.status_code == 200:
                dados = resposta.json()
                resultados_formatados = []
                bloco_result = dados.get("result", {})
                if isinstance(bloco_result, dict):
                    for chave, item in bloco_result.items():
                        if isinstance(item, dict) and "string" in item:
                            resultados_formatados.append({
                                "termo": item["string"].strip(),
                                "id": f"VCB-{item.get('term_id', chave)}",
                                "note": "Termo oficial homologado pelo Vocabulário Controlado do Senado Federal."
                            })
                return resultados_formatados
        except Exception:
            return []
        return []

    def gerar_docx_lote(lista_fichas):
        doc = Document()
        
        # Configuração das margens
        section = doc.sections[0]
        section.left_margin = Pt(72)
        section.right_margin = Pt(72)

        for idx, ficha_texto in enumerate(lista_fichas):
            if idx > 0:
                doc.add_page_break()
            
            # Adiciona a tabela com o estilo 'Table Grid' que força as bordas
            table = doc.add_table(rows=1, cols=1)
            table.style = 'Table Grid' 
            table.autofit = False
            table.allow_autofit = False
            table.columns[0].width = Pt(400)
            
            # Acessa a célula
            cell = table.cell(0, 0)
            
            # Remove parágrafos padrão para garantir controle total
            cell._element.clear_content()
            
            # Adiciona o texto configurando a fonte
            p = cell.add_paragraph()
            run = p.add_run(ficha_texto)
            run.font.name = 'Courier New'
            run.font.size = Pt(10)
            
            # Ajusta o alinhamento e recuo dentro da caixa
            p.paragraph_format.left_indent = Pt(10)
            p.paragraph_format.right_indent = Pt(10)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(10)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    def formatar_entrada_e_corpo(tipo_autor, autores_lista, entidade, titulo, tem_organizador, organizador_nome, tipo_org, tem_tradutor, tradutor_nome):
        entrada = ""
        corpo_autores = ""
        entrada_por_titulo = False
        
        if tem_organizador and tipo_autor == "Pessoa Física" and not any(a.strip() for a in autores_lista):
            entrada_por_titulo = True
            entrada = ""
            corpo_autores = f"{tipo_org} por {organizador_nome.strip()}"
        elif tipo_autor == "Entidade (Órgão/Instituição)":
            entrada = entidade.strip().upper()
            corpo_autores = ""
        else:
            autores = [a.strip() for a in autores_lista if a.strip()]
            qtd = len(autores)
            
            if qtd == 1:
                partes = autores[0].split()
                entrada = f"{partes[-1].upper()}, {' '.join(partes[:-1])}." if len(partes) > 1 else f"{autores[0].upper()}."
                corpo_autores = autores[0]
            elif qtd >= 2 and qtd <= 3:
                partes = autores[0].split()
                entrada = f"{partes[-1].upper()}, {' '.join(partes[:-1])}." if len(partes) > 1 else f"{autores[0].upper()}."
                corpo_autores = ", ".join(autores)
            elif qtd >= 4:
                entrada_por_titulo = True
                entrada = ""  
                corpo_autores = f"{autores[0]} [et al.]"
                
            if tem_organizador and organizador_nome.strip() and qtd < 4:
                corpo_autores += f" ; {tipo_org} por {organizador_nome.strip()}"

        if tem_tradutor and tradutor_nome.strip():
            if corpo_autores:
                corpo_autores += f" ; tradução por {tradutor_nome.strip()}"
            else:
                corpo_autores = f"tradução por {tradutor_nome.strip()}"
                
        return entrada, corpo_autores, entrada_por_titulo

    def buscar_na_tabela_cutter(texto_para_busca, titulo_obra):
        if not texto_para_busca or not titulo_obra: return "X000x"
        url_csv = "https://raw.githubusercontent.com/Biblio-Khan/gerador-ficha-cat/refs/heads/main/cutter.csv"
        try:
            df = pd.read_csv(url_csv, sep=',', encoding='utf-8', quotechar='"')
        except Exception:
            return f"{texto_para_busca.strip().upper()[0]}200{titulo_obra.strip().lower()[0]}"
        
        df.columns = df.columns.str.strip().str.lower()
        col_nome = 'name' if 'name' in df.columns else df.columns[0]
        col_id = 'id' if 'id' in df.columns else df.columns[1]
        
        df['Name_Clean'] = df[col_nome].astype(str).str.strip().str.upper()
        sub_busca = texto_para_busca.strip().upper()
        
        match = df[df['Name_Clean'] <= sub_busca].sort_values(by='Name_Clean').tail(1)
        num = "200"
        if not match.empty:
            num = str(match[col_id].values[0]).strip().split('.')[0]
            
        titulo_limpo = titulo_obra.strip().upper()
        artigos = ["O ", "A ", "OS ", "AS ", "UM ", "UMA ", "UNS ", "UMAS "]
        for artigo in artigos:
            if titulo_limpo.startswith(artigo):
                titulo_limpo = titulo_limpo[len(artigo):].strip()
                break
        letra_titulo = titulo_limpo[0].lower() if titulo_limpo else "t"
        return f"{sub_busca[0]}{num}{letra_titulo}"

    def calcular_cutter(tipo_autor, autores_lista, entidade="", titulo="", tem_organizador=False, organizador_nome=""):
        if tipo_autor == "Entidade (Órgão/Instituição)" and entidade:
            texto_base = entidade
        elif tipo_autor == "Pessoa Física" and autores_lista and any(a.strip() for a in autores_lista):
            autor_principal = [a.strip() for a in autores_lista if a.strip()][0]
            partes = autor_principal.split()
            texto_base = partes[-1] if len(partes) > 1 else autor_principal
        elif tem_organizador or tipo_autor == "Organizador":
            partes_org = organizador_nome.strip().split()
            texto_base = partes_org[-1] if len(partes_org) > 1 else organizador_nome
        else:
            texto_base = "Autor"
        return buscar_na_tabela_cutter(texto_base, titulo)

    # =========================================================================
    # SISTEMA DE ABAS (CATALOGAÇÃO & CRÉDITOS LIMITADOS ATÉ 300)
    # =========================================================================
    tab_gerador, tab_financeiro = st.tabs(["⚖️ Gerar Ficha", "💳 Compra e Gestão de Créditos"])

    with tab_gerador:
        if st.session_state["creditos_ativos"] <= 0:
            st.warning("🔒 O painel de salvamento está bloqueado. Adquira créditos ou aguarde a restauração para continuar.")

        st.title("⚖️ Gerador de Fichas Jurídicas — NBR/AACR2")
        st.caption("Mesa técnica integrada via Web Service ao Vocabulário Controlado Básico (VCB) do Senado Federal.")

        st.markdown("---")
        container_lote = st.container()
        with container_lote:
            col_lote_1, col_lote_2 = st.columns([2, 1])
            qtd_fichas = len(st.session_state.lote_fichas)
            col_lote_1.subheader(f"📦 Lote de Trabalho Atual: {qtd_fichas} Ficha(s) Acumulada(s)")
            
            if qtd_fichas > 0:
                arquivo_word = gerar_docx_lote(st.session_state.lote_fichas)
                col_lote_2.download_button(
                    label="📥 Baixar Lote Completo (.DOCX / Word)",
                    data=arquivo_word,
                    file_name="lote_fichas_aacr2.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                if col_lote_2.button("🗑️ Limpar Lote"):
                    st.session_state.lote_fichas = []
                    st.rerun()
            else:
                col_lote_2.info("O lote está vazio. Conclua uma ficha abaixo.")

        st.markdown("---")
        col_esquerda, col_direita = st.columns(2)

        with col_esquerda:
            st.subheader("1. Metadados & Responsabilidade")
            classificacao = st.text_input("Número de Classificação (CDD ou CDU)", value="340.1")
            tipo_autor = st.radio("Tipo de Autoria Principal", ["Pessoa Física", "Entidade (Órgão/Instituição)"], horizontal=True)
            
            autores_lista = []
            entidade_nome = ""
            
            if tipo_autor == "Pessoa Física":
                qtd_autores_input = st.number_input("Quantidade de autores principais (0 se houver apenas Organizador)", min_value=0, max_value=10, value=1)
                for i in range(int(qtd_autores_input)):
                    autores_lista.append(st.text_input(f"Autor {i+1} (Nome Sobrenome)", key=f"autor_{i}"))
            else:
                entidade_nome = st.text_input("Nome da Entidade (Ex: Brasil. Supremo Tribunal Federal)")
                
            titulo = st.text_input("Título Principal")
            st.markdown("---")
            col_resp_1, col_resp_2 = st.columns(2)
            
            with col_resp_1:
                tem_organizador = st.checkbox("Possui Organizador/Coordenador?")
                organizador_nome = ""
                tipo_org, abreviatura_org = "", ""
                if tem_organizador:
                    papel = st.selectbox("Função:", ["Organizador", "Coordenador", "Compilador"])
                    organizador_nome = st.text_input("Nome do Responsável")
                    if papel == "Organizador": tipo_org, abreviatura_org = "organizado", "org."
                    elif papel == "Coordenador": tipo_org, abreviatura_org = "coordenado", "coord."
                    else: tipo_org, abreviatura_org = "compilado", "comp."
                    
            with col_resp_2:
                tem_tradutor = st.checkbox("A obra possui Tradutor?")
                tradutor_nome = ""
                if tem_tradutor:
                    tradutor_nome = st.text_input("Nome do Tradutor (Nome Sobrenome)", key="trad_nome")

            st.markdown("---")
            st.subheader("2. Publicação & Descrição Física")
            edicao = st.text_input("Edição (Ex: 2. ed., 3. ed. rev. e ampl.)", value="1. ed.")
            editora = st.text_input("Editora")
            cidade = st.text_input("Cidade de Publicação", value="Brasília")
            ano = st.text_input("Ano de Publicação", value="2026")
            paginas = st.text_input("Número de Páginas/Folhas", value="180")
            
            tem_colecao = st.checkbox("Esta obra faz parte de uma Coleção / Série?")
            colecao_nome = ""
            if tem_colecao:
                colecao_nome = st.text_input("Nome da Coleção e Volume (Ex: Biblioteca jurídica, v. 12)")
                
            isbn = st.text_input("ISBN (Ex: 978-65-0000-00-0)")
            suporte = st.radio("Suporte da Obra", ["Impresso", "Digital"], horizontal=True)
            url_acesso = st.text_input("URL de Acesso / DOI") if suporte == "Digital" else ""

        with col_direita:
            st.subheader("3. Indexação por Assunto")
            st.markdown("##### 🏛️ Buscar no VCB do Senado Federal")
            termo_busca = st.text_input("Digite um termo jurídico para pesquisar:")
            
            if termo_busca:
                resultados_vcb = buscar_vcb_senado(termo_busca)
                if resultados_vcb:
                    st.success(f"{len(resultados_vcb)} conceitos localizados no Senado!")
                    mapeamento_opcoes = {item["termo"]: item for item in resultados_vcb}
                    lista_opcoes = sorted(list(mapeamento_opcoes.keys()))
                    termo_selecionado = st.selectbox("Selecione o conceito oficial:", lista_opcoes)
                    
                    if st.button("➕ Vincular Assunto do Senado"):
                        if termo_selecionado not in st.session_state.assuntos_selecionados:
                            st.session_state.assuntos_selecionados.append(termo_selecionado)
                            st.rerun()
                else:
                    st.warning("Nenhum termo correspondente retornado pela API do Senado.")

            st.markdown("##### ✍️ Adicionar Assunto Manualmente")
            assunto_manual = st.text_input("Digite um assunto customizado:")
            if st.button("➕ Vincular Assunto Manual"):
                if assunto_manual.strip():
                    termo_limpo = assunto_manual.strip()
                    if termo_limpo not in st.session_state.assuntos_selecionados:
                        st.session_state.assuntos_selecionados.append(termo_limpo)
                        st.rerun()

            if st.session_state.assuntos_selecionados:
                st.write("**Assuntos Vinculados à Ficha:**")
                
                # Criamos um botão de exclusão individual para cada assunto
                for idx, ass in enumerate(st.session_state.assuntos_selecionados):
                    col_assunto, col_excluir = st.columns([9, 1])
                    with col_assunto:
                        st.write(f"{idx+1}. {ass}")
                    with col_excluir:
                        if st.button("❌", key=f"remover_assunto_{idx}", help="Remover apenas este assunto"):
                            st.session_state.assuntos_selecionados.pop(idx)
                            st.rerun()
                            
                # Mantemos o botão de limpar tudo, caso o usuário queira zerar a lista
                if st.button("🗑️ Limpar Todos os Assuntos"):
                    st.session_state.assuntos_selecionados = []
                    st.rerun()

            st.markdown("---")
            st.subheader("4. Fechamento e Visualização da Ficha")
            
            entrada_principal, responsabilidade, entrada_por_titulo = formatar_entrada_e_corpo(
                tipo_autor=tipo_autor, autores_lista=autores_lista, entidade=entidade_nome, titulo=titulo, 
                tem_organizador=tem_organizador, organizador_nome=organizador_nome, tipo_org=tipo_org, 
                tem_tradutor=tem_tradutor, tradutor_nome=tradutor_nome
            )
            
            cutter = calcular_cutter(tipo_autor, autores_lista, entidade=entidade_nome, titulo=titulo, tem_organizador=tem_organizador, organizador_nome=organizador_nome)
            dgm = " [recurso eletrônico]" if suporte == "Digital" else ""
            desc_fisica = f"1 recurso online ({paginas} f.) " if suporte == "Digital" else f"{paginas} f"
            
            bloco_colecao = ""
            if tem_colecao and colecao_nome.strip():
                text_colecao = colecao_nome.strip()
                text_colecao = text_colecao[0].upper() + text_colecao[1:]
                bloco_colecao = f" ({text_colecao})"
                
            nota_acesso = f"\n            Modo de acesso: {url_acesso}" if suporte == "Digital" and url_acesso else ""
            isbn_bloco = f"\n            ISBN {isbn}" if isbn.strip() else ""
            nota_traducao = f"\n            Traduzido de obra original." if tem_tradutor and tradutor_nome.strip() else ""
            ed_bloco = f"{edicao.strip()} – " if edicao.strip() else ""
            pub_bloco = f"{cidade.strip()} : {editora.strip()}, {ano.strip()}."
            
            string_assuntos = " ".join([f"{i+1}. {ass}" for i, ass in enumerate(st.session_state.assuntos_selecionados)])
            rastreabilidade = ""
            romanos = ["I", "II", "III", "IV", "V"]
            r_idx = 0
            
            if not entrada_por_titulo:
                rastreabilidade += f" {romanos[r_idx]}. Título."
                r_idx += 1
                
            if tem_organizador and organizador_nome.strip():
                partes_org = organizador_nome.strip().split()
                nome_invertido_org = f"{partes_org[-1].upper()}, {' '.join(partes_org[:-1])}" if len(partes_org) > 1 else organizador_nome.strip().upper()
                rastreabilidade += f" {romanos[r_idx]}. {nome_invertido_org}, {abreviatura_org}."
                r_idx += 1
                
            if tem_tradutor and tradutor_nome.strip():
                partes_trad = tradutor_nome.strip().split()
                nome_invertido_trad = f"{partes_trad[-1].upper()}, {' '.join(partes_trad[:-1])}" if len(partes_trad) > 1 else tradutor_nome.strip().upper()
                rastreabilidade += f" {romanos[r_idx]}. {nome_invertido_trad}, trad."
                r_idx += 1

            if entrada_por_titulo:
                txt_ficha = f"""{classificacao}
{cutter}   {titulo.strip()}{dgm} / {responsabilidade}. – {ed_bloco}{pub_bloco}
            {desc_fisica}.{bloco_colecao}{nota_traducao}{nota_acesso}{isbn_bloco}
            
            {string_assuntos}{rastreabilidade}"""
            else:
                txt_ficha = f"""{classificacao}
{cutter}   {entrada_principal}
            {titulo.strip()}{dgm} / {responsabilidade}. – {ed_bloco}{pub_bloco}
            {desc_fisica}.{bloco_colecao}{nota_traducao}{nota_acesso}{isbn_bloco}
            
            {string_assuntos}{rastreabilidade}"""
                    
            st.text_area("Visualização Normativa (Fonte Monoespaçada)", value=txt_ficha, height=240)
            
            if st.button("💾 CONCLUIR FICHA E ENVIAR AO LOTE", disabled=st.session_state["creditos_ativos"] <= 0):
                valido = True
                if tipo_autor == "Pessoa Física" and not any(a.strip() for a in autores_lista) and not tem_organizador:
                    valido = False
                    
                if valido and titulo.strip():
                    with st.spinner("Gravando ficha e atualizando saldo na nuvem..."):
                        try:
                            # 1. Envia a ordem de desconto para o Google Apps Script da Planilha
                            url_script = st.secrets["URL_SCRIPT_GOOGLE"]
                            payload = {
                                "email": st.session_state["usuario_atual"],
                                "acao": "descontar"
                            }
                            # Faz a requisição POST para rodar o script do Google
                            resposta_google = requests.post(url_script, json=payload, timeout=15)
                            
                            # Verificação de segurança: O Google respondeu com sucesso HTTP (200)?
                            if resposta_google.status_code == 200:
                                try:
                                    resultado_json = resposta_google.json()
                                    if resultado_json.get("status") == "sucesso":
                                        # 2. Se correu bem na planilha, atualiza localmente e insere no lote
                                        st.session_state.lote_fichas.append(txt_ficha)
                                        st.session_state["creditos_ativos"] -= 1
                                        st.session_state.assuntos_selecionados = [] 
                                        st.success("✅ Ficha guardada com sucesso! Saldo deduzido diretamente na planilha.")
                                        st.rerun()
                                    else:
                                        erro_msg = resultado_json.get("mensagem", "Erro desconhecido")
                                        st.error(f"❌ Não foi possível deduzir o saldo na planilha: {erro_msg}")
                                except Exception:
                                    # Captura se o script do Google mandou HTML de erro ou texto bruto em vez de JSON
                                    st.error(f"❌ O Google Script não retornou um formato JSON válido. Resposta recebida do servidor: {resposta_google.text[:250]}")
                            else:
                                st.error(f"❌ Falha crítica de conexão com o servidor Google Script (Status HTTP: {resposta_google.status_code})")
                                
                        except Exception as e:
                            st.error(f"❌ Erro de comunicação com a planilha: {e}")
                else:
                    st.error("Preencha os campos de autoria/organização e o título.")

    with tab_financeiro:
        st.header("💳 Gestão Financeira e Saldo")
        col_f1, col_f2 = st.columns(2)
        
        with col_f1:
            st.subheader("🔄 Sincronização")
            st.info(f"Seu sistema está vinculado ao e-mail: **{st.session_state['usuario_atual']}**")
            if st.button("Atualizar meu Saldo"):
                with st.spinner("Puxando dados atualizados do Sheets..."):
                    atualizar_saldo_usuario(st.session_state["usuario_atual"])
                    st.success("Saldo checado com sucesso!")
                    st.rerun()

        with col_f2:
            st.subheader("🛒 Tabela de Preços")
            st.markdown("""
            * **30 Fichas** — R$ 70,00 
            * **60 Fichas** — R$ 160,00 
            * **100 Fichas** — R$ 240,00 
            * **200 Fichas** — R$ 420,00 
            * **300 Fichas** — R$ 570,00 
            """)
            st.info("🔑 **PIX:** `bibliokhancontato@gmail.com`")

        st.markdown("---")
        st.subheader("📩 Envio de Comprovante")
        
        with st.form("pix_form_original"):
            email_cliente = st.text_input("E-mail de Cadastro no Sistema", value=st.session_state["usuario_atual"], disabled=True)
           
            pacote_escolhido = st.selectbox(
                "Qual pacote de créditos você comprou?",
                options=[
                    "30 Fichas (R$ 70,00)",
                    "60 Fichas (R$ 160,00)",
                    "100 Fichas (R$ 240,00)",
                    "200 Fichas (R$ 420,00)",
                    "300 Fichas (R$ 570,00)"
                ]
            )
            
            comprovante = st.file_uploader("Anexe a imagem ou PDF do comprovante do PIX", type=["jpg", "png", "jpeg", "pdf"])
            
            if st.form_submit_button("Enviar para Restauração de Saldo"):
                if comprovante is not None:
                    with st.spinner("Enviando comprovante para o suporte... Por favor, aguarde."):
                        try:
                            tg_token = st.secrets["TELEGRAM_BOT_TOKEN"]
                            tg_chat = st.secrets["TELEGRAM_CHAT_ID"]
    
                            fuso_brasilia = timezone(timedelta(hours=-3))
                            data_hora_br = datetime.now(fuso_brasilia).strftime('%d/%m/%Y %H:%M:%S')
                            
                            texto_notificacao = (
                                f"🔥 *NOVO COMPROVANTE RECEBIDO!*\n\n"
                                f"📧 *E-mail do Cliente:* {st.session_state['usuario_atual']}\n"
                                f"💰 *Pacote Escolhido:* {pacote_escolhido}\n"
                                f"📅 *Data/Hora:* {data_hora_br}"
                            )
                            
                            url_api_telegram = f"https://api.telegram.org/bot{tg_token}/sendPhoto"
                            ficheiro_envio = {"photo": (comprovante.name, comprovante.getvalue(), comprovante.type)}
                            dados_requisicao = {"chat_id": tg_chat, "caption": texto_notificacao, "parse_mode": "Markdown"}
                            
                            resposta_tg = requests.post(url_api_telegram, data=dados_requisicao, files=ficheiro_envio, timeout=15)
                            
                            if resposta_tg.status_code == 200:
                                st.success("✅ Comprovante enviado com sucesso!")
                                st.info("⏳ O seu saldo será atualizado assim que a validação for concluída.")
                            else:
                                st.error(f"Erro na API de comunicação (Código {resposta_tg.status_code}).")
                        except Exception as e:
                            st.error(f"Erro ao disparar arquivo de envio: {e}")
                else:
                    st.error("❌ Por favor, informe o seu nome completo e anexe o arquivo do comprovante.")
