importar streamlit como st
import pandas as pd
importar io
solicitações de importação
importar re
importar urllib3
import xml.etree.ElementTree as ET
importar data e hora
Importar documento do tipo docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
import firebase_admin
from firebase_admin import auth
from firebase_admin import credentials
from google.oauth2 import service_account
from datetime import datetime, timezone, timedelta

# =========================================================================
# 1. CONFIGURAÇÕES TÉCNICAS DA PÁGINA & INICIALIZAÇÃO SEGURA DO FIREBASE
# =========================================================================

st.set_page_config(
    page_title="Gerador de Fichas Catalográficas - VCB Senado",
    page_icon="logo_bibliokhan.ico",
    layout="amplo"
)

# --- ADICIONAR UM LOGOTIPO NA BARRA LATERAL ---
st.sidebar.image("logo_bibliokhan.png", use_container_width=True)

# --- BARRA LATERAL (Tudo encostado na esquerda) ---
com st.sidebar:
    st.title("**BiblioKhan**")
    st.write("**Inteligência e Automação para Bibliotecas**")
    st.write("bibliokhancontato@gmail.com")
    st.markdown("---")

se não firebase_admin._apps:
    tentar:
        firebase_secrets = dict(st.secrets["firebase"])
        firebase_secrets["private_key"] = firebase_secrets["private_key"].replace("\\n", "\n")
        
        cred = credentials.Certificate(firebase_secrets)
        firebase_admin.initialize_app(cred)
    exceto Exception como e:
        st.error(f"❌ Erro crítico nas credenciais do Firebase: {str(e)}")

# =========================================================================
# 🌟 RECARGA AUTOMÁTICA EM BACKEND
# =========================================================================
importar re

def carregar_e_filtrar_saldo(url_planilha, token_usuario):
    url_tratada = tratar_url_google_sheets(url_planilha)
    
    # header=None: tratamos a primeira linha como dado, não como título
    df = pd.read_csv(url_tratada, header=None)
    
    # Agora só temos duas colunas: 0 (token) e 1 (créditos)
    df.columns = ['token', 'creditos']
    
    # Remova espaços em branco por segurança
    df['token'] = df['token'].astype(str).str.strip()
    
    # Filtrar o usuário pelo token
    usuário = df[df['token'] == token_usuario.strip()]
    
    se não for usuário.vazio:
        # Pega o valor da coluna créditos
        valor = float(usuário['créditos'].iloc[0])
        retornar int(valor)
    outro:
        st.warning(f"Token '{token_usuario}' não encontrado!")
        retornar 0
        
def carregar_creditos_planilha(url_planilha):
    tentar:
        url_tratada = tratar_url_google_sheets(url_planilha)
        
        # Lemos o CSV garantindo que ele entenda o título
        df = pd.read_csv(url_tratada)
        
        # Debug: veja quais colunas o pandas enxergou
        # st.write("Colunas encontradas:", df.columns.tolist())
        
        # Limpeza forçada: converter a coluna de créditos para número
        # Substitua 'CREDITOS' pelo nome exato de sua coluna de saldo
        coluna_saldo = 'créditos'
        se coluna_saldo em df.columns:
            df[coluna_saldo] = pd.to_numeric(df[coluna_saldo], errors='coerce').fillna(0)
            
        retornar df
    exceto Exception como e:
        st.error(f"Erro ao processar o CSV: {e}")
        retornar Nenhum


import pandas as pd
importar streamlit como st

def atualizar_saldo_usuario(token_usuario):
    url_direta = "https://docs.google.com/spreadsheets/d/1epaFSWFhnd2Q_ZjGq32wdL3LeWpEqmFn1JFRBCh0j_U/export?format=csv&gid=0"
    
    tentar:
        # header=0 diz ao pandas: "a primeira linha é o cabeçalho (títulos)"
        df = pd.read_csv(url_direta, header=0)
        
        # Agora vamos renomear as colunas para garantir que o Python as tenha
        # (ajuste os nomes abaixo se a sua planilha tiver nomes diferentes na primeira linha)
        df.columns = ['token', 'creditos']
        
        # Remover espaços em branco dos nomes das colunas por segurança
        df.columns = df.columns.str.strip()
        
        # Filtra o token
        token_buscado = str(token_usuario).strip()
        df['token'] = df['token'].astype(str).str.strip()
        
        usuário = df[df['token'] == token_buscado]
        
        se não for usuário.vazio:
            # Pega o valor e converte para float (o erro de string sumiu porque pulamos a linha de títulos)
            saldo = int(float(usuário['créditos'].iloc[0]))
            st.session_state["créditos_ativos"] = saldo
            st.success(f"✅ Sincronizado: {saldo:.0f} créditos")
        outro:
            st.session_state["créditos_ativos"] = 0
            st.error("❌ Token não encontrado na planilha.")
            
    exceto Exception como e:
        st.error(f"Erro ao processar os dados: {e}")
        
def api_obter_produtividade_juridica(usuário):
    url_produtividade = "https://docs.google.com/spreadsheets/d/1epaFSWFhnd2Q_ZjGq32wdL3LeWpEqmFn1JFRBCh0j_U/export?format=csv&gid=54763437"
    
    # 1. Carrega a planilha
    df = pd.read_csv(url_produtividade, header=0)
    
    # 2. Renomeia APENAS as colunas que você sabe que existem,
    # mantendo o resto intacto (evita o ValueError)
    # Supondo que as 4 primeiras colunas são como você listou:
    df = df.rename(columns={
        df.columns[0]: 'dados',
        df.columns[1]: 'email',
        df.columns[2]: 'título',
        df.columns[3]: 'assunto'
    })
    
    # 3. Agora o filtro funcionará normalmente
    df['email'] = df['email'].astype(str).str.strip().str.lower()
    filtro = df[df['email'] == usuario.strip().lower()]
    
    retornar filtro
    
    retornar df
# =========================================================================
# 2. SISTEMA DE AUTENTICAÇÃO E CONTROLE DE SESSÃO COMERCIAL
# =========================================================================

def verificar_login_firebase(e-mail, senha):
    tentar:
        usuário = auth.get_user_by_email(email)
        st.session_state["logado"] = True
        st.session_state["usuario_atual"] = usuário.email
        atualizar_saldo_usuario(usuario.email)
        retornar Verdadeiro
    exceto Exception como e:
        st.error("❌ Acesso negado: E-mail não cadastrado ou credenciais inválidas.")
        retornar Falso

se "logado" não estiver em st.session_state:
    st.session_state["logado"] = Falso

se "creditos_ativos" não estiver em st.session_state:
    st.session_state["créditos_ativos"] = 0

# Exibe o saldo na barra lateral caso o usuário esteja logado
se st.session_state["logado"]:
    com st.sidebar:
        if st.session_state["créditos_ativos"] > 0:
            st.success(f"💳 Saldo: {st.session_state['créditos_ativos']} fichas")
        outro:
            st.error("💳 Sem créditos ativos")

# =========================================================================
# 3. INTERFACE DE LOGIN OU FLUXO DO APLICATIVO PROTEGIDO
# =========================================================================

se não st.session_state["logado"]:
    st.markdown("# 🔒 Área do Cliente")
    st.markdown("### Faça o login para acessar o Gerador de Fichas Catalográficas.")
    
    com st.form("login_form"):
        email_input = st.text_input("E-mail do Usuário").strip()
        senha_input = st.text_input("Senha de acesso", type="password").strip()
        botao_entrar = st.form_submit_button("Entrar no Sistema")
        
        se botao_entrar:
            se email_input e senha_input:
                verificar_login_firebase(email_input, senha_input)
                se st.session_state["logado"]:
                    st.rerun()
            outro:
                st.warning("⚠️ Por favor, preencha o e-mail e a senha.")

    st.markdown("---")
    with st.expander("🔑 Esqueceu sua senha ou quer trocar a senha provisória?"):
        st.markdown("""
        Como medida de segurança, a alteração de credenciais é validada diretamente pela administração.
        
        Para redefinir sua senha, entre em contato diretamente com o suporte técnico através do e-mail informado na lateral do sistema ou pelo canal de atendimento onde adquiriu o produto. Um link oficial de redefinição será enviado para o seu e-mail cadastrado.
        "")

# --- TELA DE CADASTRO (Abaixo do Login) ---
se não st.session_state["logado"]:
    with st.expander("📝 Ainda não tem conta? Clique aqui para se cadastrar"):
        com st.form("cadastro_form"):
            novo_email = st.text_input("Novo E-mail").strip()
            nova_senha = st.text_input("Escolha uma senha", type="password")
            botao_cadastrar = st.form_submit_button("Criar Conta")
            
            se botao_cadastrar:
                se novo_email e nova_senha:
                    tentar:
                        # Chama o Firebase para criar o usuário
                        auth.create_user(email=novo_email, senha=nova_senha)
                        st.success("✅ Conta criada com sucesso! Faça o login agora.")
                    exceto Exception como e:
                        st.error(f"❌ Erro ao criar conta: {e}")
                outro:
                    st.warning("⚠️ Preencha e-mail e senha.")

outro:
    # --- CONTEÚDO DO APLICATIVO COMERCIAL ---
    st.markdown("""
        <style>
        textarea {
            font-family : 'Courier New', Courier, monospace !important;
        }
        .stTabs [data-baseweb="tab-list"] { gap: 24px; }
        .stTabs [data-baseweb="tab"] {
            altura: 50px;
            espaço em branco: pré-quebra;
            cor de fundo: #f0f2f6;
            raio da borda: 5px 5px 0px 0px;
            espaço: 1px;
            preenchimento superior: 10px;
            preenchimento-inferior: 10px;
        }
        .stTabs [aria-selected="true"] { background-color: #B19FFB !important; color: black !important; font-weight: bold; }
        </style>
        "", unsafe_allow_html=True)

    se "lote_fichas" não estiver em st.session_state:
        st.session_state.lote_fichas = []

    se "assuntos_selecionados" não estiver em st.session_state:
        st.session_state.assuntos_selecionados = []

    def buscar_vcb_senado(termo_busca):
        url_api = "https://adm.senado.leg.br/vcb/vocab/services.php"
        params = {"task": "search", "arg": termo_busca, "output": "json"}
        tentar:
            resposta = requests.get(url_api, params=params, timeout=8, verify=False)
            se resposta.status_code == 200:
                dados = resposta.json()
                resultados_formatados = []
                bloco_resultado = dados.get("resultado", {})
                se isinstance(bloco_result, dict):
                    para chave, item em bloco_result.items():
                        Se isinstance(item, dict) e "string" em item:
                            resultados_formatados.append({
                                "termo": item["string"].strip(),
                                "id": f"VCB-{item.get('term_id', chave)}",
                                "note": "Termo oficial homologado pelo Vocabulário Controlado do Senado Federal."
                            })
                retornar resultados_formatados
        exceto Exceção:
            retornar []
        retornar []

    def gerar_docx_lote(lista_fichas):
        doc = Documento()
        
        # Configuração das margens
        seção = doc.seções[0]
        seção.margem_esquerda = Pt(72)
        seção.margem_direita = Pt(72)

        para idx, ficha_texto em enumerate(lista_fichas):
            se idx > 0:
                doc.add_page_break()
            
            # Adicionado a tabela com o estilo 'Table Grid' que força as bordas
            tabela = doc.add_table(linhas=1, colunas=1)
            table.style = 'Tabela em Grade'
            tabela.ajusteautomático = Falso
            tabela.permitir_ajuste_automático = Falso
            tabela.colunas[0].largura = Pt(400)
            
            # Acessa ase
            célula = tabela.célula(0, 0)
            
            # Remover parâmetros padrão para garantir o controle total
            célula._elemento.limpar_conteúdo()
            
            # Adicionado o texto configurando a fonte
            p = célula.adicionar_parágrafo()
            executar = p.adicionar_execução(ficha_texto)
            run.font.name = 'Courier New'
            tamanho.da.fonte.da.execução = Pt(10)
            
            # Ajusta o alinhamento e recuo dentro da caixa
            p.paragraph_format.left_indent = Pt(10)
            p.paragraph_format.right_indent = Pt(10)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(10)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        buffer de retorno
    def formatar_entrada_e_corpo(tipo_autor, autores_lista, entidade, titulo, tem_organizador, organizador_nome, tipo_org, tem_tradutor, tradutor_nome):
        entrada = ""
        corpo_autores = ""
        entrada_por_título = Falso
        
        if tem_organizador e tipo_autor == "Pessoa Física" e não any(a.strip() para a in autores_lista):
            entrada_por_título = Verdadeiro
            entrada = ""
            corpo_autores = f"{tipo_org} por {organizador_nome.strip()}"
        elif tipo_autor == "Entidade (Órgão/Instituição)":
            entrada = entidade.strip().upper()
            corpo_autores = ""
        outro:
            autores = [a.strip() para a em autores_lista se a.strip()]
            qtd = len(autores)
            
            se qtd == 1:
                partes = autores[0].split()
                entrada = f"{partes[-1].upper()}, {' '.join(partes[:-1])}." if len(partes) > 1 else f"{autores[0].upper()}."
                corpo_autores = autores[0]
            senão se qtd >= 2 e qtd <= 3:
                partes = autores[0].split()
                entrada = f"{partes[-1].upper()}, {' '.join(partes[:-1])}." if len(partes) > 1 else f"{autores[0].upper()}."
                corpo_autores = ", ".join(autores)
            senão se qtd >= 4:
                entrada_por_título = Verdadeiro
                entrada = ""  
                corpo_autores = f"{autores[0]} [et al.]"
                
            se tem_organizador e organizador_nome.strip() e qtd < 4:
                corpo_autores += f" ; {tipo_org} por {organizador_nome.strip()}"

        if tem_tradutor e tradutor_nome.strip():
            se corpo_autores:
                corpo_autores += f" ; tradução por {tradutor_nome.strip()}"
            outro:
                corpo_autores = f"tradução por {tradutor_nome.strip()}"
                
        return entrada, corpo_autores, entrada_por_titulo

    def buscar_na_tabela_cutter(texto_para_busca, titulo_obra):
        se não for texto_para_busca ou não titulo_obra: retorne "X000x"
        url_csv = "https://raw.githubusercontent.com/Biblio-Khan/gerador-ficha-cat/refs/heads/main/cutter.csv"
        tentar:
            df = pd.read_csv(url_csv, sep=',', encoding='utf-8', quotechar='"')
        exceto Exceção:
            return f"{texto_para_busca.strip().upper()[0]}200{titulo_obra.strip().lower()[0]}"
        
        df.columns = df.columns.str.strip().str.lower()
        col_nome = 'nome' se 'nome' estiver em df.columns senão df.columns[0]
        col_id = 'id' se 'id' estiver em df.columns senão df.columns[1]
        
        df['Name_Clean'] = df[col_nome].astype(str).str.strip().str.upper()
        sub_busca = texto_para_busca.strip().upper()
        
        correspondência = df[df['Name_Clean'] <= sub_busca].sort_values(by='Name_Clean').tail(1)
        num = "200"
        se não corresponder.vazio:
            num = str(match[col_id].values[0]).strip().split('.')[0]
            
        titulo_limpo = titulo_obra.strip().upper()
        artigos = ["O ", "A", "OS", "AS", "UM", "UMA", "UNS", "UMAS "]
        para artigo em grupo:
            if titulo_limpo.startswith(artigo):
                titulo_limpo = titulo_limpo[len(artigo):].strip()
                quebrar
        letra_titulo = titulo_limpo[0].lower() if titulo_limpo else "t"
        retornar f"{sub_busca[0]}{num}{letra_titulo}"

    def calcular_cutter(tipo_autor, autores_lista, entidade="", titulo="", tem_organizador=False, organizador_nome=""):
        if tipo_autor == "Entidade (Órgão/Instituição)" e entidade:
            texto_base = entidade
        elif tipo_autor == "Pessoa Física" and autores_lista and any(a.strip() for a in autores_lista):
            autor_principal = [a.strip() para a em autores_lista se a.strip()][0]
            partes = autor_principal.split()
            texto_base = partes[-1] if len(partes) > 1 else autor_principal
        elif tem_organizador ou tipo_autor == "Organizador":
            partes_org = organizador_nome.strip().split()
            texto_base = partes_org[-1] if len(partes_org) > 1 else organizador_nome
        outro:
            texto_base = "Autor"
        return buscar_na_tabela_cutter(texto_base, titulo)

    def gerar_marc21_completo(dados):
        linhas_marc = [
            "000 00000nam a2200000 i 4500",
            f"100 1#$a{dados.get('entrada', '')}",
            f"245 10$a{dados.get('título', '')}"
        ]
    
        # Tag 260: Só adicione se houver pelo menos a cidade ou a editora
        local = dados.get('local_editora', '')
        if local e local != " : ": # Verifica se não está vazio
            marc_lines.append(f"260 ##$a{local}")
    
        # Tag 300: Só adicione se houver páginas OU dimensões
        paginas = dados.get('páginas', '')
        dimensões = dados.get('dimensões', '')
        se forem páginas ou dimensões:
            marc_lines.append(f"300 ##$a{paginas} p. ; {dimensões} cm.")
    
        # Etiqueta 502
        tipo = dados.get('tipo', '')
        se tipo e "Livro" não estiverem no tipo:
            inst = dados.get('instituicao', '')
            ano = dados.get('ano', '')
            marc_lines.append(f"502 ##$a{tipo} - {inst}, {ano}.")
    
        # Etiquetas 650
        para assunto em dados.get('assuntos', []):
            se for relevante:
                marc_lines.append(f"650 #4$a{assunto}")
    
        se dados.get('area'):
            marc_lines.append(f"650 #4$a{dados.get('area')}")

        retornar "\n".join(marc_lines)

   

    # =========================================================================
    # SISTEMA DE ABAS (CATALOGAÇÃO & CRÉDITOS LIMITADOS ATÉ 300)
    # =========================================================================
    tab_gerador, tab_financeiro, tab_produtividade = st.tabs([
    "⚖️ Gerar Ficha",
    "💳 Compra e Gestão de Créditos",
    "📊 Painel de Produtividade"
])

    com tab_gerador:
        if st.session_state["créditos_ativos"] <= 0:
            st.warning("🔒 O painel de salvamento está bloqueado. Adquira créditos ou aguarde a restauração para continuar.")

        st.title("Gerador de Fichas Catalográficas — NBR/AACR2")
        st.caption("Mesa técnica integrada via Web Service ao Vocabulário Controlado Básico (VCB) do Senado Federal.")

        st.markdown("---")
        container_lote = st.container()
        com container_lote:
            # Adicionei uma terceira coluna (col_lote_3) para o botão MARC
            col_lote_1, col_lote_2, col_lote_3 = st.columns([2, 1, 1])
            qtd_fichas = len(st.session_state.lote_fichas)
            col_lote_1.subheader(f"📦 Lote: {qtd_fichas} Ficha(s)")
            
            se qtd_fichas > 0:
                # 1. . Palavra (Mantido)
                arquivo_word = gerar_docx_lote([f["texto_ficha"] for f in st.session_state.lote_fichas])
                col_lote_2.download_button(
                    label="📥 Palavra",
                    dados=arquivo_word,
                    file_name="lote_fichas.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # 2. MARC 21 (Novo)
                # Como a ficha atual é apenas texto, falei o texto para a função
                # (A função tratará de criar um registro básico a partir do texto)
                conteudo_marc = "\n\n".join([gerar_marc21_completo(f["dados_marc"]) for f em st.session_state.lote_fichas])

                # Primeiro botão (MARC 21 original)
                col_lote_3.download_button(
                label="📥 MARC 21 (.mrc)",
                dados=conteudo_marc,
                file_name="lote_juridico.mrc",
                mime="texto/simples"
                )

                # Segundo botão (TXT para copiar/colar)
                col_lote_3.download_button(
                label="📋 MARC 21 (.txt)",
                dados=conteudo_marc,
                file_name="lote_juridico.txt",
                mime="texto/simples"
                )
                
                # 3. Limpar (Mantido)
                if col_lote_2.button("🗑️ Limpar"):
                    st.session_state.lote_fichas = []
                    st.rerun()
            outro:
                col_lote_2.info("O lote está vazio.")
        
        st.markdown("---")
        col_esquerda, col_direita = st.columns(2)

        com col_esquerda:
            st.subheader("1. Metadados e Responsabilidade")
            classificação = st.text_input("Número de Classificação (CDD ou CDU)", valor="340.1")
            tipo_autor = st.radio("Tipo de Autoria Principal", ["Pessoa Física", "Entidade (Órgão/Instituição)"], horizontal=True)
            
            autores_lista = []
            entidade_nome = ""
            
            if tipo_autor == "Pessoa Física":
                qtd_autores_input = st.number_input("Quantidade de autores principais (0 se houver apenas Organizador)", min_value=0, max_value=10, value=1)
                para i em range(int(qtd_autores_input)):
                    autores_lista.append(st.text_input(f"Autor {i+1} (Nome Sobrenome)", key=f"autor_{i}"))
            outro:
                entidade_nome = st.text_input("Nome da Entidade (Ex: Brasil. Supremo Tribunal Federal)")
                
            título = st.text_input("Título Principal")
            st.markdown("---")
            col_resp_1, col_resp_2 = st.columns(2)
            
            com col_resp_1:
                tem_organizador = st.checkbox("Possui Organizador/Coordenador?")
                organizador_nome = ""
                tipo_org, abreviatura_org = "", ""
                se tem_organizador:
                    papel = st.selectbox("Função:", ["Organizador", "Coordenador", "Compilador"])
                    organizador_nome = st.text_input("Nome do Responsável")
                    if papel == "Organizador": tipo_org, abreviatura_org = "organizado", "org."
                    elif papel == "Coordenador": tipo_org, abreviatura_org = "coordenado", "coord."
                    else: tipo_org, abreviatura_org = "compilado", "comp."
                    
            com col_resp_2:
                tem_tradutor = st.checkbox("A obra possui tradutor?")
                tradutor_nome = ""
                se tem_tra menina:
                    tradutor_nome = st.text_input("Nome do Tradutor (Nome Sobrenome)", key="trad_nome")

            st.markdown("---")
            st.subheader("2. Publicação e Descrição Física")
            edição = st.text_input("Edição e Volume (Ex: 2. ed., 3. ed. rev. e ampl.)", valor="1. ed.")
            editora = st.text_input("Editora")
            cidade = st.text_input("Cidade de Publicação", value="Brasília")
            ano = st.text_input("Ano de Publicação", value="2026")
            paginas_input = st.text_input("Número de páginas/folhas", value="180")
            dimensoes_input = st.text_input("Dimensões", valor="30 cm")
            
            tem_colecao = st.checkbox("Esta obra faz parte de uma Coleção / Série?")
            colecao_nome = ""
            se tem_colecao:
                colecao_nome = st.text_input("Nome da Coleção e Volume (Ex: Biblioteca jurídica, v. 12)")

            # === BLOCO CORRIGIDO: Tipo de Trabalho Acadêmico ===
            st.markdown("---")
            st.subheader("3. Tipo de Documento / Trabalho Acadêmico")
            
            grau_acadêmico = st.selectbox(
                "Tipo de Obra:",
                ["Livro / Código / Obra Geral", "Tese (Doutorado)", "Dissertação (Mestrado)", "Monografia (Especialização)", "Monografia (Graduação)"]
            )

            # Inicializa as variáveis ​​vazias por padrão
            instituição = ""
            área_concentração = ""

            # Se for selecionado qualquer trabalho acadêmico, mostre os campos adicionais
            if grau_academico != "Livro / Código / Obra Geral":
                instituicao = st.text_input("Instituição / Universidade (Ex: Faculdade de Direito da USP):")
                area_concentracao = st.text_input("Área de Concentração / Curso (Ex: Direito Civil):")
                
            isbn = st.text_input("ISBN (Ex: 978-65-0000-00-0)")
            suporte = st.radio("Suporte da Obra", ["Impresso", "Digital"], horizontal=True)
            url_acesso = st.text_input("URL de Acesso / DOI") if suporte == "Digital" else ""

            com col_sentido:
                st.subheader("3. Indexação por Assunto")
    
                # Campo direto de busca para o VCB do Senado
                termo_busca = st.text_input("Digite um termo para pesquisar no Vocabulário Controlado do Senado:")
    
                se termo_busca:
                    resultados_vcb = buscar_vcb_senado(termo_busca)
        
                    se resultados_vcb:
                        st.success(f"{len(resultados_vcb)} conceitos localizados no Senado!")
                        mapeamento_opcoes = {item["termo"]: item por item nos resultados_vcb}
                        lista_opcoes = sorted(list(mapeamento_opcoes.keys()))
                        termo_selecionado = st.selectbox("Seleção do conceito oficial (Senado):", lista_opcoes)
            
                        if st.button("➕ Vincular Assunto do Senado"):
                            se termo_selecionado não estiver em st.session_state.assuntos_selecionados:
                                st.session_state.assuntos_selecionados.append(termo_selecionado)
                                st.rerun()
                        
                            
                            

                st.markdown("---")

                st.markdown("##### Adicionar Assunto Manualmente")
                assunto_manual = st.text_input("Digite um assunto personalizado:")
                if st.button("➕ Vincular Assunto Manual"):
                    se o assunto_manual.strip():
                        termo_limpo = assunto_manual.strip()
                        se termo_limpo não estiver em st.session_state.assuntos_selecionados:
                            st.session_state.assuntos_selecionados.append(termo_limpo)
                            st.rerun()

                if st.session_state.get("assuntos_selecionados"):
                    st.write("**Assuntos Vinculados à Ficha:**")

                    #Exclusão individual do assunto
                    para idx, ass in enumerate(st.session_state.assuntos_selecionados):
                        col_assunto, col_excluir = st.columns([9, 1])
                        com col_assunto:
                            st.write(f"{idx + 1}. {ass}")
                        com col_excluir:
                            se st.botão(
                                "❌",
                                chave=f"remover_assunto_{idx}",
                                help="Remover apenas este assunto",
                            ):
                                st.session_state.assuntos_selecionados.pop(idx)
                                st.rerun()

                    # Botão para limpar toda a lista
                    if st.button("🗑️ Limpar Todos os Assuntos"):
                        st.session_state.assuntos_selecionados = []
                        st.rerun()

                st.markdown("---")
                st.subheader("4. Fechamento e Visualização da Ficha")

                # Formatação dos dados da ficha
                (
                    entrada_principal,
                    responsabilidade,
                    entrada_por_título,
                ) = formatar_entrada_e_corpo(
                    tipo_autor=tipo_autor,
                    autores_lista=autores_lista,
                    entidade=entidade_nome,
                    título=título,
                    tem_organizador=tem_organizador,
                    organizador_nome=organizador_nome,
                    tipo_org=tipo_org,
                    tem_tracord=tem_tracord,
                    tradutor_nome=tradução_nome,
                )

                cortador = cortador_de_casa(
                    tipo_autor,
                    autores_lista,
                    entidade=entidade_nome,
                    título=título,
                    tem_organizador=tem_organizador,
                    organizador_nome=organizador_nome,
                )
                dgm = " [recurso eletrônico]" if suporte == "Digital" else ""
                desc_física = (
                    f"1 recurso online ({paginas_input} p.)"
                    se suporte == "Digital"
                    senão f"{paginas_input} p."
                )
                if suporte != "Digital" and dimensionoes_input.strip():
                    desc_fisica = f"{desc_fisica} ; {dimensoes_input.strip()}"

                bloco_colecao = ""
                if tem_colecao e colecao_nome.strip():
                    text_colecao = colecao_nome.strip()
                    text_colecao = text_colecao[0].upper() + text_colecao[1:]
                    bloco_colecao = f" ({text_colecao})"

                # Nota de trabalho acadêmico (ABNT)
                nota_boa_str = ""
                if grau_academico != "Livro / Código / Obra Geral":
                    inst_str = (
                        f" – {instituicao.strip()}" if instituicao.strip() else ""
                    )
                    área_str = (
                        f" em {area_concentracao.strip()}"
                        se area_concentracao.strip()
                        outro ""
                    )
                    nota_trabalho_str = f"\n {grau_academico}{area_str}{inst_str}, {ano.strip()}."

                nota_acesso = (
                    f"\n Modo de acesso: {url_acesso}"
                    if suporte == "Digital" e url_acesso
                    outro ""
                )
                isbn_bloco = f"\n ISBN {isbn}" if isbn.strip() else ""
                nota_traducao = (
                    "\nTraduzido da obra original."
                    if tem_tradutor e tradutor_nome.strip()
                    outro ""
                )
                ed_bloco = f"{edicao.strip()} – " if edicao.strip() else ""
                pub_bloco = f"{cidade.strip()} : {editora.strip()}, {ano.strip()}."

                string_assuntos = " ".join(
                    [
                        f"{i + 1}. {ass}"
                        para eu, bunda em enumerar(st.session_state.assuntos_selecionados)
                    ]
                )
                rastreabilidade = ""
                romanos = ["I", "II", "III", "IV", "V"]
                r_idx = 0

                se não entrada_por_titulo:
                    rastreabilidade += f" {romanos[r_idx]}. Título."
                    r_idx += 1

                if tem_organizador e organizador_nome.strip():
                    partes_org = organizador_nome.strip().split()
                    nome_inba_org = (
                        f"{partes_org[-1].upper()}, {' '.join(partes_org[:-1])}"
                        se len(partes_org) > 1
                        senão organizador_nome.strip().upper()
                    )
                    rastreabilidade += f" {romanos[r_idx]}. {nome_invertido_org}, {abreviatura_org}."
                    r_idx += 1

                if tem_tradutor e tradutor_nome.strip():
                    partes_trad = tradutor_nome.strip().split()
                    nome_inla_trad = (
                        f"{partes_trad[-1].upper()}, {' '.join(partes_trad[:-1])}"
                        se len(partes_trad) > 1
                        senão tradutor_nome.strip().upper()
                    )
                    rastreabilidade += f" {romanos[r_idx]}. {nome_invertido_trad}, trad."
                    r_idx += 1

                # Montagem da string final da Ficha
                se entrada_por_título:
                    txt_ficha = f"""{classificação}
            {cutter} {titulo.strip()}{dgm} / {responsabilidade}. – {ed_bloco}{pub_bloco}
                        {desc_fisica}.{bloco_colecao}{nota_trabalho_str}{nota_traducao}{nota_acesso}{isbn_bloco}
            
                        {string_assuntos}{rastreabilidade}"""
                outro:
                    txt_ficha = f"""{classificação}
            {cortador} {entrada_principal}
                        {titulo.strip()}{dgm} / {responsabilidade}. – {ed_bloco}{pub_bloco}
                        {desc_fisica}.{bloco_colecao}{nota_trabalho_str}{nota_traducao}{nota_acesso}{isbn_bloco}
            
                        {string_assuntos}{rastreabilidade}"""

                st.text_area(
                    "Visualização Normativa (Fonte Monoespaçada)",
                    valor=txt_ficha,
                    altura=240,
                )

                # Validação e salvamento
                btn_salvar = st.button(
                    "💾CONCLUIR FICHA E ENVIAR AO LOTE",
                    disabled=st.session_state.get("creditos_ativos", 0) <= 0,
                )

                se btn_salvar:
                    válido = Verdadeiro
                    se (
                        tipo_autor == "Pessoa Física"
                        e não qualquer(a.strip() para a em autores_lista)
                        e não tem_organizador
                    ):
                        válido = Falso
                        st.erro(
                            "❌É necessário informar ao menos um autor ou organizador."
                        )

                    se não titulo.strip():
                        válido = Falso
                        st.error("❌ O título é obrigatório.")

                    se válido:
                        with st.spinner("Gravando ficha e atualizando saldo na nuvem..."):
                            tentar:
                                importar json

                                solicitações de importação

                                url_script = st.secrets["URL_SCRIPT_GOOGLE"]
                                lista_assuntos = st.session_state.get(
                                    "assuntos_selecionados", []
                                )
                                sência_texto = (
                                    ", ".join(lista_assuntos)
                                    se lista_assuntos
                                    caso contrário "Não informado"
                                )

                                carga útil = {
                                    "e-mail": st.session_state["usuario_atual"],
                                    "acao": "descontar",
                                    "titulo": titulo if titulo else "Não Informado",
                                    "assunto":jismo_texto,
                                }

                                resposta_google = solicitações.post(
                                url_script, json=payload, timeout=15
                                )

                                se resposta_google.status_code == 200:
                                    contado = (
                                        resposta_google.content.decode("utf-8-sig").strip()
                                    )
                                    se "{" no contexto:
                                        conteudo = conteudo[conteudo.find("{") :]

                                    resultado_json = json.loads(conteudo)

                                    if resultado_json.get("status") == "sucesso":
                                        cabeça_completa = {
                                            "texto_ficha": txt_ficha,
                                            "dados_marc": {
                                                "entrada": entrada_principal,
                                                "título": título,
                                                "local_editora": f"{cidade.strip()} : {editora.strip()}",
                                                "tipo": grau_academico,
                                                "instituicao": instituicao.strip(),
                                                "área": ​​área_concentração.strip(),
                                                "assuntos": st.session_state.assuntos_selecionados,
                                                "ano": ano.strip(),
                                                "páginas": paginas_input,
                                                "dimensões": dimensões_entrada,
                                            },
                                        }
                                        st.session_state.lote_fichas.append(
                                            cabeça_completa
                                        )
                                        st.session_state["créditos_ativos"] -= 1
                                        st.session_state.assuntos_selecionados = []
                                        st.success("✅ Ficha guardada com sucesso!")
                                        st.rerun()
                                    outro:
                                        st.erro(
                                            f"❌ Erro na planilha: {resultado_json.get('mensagem', 'Erro desconhecido')}"
                                        )
                                outro:
                                    st.erro(
                                        f"❌ Falha de conexão. Status: {resposta_google.status_code}"
                                    )

                            exceto Exception como e:
                                st.error(f"❌ Erro ao processar requisição: {e}")

               
                
# Abaixo, fora de qualquer bloco 'if' ou 'try', começa o tab_financeiro
    com tab_financeiro:
        st.header("💳 Gestão Financeira e Saldo")
        # ... resto do seu código da aba
        col_f1, col_f2 = st.columns(2)
    
    com col_f1:
        st.subheader("🔄 Sincronização")
        st.info(f"Seu sistema está vinculado ao e-mail: **{st.session_state['usuario_atual']}**")
        if st.button("Atualizar meu Saldo"):
            with st.spinner("Puxando dados atualizados do Sheets..."):
                atualizar_saldo_usuario(st.session_state["usuario_atual"])
                st.success("Saldo verificado com sucesso!")
                st.rerun()

        com col_f2:
            st.subheader("🛒 Tabela de Preços")
            st.markdown("""
            * **20 Fichas** — R$ 55,00
            * **30 Fichas** — R$ 80,00
            * **100 Fichas** — R$ 240,00
            * **300 Fichas** — R$ 660,00
            * **600 Fichas** — R$ 1.200,00
            "")
            st.info("🔑 **FOTO:** `bibliokhancontato@gmail.com`")

        st.markdown("---")
        st.subheader("📩 Envio de Comprovante")
        
        com st.form("pix_form_original"):
            email_cliente = st.text_input("E-mail de Cadastro no Sistema", value=st.session_state["usuario_atual"], desabilitado=True)
           
            pacote_escolhido = st.selectbox(
                "Qual pacote de créditos você comprou?",
                opções=[
                    "20 Fichas (R$ 55,00)",
                    "30 Fichas (R$ 80,00)",
                    "100 Fichas (R$ 240,00)",
                    "300 Fichas (R$ 660,00)",
                    "600 Fichas (R$ 1.200,00)"
                ]
            )
            
            comprovante = st.file_uploader("Anexe a imagem ou PDF do comprovante do PIX", type=["jpg", "png", "jpeg", "pdf"])
            
            if st.form_submit_button("Enviar para Restauração de Saldo"):
                se comprovante não for None:
                    with st.spinner("Enviando comprovante para o suporte... Por favor, aguarde."):
                        tentar:
                            tg_token = st.secrets["TELEGRAM_BOT_TOKEN"]
                            tg_chat = st.secrets["TELEGRAM_CHAT_ID"]
    
                            fuso_brasilia = fuso horário(timedelta(horas=-3))
                            data_hora_br = datetime.now(fuso_brasilia).strftime('%d/%m/%Y %H:%M:%S')
                            
                            texto_notificação = (
                                f"🔥 *NOVO COMPROVANTE RECEBIDO!*\n\n"
                                f"📧 *E-mail do Cliente:* {st.session_state['usuario_atual']}\n"
                                f"💰 *Pacote Escolhido:* {pacote_escolhido}\n"
                                f"📅 *Dados/Hora:* {data_hora_br}"
                            )
                            
                            url_api_telegram = f"https://api.telegram.org/bot{tg_token}/sendPhoto"
                            arquivo_envio = {"photo": (comprovante.name, comprovante.getvalue(), comprovante.type)}
                            dados_requisicao = {"chat_id": tg_chat, "caption": texto_notificacao, "parse_mode": "Markdown"}
                            
                            resposta_tg = requests.post(url_api_telegram, data=dados_requisicao, files=ficheiro_envio, timeout=15)
                            
                            se resposta_tg.status_code == 200:
                                st.success("✅ Comprovante enviado com sucesso!")
                                st.info("⏳ O seu saldo será atualizado assim que a validação for concluída.")
                            outro:
                                st.error(f"Erro na API de comunicação (Código {resposta_tg.status_code}).")
                        exceto Exception como e:
                            st.error(f"Erro ao disparar arquivo de envio: {e}")
                outro:
                    st.error("❌ Por favor, informe seu nome completo e anexo o arquivo do comprovante.")

# ---------------------------------------------------------------------
#NOVA ABA: PAINEL DE PRODUTIVIDADE JURÍDICA (COM TRAVA DE LOGIN)
# ---------------------------------------------------------------------
# Só execute este bloco se o usuário já tiver passado pela tela de login
se st.session_state.get("usuario_atual"):

    # Verifica dinamicamente se a aba foi criada no topo do arquivo
    Se 'tab_produtividade' não estiver em locals() e 'tab_produtividade' não estiver em globals():
        st.markdown("---")
        tab_produtividade = st.container()

    com a aba_produtividade:
        st.title("Painel de Produtividade")
        st.subheader(f"Análise de Obras Processadas por {st.session_state.get('usuario_atual', 'Usuário')}")

        with st.spinner("Carregando dados de produtividade..."):
            dados = api_obter_produtividade_juridica(st.session_state.get("usuario_atual", ""))

        # Verifique se o objeto 'dados' é um DataFrame válido e não está vazio
        import pandas as pd
        
        Se isinstance(dados, pd.DataFrame) e não dados.empty:
            st.write(f"Total de registros encontrados: {len(dados)}")
            st.dataframe(dados) # Aqui os dados aparecem
        outro:
            st.info("Você ainda não possui registros de fichas geradas.")

            # 1. Converta os dados obtidos da API para um DataFrame do Pandas
            df = pd.DataFrame(dados)

            # 2. Coleta todos os assuntos, quebra pelas vírgulas e limpa os espaços
            todos_assuntos = []
            para linha_assunto em df['assunto']:
                se linha_assunto:
                    if str(linha_assunto) != "Não informado":
                        partes = [a.strip().title() for a in str(linha_assunto).split(",") if a.strip()]
                        todos_assuntos.extend(partes)

            # 3. Contar a frequência de cada assunto individual
            se todos_assuntos:
                df_contagem = pd.DataFrame(todos_assuntos, columns=["Área/Assunto"]).value_counts().reset_index(name="Quantidade")
            outro:
                df_contagem = pd.DataFrame()

            # 4. Mostra os cartões de resumo (Métricas)
            col_card1, col_card2 = st.columns(2)
            com col_card1:
                st.metric("Total de Processos/Livros", len(df))
            com col_card2:
                st.metric("Total de Assuntos Mapeados", len(df_contagem))

            st.markdown("---")
            
            # 5. Renderiza o Gráfico de Barras se houver assuntos mapeados
            se não df_contagem.empty:
                st.write("### Temas mais exigidos nas suas fichas")
                st.bar_chart(
                    dados=df_contagem,
                    x="Área/Assunto",
                    y="Quantidade",
                    cor="#0077B6",
                    use_container_width=True
                )
                st.markdown("---")

            # 6. Histórico de Obras Processadas e Opção de Download
            st.write("### Histórico de Fichas Emitidas")
            
            df_exibicao = df.copy()
            
            df_exibicao = df_exibicao.rename(colunas={
                "dados": "Dados/Hora",
                "título": "Título da Obra",
                "assunto": "Assuntos Indexados"
            })
            
            if "Dados/Hora" em df_exibicao.columns:
                tentar:
                    df_exibicao["Dados/Hora"] = pd.to_datetime(df_exibicao["Dados/Hora"]).dt.strftime('%d/%m/%Y %H:%M')
                exceto:
                    passar

            colunas_relatorio = ["Dados/Hora", "Título da Obra", "Assuntos Indexados"]
            df_final = df_exibicao[colunas_relatorio]

            # === BOTÃO DE DOWNLOAD ===
            csv_dados = df_final.to_csv(index=False, sep=";").encode('utf-8-sig')
            
            st.download_button(
                label="📥 Baixar Relatório em CSV (Excel)",
                dados=csv_dados,
                file_name=f"produtividade_juridica_{st.session_state['usuario_atual'].split('@')[0]}.csv",
                mime="texto/csv",
                use_container_width=True
            )
            
            st.write("")
            
            st.dataframe(
                df_final,
                use_container_width=True,
                ocultar_índice=Verdadeiro
            )
