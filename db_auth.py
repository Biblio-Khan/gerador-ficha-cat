import streamlit as st
import libsql_experimental as libsql
import bcrypt

def conectar_turso():
    """Conecta ao banco de dados Turso usando os Secrets do Streamlit."""
    url = st.secrets["turso"]["TURSO_DATABASE_URL"]
    token = st.secrets["turso"]["TURSO_AUTH_TOKEN"]
    return libsql.connect(database=url, auth_token=token)

def cadastrar_usuario(email, nome, senha):
    """Cadastra um novo usuário no Turso. Ganha 4 créditos automaticamente."""
    conn = conectar_turso()
    cursor = conn.cursor()
    
    # Criptografa a senha com bcrypt
    senha_bytes = senha.encode('utf-8')
    salt = bcrypt.gensalt()
    senha_hash = bcrypt.hashpw(senha_bytes, salt).decode('utf-8')
    
    try:
        cursor.execute(
            "INSERT INTO usuarios (email, nome, senha_hash) VALUES (?, ?, ?)",
            (email.lower().strip(), nome.strip(), senha_hash)
        )
        conn.commit()
        return True, "Usuário cadastrado com sucesso! Você ganhou 4 créditos de boas-vindas."
    except Exception as e:
        if "UNIQUE" in str(e).upper():
            return False, "Este e-mail já está cadastrado."
        return False, f"Erro ao cadastrar: {e}"

def autenticar_usuario(email, senha):
    """Valida o login do usuário e retorna seus dados (incluindo créditos)."""
    conn = conectar_turso()
    cursor = conn.cursor()
    
    res = cursor.execute(
        "SELECT id, nome, email, senha_hash, creditos, is_admin FROM usuarios WHERE email = ?",
        (email.lower().strip(),)
    )
    usuario = res.fetchone()
    
    if not usuario:
        return False, "E-mail ou senha incorretos.", None
        
    user_id, nome, user_email, senha_hash_db, creditos, is_admin = usuario
    
    # Compara a senha informada com o hash salvo no banco
    if bcrypt.checkpw(senha.encode('utf-8'), senha_hash_db.encode('utf-8')):
        dados_usuario = {
            "id": user_id, 
            "nome": nome, 
            "email": user_email,
            "creditos": creditos,
            "is_admin": is_admin
        }
        return True, "Login realizado com sucesso!", dados_usuario
    else:
        return False, "E-mail ou senha incorretos.", None

def descontar_credito_e_registrar(usuario_id, autor, titulo, assunto):
    """
    Desconta 1 crédito do usuário e salva a ficha na tabela de produtividade.
    """
    conn = conectar_turso()
    cursor = conn.cursor()
    
    # 1. Verifica se o usuário ainda tem créditos
    res = cursor.execute("SELECT creditos FROM usuarios WHERE id = ?", (usuario_id,))
    row = res.fetchone()
    
    if not row or row[0] <= 0:
        return False, "Você não possui créditos suficientes para gerar a ficha."
        
    try:
        # 2. Desconta 1 crédito
        cursor.execute("UPDATE usuarios SET creditos = creditos - 1 WHERE id = ?", (usuario_id,))
        
        # 3. Registra os dados da ficha gerada para a coleta de produtividade
        cursor.execute(
            "INSERT INTO registros_produtividade (usuario_id, autor, titulo, assunto) VALUES (?, ?, ?, ?)",
            (usuario_id, autor, titulo, assunto)
        )
        
        conn.commit()
        return True, "Ficha registrada e 1 crédito consumido com sucesso."
    except Exception as e:
        return False, f"Erro ao processar registro: {e}"

def adicionar_creditos(email_destino, quantidade):
    """Adiciona ou remove créditos de um usuário pelo e-mail."""
    conn = conectar_turso()
    cursor = conn.cursor()
    
    email_limpo = email_destino.lower().strip()
    res = cursor.execute("SELECT id, creditos FROM usuarios WHERE email = ?", (email_limpo,))
    row = res.fetchone()
    
    if not row:
        return False, "Usuário não encontrado."
        
    try:
        cursor.execute(
            "UPDATE usuarios SET creditos = creditos + ? WHERE email = ?",
            (quantidade, email_limpo)
        )
        conn.commit()
        return True, f"Sucesso! {quantidade} crédito(s) adicionado(s) para {email_limpo}."
    except Exception as e:
        return False, f"Erro ao atualizar créditos: {e}"

def listar_usuarios():
    """Retorna a lista de todos os usuários cadastrados."""
    conn = conectar_turso()
    cursor = conn.cursor()
    res = cursor.execute("SELECT id, nome, email, creditos, is_admin FROM usuarios ORDER BY criado_em DESC")
    return res.fetchall()
