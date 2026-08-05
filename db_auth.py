import streamlit as st
import libsql_client
import bcrypt

def conectar_turso():
    # Lê diretamente da raiz do st.secrets
    url = st.secrets.get("TURSO_DATABASE_URL")
    auth_token = st.secrets.get("TURSO_AUTH_TOKEN")
    
    if not url or not auth_token:
        st.error("⚠️ Credenciais do Turso não foram encontradas no st.secrets.")
        st.stop()
        
    return libsql.connect(database=url, auth_token=auth_token)

def cadastrar_usuario(email, nome, senha):
    """Cadastra um novo usuário no Turso. Ganha 4 créditos automaticamente."""
    client = conectar_turso()
    
    senha_bytes = senha.encode('utf-8')
    salt = bcrypt.gensalt()
    senha_hash = bcrypt.hashpw(senha_bytes, salt).decode('utf-8')
    
    try:
        client.execute(
            "INSERT INTO usuarios (email, nome, senha_hash) VALUES (?, ?, ?)",
            [email.lower().strip(), nome.strip(), senha_hash]
        )
        return True, "Usuário cadastrado com sucesso! Você ganhou 4 créditos de boas-vindas."
    except Exception as e:
        if "UNIQUE" in str(e).upper():
            return False, "Este e-mail já está cadastrado."
        return False, f"Erro ao cadastrar: {e}"

def autenticar_usuario(email, senha):
    """Valida o login do usuário e retorna seus dados (incluindo créditos)."""
    client = conectar_turso()
    
    res = client.execute(
        "SELECT id, nome, email, senha_hash, creditos, is_admin FROM usuarios WHERE email = ?",
        [email.lower().strip()]
    )
    
    if not res.rows:
        return False, "E-mail ou senha incorretos.", None
        
    usuario = res.rows[0]
    user_id, nome, user_email, senha_hash_db, creditos, is_admin = usuario
    
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
    """Desconta 1 crédito do usuário e salva a ficha na tabela de produtividade."""
    client = conectar_turso()
    
    res = client.execute("SELECT creditos FROM usuarios WHERE id = ?", [usuario_id])
    
    if not res.rows or res.rows[0][0] <= 0:
        return False, "Você não possui créditos suficientes para gerar a ficha."
        
    try:
        client.execute("UPDATE usuarios SET creditos = creditos - 1 WHERE id = ?", [usuario_id])
        client.execute(
            "INSERT INTO registros_produtividade (usuario_id, autor, titulo, assunto) VALUES (?, ?, ?, ?)",
            [usuario_id, autor, titulo, assunto]
        )
        return True, "Ficha registrada e 1 crédito consumido com sucesso."
    except Exception as e:
        return False, f"Erro ao processar registro: {e}"

def adicionar_creditos(email_destino, quantidade):
    """Adiciona ou remove créditos de um usuário pelo e-mail."""
    client = conectar_turso()
    
    email_limpo = email_destino.lower().strip()
    res = client.execute("SELECT id, creditos FROM usuarios WHERE email = ?", [email_limpo])
    
    if not res.rows:
        return False, "Usuário não encontrado."
        
    try:
        client.execute(
            "UPDATE usuarios SET creditos = creditos + ? WHERE email = ?",
            [quantidade, email_limpo]
        )
        return True, f"Sucesso! {quantidade} crédito(s) adicionado(s) para {email_limpo}."
    except Exception as e:
        return False, f"Erro ao atualizar créditos: {e}"

def listar_usuarios():
    """Retorna a lista de todos os usuários cadastrados."""
    client = conectar_turso()
    res = client.execute("SELECT id, nome, email, creditos, is_admin FROM usuarios ORDER BY criado_em DESC")
    return res.rows
