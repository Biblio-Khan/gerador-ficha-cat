import streamlit as st
import libsql_client

def conectar_turso():
    """Estabelece a conexão com o banco Turso."""
    turso_secrets = st.secrets.get("turso", {})
    url = turso_secrets.get("TURSO_DATABASE_URL") or st.secrets.get("TURSO_DATABASE_URL")
    auth_token = turso_secrets.get("TURSO_AUTH_TOKEN") or st.secrets.get("TURSO_AUTH_TOKEN")
    
    if not url or not auth_token:
        st.error("⚠️ Credenciais do Turso ausentes no st.secrets.")
        st.stop()
        
    try:
        return libsql_client.create_client_sync(url=url, auth_token=auth_token)
    except Exception as e:
        st.error(f"❌ Erro ao conectar no Turso: {e}")
        st.stop()

def autenticar_usuario(email, senha):
    """Verifica login e retorna os dados do usuário."""
    client = conectar_turso()
    try:
        # Busca pela coluna senha_hash
        result = client.execute(
            "SELECT id, nome, email, creditos, is_admin, senha_hash FROM usuarios WHERE email = ?", 
            (email,)
        )
        if result.rows:
            row = result.rows[0]
            if str(row[5]) == str(senha):  # Confirma a senha
                user_data = {
                    "id": row[0],
                    "nome": row[1],
                    "email": row[2],
                    "creditos": row[3],
                    "is_admin": row[4]
                }
                return True, "Login realizado com sucesso!", user_data
            else:
                return False, "Senha incorreta.", None
        return False, "Usuário não encontrado.", None
    finally:
        client.close()

def criar_usuario(nome, email, senha):
    """Cadastra um novo usuário no sistema com 4 créditos iniciais."""
    client = conectar_turso()
    try:
        # Verifica se o e-mail já existe
        check = client.execute("SELECT id FROM usuarios WHERE email = ?", (email,))
        if check.rows:
            return False, "E-mail já cadastrado!"
        
        # Insere na coluna senha_hash e dá 4 créditos
        client.execute(
            "INSERT INTO usuarios (nome, email, senha_hash, creditos, is_admin) VALUES (?, ?, ?, ?, ?)",
            (nome, email, senha, 4, 0)
        )
        return True, "Cadastro realizado com sucesso!"
    except Exception as e:
        return False, f"Erro ao cadastrar: {e}"
    finally:
        client.close()

def listar_usuarios():
    """Lista todos os usuários para o painel do Admin."""
    client = conectar_turso()
    try:
        result = client.execute("SELECT id, nome, email, creditos, is_admin FROM usuarios")
        return result.rows
    except Exception as e:
        st.error(f"Erro ao listar usuários: {e}")
        return []
    finally:
        client.close()

def adicionar_creditos(email, quantidade):
    """Admin adiciona créditos a um usuário específico."""
    client = conectar_turso()
    try:
        check = client.execute("SELECT creditos FROM usuarios WHERE email = ?", (email,))
        if not check.rows:
            return False, "Usuário não encontrado."
        
        saldo_atual = check.rows[0][0]
        novo_saldo = saldo_atual + quantidade
        
        client.execute(
            "UPDATE usuarios SET creditos = ? WHERE email = ?", 
            (novo_saldo, email)
        )
        return True, f"Créditos atualizados! Novo saldo: {novo_saldo} fichas."
    except Exception as e:
        return False, f"Erro ao adicionar créditos: {e}"
    finally:
        client.close()

def descontar_credito_e_registrar(email, usuario_id, autor, titulo, assunto):
    """
    Desconta 1 crédito e registra a ficha gerada.
    """
    client = conectar_turso()
    try:
        check = client.execute("SELECT creditos FROM usuarios WHERE email = ?", (email,))
        if check.rows:
            saldo_atual = check.rows[0][0]
            
            if saldo_atual > 0:
                novo_saldo = saldo_atual - 1
                
                # Desconta o crédito
                client.execute(
                    "UPDATE usuarios SET creditos = ? WHERE email = ?", 
                    (novo_saldo, email)
                )
                
                # Registra na tabela produtividade
                client.execute(
                    "INSERT INTO produtividade (usuario_id, autor, titulo, assunto) VALUES (?, ?, ?, ?)",
                    (usuario_id, autor, titulo, assunto)
                )
                
                return True, novo_saldo
                
        return False, 0
    except Exception as e:
        print(f"Erro ao descontar crédito e registrar: {e}")
        return False, 0
    finally:
        client.close()

def listar_produtividade():
    """Lista o histórico de fichas para o Painel Admin."""
    client = conectar_turso()
    try:
        result = client.execute("""
            SELECT p.id, u.nome, u.email, p.autor, p.titulo, p.assunto, p.data_registro 
            FROM produtividade p
            JOIN usuarios u ON p.usuario_id = u.id
            ORDER BY p.data_registro DESC
        """)
        return result.rows
    except Exception as e:
        print(f"Erro ao listar produtividade: {e}")
        return []
    finally:
        client.close()
