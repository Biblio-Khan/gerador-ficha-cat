from datetime import datetime
import json
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
import streamlit as st

# Configuração da página
st.set_page_config(page_title="Admin - Créditos", page_icon="🛠️", layout="wide")


# Função de Conexão com o Firestore
@st.cache_resource
def init_firebase():
  if not firebase_admin._apps:
    if "firebase" in st.secrets:
      key_dict = dict(st.secrets["firebase"])
      creds = credentials.Certificate(key_dict)
      firebase_admin.initialize_app(creds)
    else:
      creds = credentials.Certificate("credenciais.json")
      firebase_admin.initialize_app(creds)
  return firestore.client()


st.title("🛠️ Painel de Gestão e Análise (Firestore)")

# --- LOGIN SIMPLES ---
password = st.sidebar.text_input("Senha Admin:", type="password")

if password == st.secrets["admin_senha"]:
  try:
    db = init_firebase()

    # 1. Carrega usuários do Firestore
    usuarios_ref = db.collection("usuarios").stream()
    lista_usuarios = []
    for doc in usuarios_ref:
      dados = doc.to_dict()
      lista_usuarios.append({
          "uid": doc.id,
          "email": dados.get("email", "N/A"),
          "creditos": int(dados.get("creditos", 0)),
      })

    df = pd.DataFrame(lista_usuarios)
    if df.empty:
      df = pd.DataFrame(columns=["uid", "email", "creditos"])

    # --- CRIAÇÃO DAS ABAS ---
    aba_gestao, aba_graficos = st.tabs([
        "🛠️ Gestão e Recargas",
        "📊 Análise de Usuários",
    ])

    # --- ABA 1: GESTÃO E RECARGAS ---
    with aba_gestao:
      st.subheader("📋 Usuários Cadastrados")
      st.dataframe(df, use_container_width=True)

      st.divider()
      st.subheader("⚡ Ações Rápidas")

      email_input = st.text_input("E-mail do usuário")
      qtd = st.number_input("Créditos para recarga", value=10, min_value=1)

      if st.button("Recarregar Créditos"):
        if not email_input.strip():
          st.error("Digite um e-mail válido.")
        else:
          email_busca = email_input.lower().strip()
          # Busca o usuário pelo e-mail na coleção 'usuarios'
          query = (
              db.collection("usuarios")
              .where("email", "==", email_busca)
              .limit(1)
              .stream()
          )

          user_doc = None
          for doc in query:
            user_doc = doc

          if user_doc:
            dados_user = user_doc.to_dict()
            saldo_atual = int(dados_user.get("creditos", 0))
            novo_saldo = saldo_atual + int(qtd)

            # 1. Atualiza o saldo no documento do usuário
            user_doc.reference.update({"creditos": novo_saldo})

            # 2. Registra na aba/coleção Histórico automaticamente no Firestore
            try:
              db.collection("historico_recargas").add({
                  "data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                  "email": email_busca,
                  "quantidade": int(qtd),
                  "saldo_anterior": saldo_atual,
                  "novo_saldo": novo_saldo,
              })
            except Exception as e:
              st.warning(
                  f"Erro ao registrar histórico de recarga no Firestore: {e}"
              )

            st.success(
                f"Recarga feita! Saldo do usuário foi de {saldo_atual} para"
                f" {novo_saldo}."
            )
            st.rerun()
          else:
            st.error("Usuário não encontrado.")

    # --- ABA 2: ANÁLISE DE USUÁRIOS E GRÁFICOS ---
    with aba_graficos:
      st.subheader("📊 Relatórios e Comportamento de Compra")

      # Métricas gerais
      if not df.empty:
        total_usuarios = len(df)
        total_creditos = int(df["creditos"].sum())

        m1, m2 = st.columns(2)
        with m1:
          st.metric("Total de Usuários Cadastrados", total_usuarios)
        with m2:
          st.metric("Créditos Totais em Circulação", total_creditos)

      st.divider()

      # Leitura da coleção Historico Recargas para gráficos
      try:
        hist_ref = db.collection("historico_recargas").stream()
        lista_hist = []
        for doc in hist_ref:
          lista_hist.append(doc.to_dict())

        if lista_hist:
          df_hist = pd.DataFrame(lista_hist)

          # Expander para conferência dos dados brutos
          with st.expander("Ver dados brutos do Histórico de Recargas"):
            st.dataframe(df_hist)

          st.markdown("### 📈 Histórico de Recargas por Usuário (Volume)")

          if "email" in df_hist.columns and "quantidade" in df_hist.columns:
            ranking = (
                df_hist.groupby("email")["quantidade"]
                .sum()
                .sort_values(ascending=False)
            )
            st.bar_chart(ranking)
          else:
            st.warning(
                "As colunas de 'email' ou 'quantidade' não foram encontradas"
                " corretamente no histórico."
            )
        else:
          st.info(
              "A coleção 'historico_recargas' está vazia no momento. Faça uma"
              " recarga para gerar dados."
          )
      except Exception as e:
        st.error(f"Erro ao ler histórico do Firestore: {e}")

  except Exception as e:
    st.error(f"Erro ao conectar ou carregar dados do Firestore: {e}")
else:
  st.info("Insira a senha correta na barra lateral para acessar o painel.")
