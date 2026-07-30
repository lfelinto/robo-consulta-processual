# -*- coding: utf-8 -*-
"""
Robô de Consulta Processual — SEBRAE Contencioso (versão web / Streamlit).

Uso: o usuário abre o link, envia a planilha CSV de processos, clica em
"Iniciar consulta" e baixa o relatório Excel ao final.

Roda em Google Cloud Run (ver README_DEPLOY.md).
"""

import hmac
from datetime import datetime

import streamlit as st

import core

st.set_page_config(
    page_title="Consulta Processual — SEBRAE",
    page_icon="⚖️",
    layout="centered",
)


# ── Proteção por senha (definida em st.secrets["APP_PASSWORD"]) ─────────────
def checar_senha() -> bool:
    senha_config = st.secrets.get("APP_PASSWORD", "")
    if not senha_config:
        return True  # sem senha configurada → acesso livre (ex.: rodando local)
    if st.session_state.get("autenticado"):
        return True

    st.markdown("### 🔒 Acesso restrito")
    with st.form("login"):
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar", type="primary")
    if entrar:
        if hmac.compare_digest(senha, senha_config):
            st.session_state.autenticado = True
            st.rerun()
        st.error("Senha incorreta.")
    return False


if not checar_senha():
    st.stop()

# ── Cabeçalho ────────────────────────────────────────────────────────────────
st.markdown(
    """
    <div style="background:#1A2B5C;padding:18px 22px;border-radius:10px;margin-bottom:18px">
      <div style="color:white;font-size:22px;font-weight:700">⚖️ Consulta Processual — SEBRAE Contencioso</div>
      <div style="color:#C9D4EF;font-size:14px;margin-top:4px">
        Verifica na API pública do CNJ (DataJud) quais processos estão arquivados definitivamente.
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Chave da API (preparada uma vez por sessão) ──────────────────────────────
if "headers" not in st.session_state:
    with st.status("Preparando a chave de acesso à API do CNJ...", expanded=False) as status:
        logs = []
        def _log(m):
            logs.append(m)
            status.write(m)
        st.session_state.headers = core.preparar_headers(_log)
        status.update(label="✅ Chave de acesso pronta.", state="complete")

# ── Passo 1: enviar CSV ──────────────────────────────────────────────────────
st.subheader("1. Envie a planilha de processos (CSV)")
arquivo = st.file_uploader(
    "Arraste ou clique para escolher o arquivo",
    type=["csv"],
    help="A planilha deve ter as colunas do modelo (ex.: CNJ_Extraido, Estado, etc.).",
)

# ── Passo 2: rodar ───────────────────────────────────────────────────────────
st.subheader("2. Rodar a consulta")

if arquivo is None:
    st.info("Envie um arquivo CSV acima para habilitar a consulta.")
else:
    st.caption(f"Arquivo selecionado: **{arquivo.name}**")
    if st.button("▶ Iniciar consulta", type="primary", use_container_width=True):
        try:
            df = core.ler_csv(arquivo)
        except Exception as e:
            st.error(f"Não foi possível ler o CSV: {e}")
            st.stop()

        st.write(f"**{len(df)} linhas** carregadas. Iniciando consulta ao DataJud/CNJ...")
        st.caption("⏳ Cada processo leva ~0,7s (limite do CNJ). ~500 processos ≈ 6 a 10 min. "
                   "Mantenha esta aba aberta.")

        barra = st.progress(0.0, text="Consultando...")
        painel = st.empty()

        def progresso(feitos, total, cont):
            pct = feitos / total if total else 1.0
            barra.progress(min(pct, 1.0), text=f"Consultando... {feitos}/{total} ({int(pct*100)}%)")
            painel.markdown(
                f"✅ Arquivados: **{cont['ARQUIVADO']}**  |  "
                f"❌ Em andamento: **{cont['EM ANDAMENTO']}**  |  "
                f"❓ Não encontrados: **{cont['NÃO ENCONTRADO']}**  |  "
                f"⚠️ Erros/negados: **{cont['ACESSO NEGADO'] + cont['ERRO']}**"
            )

        inicio = datetime.now()
        df_res, cont = core.executar_consulta(
            df, st.session_state.headers,
            progresso_cb=progresso, log_cb=lambda m: None,
        )
        barra.progress(1.0, text="Concluído!")

        if df_res.empty:
            st.warning("Nenhum resultado gerado.")
            st.stop()

        fim = datetime.now()
        dur = int((fim - inicio).total_seconds())
        arquivados = (df_res["ARQUIVADO DEFINITIVAMENTE"] == "✅ SIM").sum()
        andamento = (df_res["ARQUIVADO DEFINITIVAMENTE"] == "❌ NÃO").sum()
        verificar = len(df_res) - arquivados - andamento

        st.success(f"Consulta concluída em {dur // 60}min {dur % 60}s.")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total", len(df_res))
        c2.metric("✅ Arquivados", int(arquivados))
        c3.metric("❌ Em andamento", int(andamento))
        c4.metric("❓ Verificar", int(verificar))

        xlsx = core.gerar_xlsx(df_res, fim)
        nome = f"Resultado_Consulta_{fim.strftime('%Y-%m-%d_%H%M')}.xlsx"

        st.subheader("3. Baixar o resultado")
        st.download_button(
            "📊 Baixar relatório Excel",
            data=xlsx,
            file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        with st.expander("Ver prévia dos resultados"):
            st.dataframe(df_res, use_container_width=True, hide_index=True)
