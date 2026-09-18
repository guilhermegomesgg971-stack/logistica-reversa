from datetime import datetime
import io
import os
from docx import Document
from docx.shared import Inches, Pt
import pandas as pd
import streamlit as st

# Configuração da Página & Tema Logístico
st.set_page_config(
    page_title="WMS & Logística Reversa - RMC Mariano",
    page_icon="🏭",
    layout="wide",
)

# Arquivo local para salvar os dados
DB_FILE = "dados_reversas.csv"


# Função para carregar os dados com segurança e tipos corretos
def carregar_dados():
  colunas_necessarias = [
      "ID_Devolucao",
      "Data_Registro",
      "Pedido",
      "NF",
      "Cliente",
      "Cidade",
      "Transportadora",
      "Motivo",
      "Itens",
      "Status",
      "Data_Autorizacao",
      "Entregador_Responsavel",
      "Observacao_Transportadora",
      "Data_Conclusao",
  ]

  if os.path.exists(DB_FILE):
    df = pd.read_csv(DB_FILE, dtype=str)
    for col in colunas_necessarias:
      if col not in df.columns:
        df[col] = ""
      else:
        df[col] = df[col].fillna("")
    return df
  else:
    df_padrao = pd.DataFrame(columns=colunas_necessarias, dtype=str)
    return df_padrao


# Função para salvar os dados
def salvar_dados(df):
  df.to_csv(DB_FILE, index=False)


# Carrega o banco de dados na sessão
if "df_reversas" not in st.session_state:
  st.session_state.df_reversas = carregar_dados()

# Cabeçalho com identidade visual de Logística / Centro de Distribuição
st.markdown(
    """
    <div style="background-color: #0f172a; padding: 20px; border-radius: 8px; border-left: 6px solid #f59e0b;">
        <h2 style="color: white; margin: 0; font-family: sans-serif;">🏭 RMC MARIANO | Gestão de Logística Reversa & WMS</h2>
        <p style="color: #94a3b8; margin: 5px 0 0 0; font-size: 14px;">Controle de fluxo de coletas, rastreio de transportadora e conferência física no CD.</p>
    </div>
""",
    unsafe_allow_html=True,
)
st.markdown("<br>", unsafe_allow_html=True)

# Criação das Abas Principais com nomenclaturas de Supply Chain
aba_gestor, aba_transportadora, aba_relatorios, aba_termo = st.tabs(
    [
        "📋 1. Painel Operacional (CD & Cadastro)",
        "🚚 2. Dock / Portal da Transportadora",
        "📊 3. Dashboard & Indicadores",
        "📄 4. Emissão de Termos (Word)",
    ]
)

# ==========================================
# ABA 1: PAINEL DO LÍDER (CADASTRO & BAIXA)
# ==========================================
with aba_gestor:
  st.subheader("📝 Abertura de Ordem de Coleta Reversa")

  with st.form("form_cadastro_reversa", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)

    with col1:
      num_pedido = st.text_input("Nº do Pedido / Carga")
      num_nf = st.text_input("Nota Fiscal (NF)")

    with col2:
      cliente = st.text_input("Nome da Revendedora / Cliente Final")
      cidade = st.text_input("Destino / Cidade")

    with col3:
      transportadora = "Transportadora José Augusto"
      st.text_input(
          "Operador Logístico / Transportadora",
          value=transportadora,
          disabled=True,
      )
      motivo = st.selectbox(
          "Motivo da Ocorrência",
          [
              "Produto Avariado",
              "Caixa Trocada",
              "Item Faltante no Pedido",
              "Desistência / Erro Operacional",
          ],
      )

    itens = st.text_area(
        "SKUs / Descrição dos Itens a Serem Coletados",
        placeholder="Ex: 1x Colônia Floratta Red, 2x Hidratante Instance...",
    )

    submitted = st.form_submit_button(
        "🚀 Emitir Ordem e Liberar para Coleta", use_container_width=True
    )

    if submitted:
      if not num_pedido or not cliente:
        st.warning(
            "⚠️ Preencha os campos obrigatórios (Nº do Pedido e Cliente)!"
        )
      else:
        novo_id = f"REV-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        nova_linha = {
            "ID_Devolucao": str(novo_id),
            "Data_Registro": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Pedido": str(num_pedido),
            "NF": str(num_nf) if num_nf else "N/A",
            "Cliente": str(cliente).upper(),
            "Cidade": str(cidade).upper(),
            "Transportadora": str(transportadora),
            "Motivo": str(motivo),
            "Itens": str(itens),
            "Status": "🟡 Aguardando Coleta",
            "Data_Autorizacao": "",
            "Entregador_Responsavel": "",
            "Observacao_Transportadora": "",
            "Data_Conclusao": "",
        }

        st.session_state.df_reversas = pd.concat(
            [st.session_state.df_reversas, pd.DataFrame([nova_linha])],
            ignore_index=True,
        )
        salvar_dados(st.session_state.df_reversas)
        st.success(
            f"✅ Ordem **{novo_id}** gerada e disponibilizada para o operador"
            " logístico!"
        )

  st.markdown("---")
  st.subheader(
      "📥 Conferência de Recebimento no CD (Baixa de Cargas Retornadas)"
  )

  df_atual = st.session_state.df_reversas
  if not df_atual.empty:
    coletados_para_baixar = df_atual[
        df_atual["Status"] == "🟢 Coletado / Autorizado"
    ]

    if not coletados_para_baixar.empty:
      st.markdown(
          "#### ⚡ Lotes Coletados em Trânsito (Aguardando conferência física"
          " no CD):"
      )
      for index, row in coletados_para_baixar.iterrows():
        with st.container(border=True):
          col_c1, col_c2 = st.columns([3, 1])
          with col_c1:
            st.markdown(
                f"**ID:** `{row['ID_Devolucao']}` | **Pedido:**"
                f" `{row['Pedido']}` | **NF:** `{row['NF']}`"
            )
            st.markdown(
                f"👤 **Cliente:** {row['Cliente']} ({row['Cidade']})"
            )
            st.markdown(
                f"🚚 **Motorista:** `{row['Entregador_Responsavel']}` em"
                f" `{row['Data_Autorizacao']}`"
            )
            st.markdown(f"📦 **Itens Esperados:** {row['Itens']}")
          with col_c2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button(
                "🏁 Validar & Dar Entrada no CD",
                key=f"finalizar_{row['ID_Devolucao']}",
            ):
              idx_real = st.session_state.df_reversas[
                  st.session_state.df_reversas["ID_Devolucao"]
                  == row["ID_Devolucao"]
              ].index[0]
              st.session_state.df_reversas.at[
                  idx_real, "Status"
              ] = "✅ Concluído / Resolvido no CD"
              st.session_state.df_reversas.at[
                  idx_real, "Data_Conclusao"
              ] = datetime.now().strftime("%d/%m/%Y %H:%M")
              salvar_dados(st.session_state.df_reversas)
              st.success(
                  "Carga validada e processada no estoque com sucesso!"
              )
              st.rerun()
    else:
      st.info(
          "ℹ️ Não há cargas em trânsito aguardando conferência no CD no"
          " momento."
      )

    st.markdown("---")
    st.subheader("📊 Histórico Geral & Auditoria de Cargas")
    filtro_status = st.selectbox(
        "Filtrar por Status Operacional",
        [
            "Todos",
            "🟡 Aguardando Coleta",
            "🟢 Coletado / Autorizado",
            "✅ Concluído / Resolvido no CD",
            "🔴 Com Ocorrência/Problema",
        ],
    )
    if filtro_status != "Todos":
      df_filtrado = df_atual[df_atual["Status"] == filtro_status]
    else:
      df_filtrado = df_atual

    st.dataframe(df_filtrado, use_container_width=True)

    st.markdown("### 🗑️ Gestão de Registros (Exclusão Direta)")
    st.markdown(
        "*(Clique na lixeira para remover definitivamente um registro do"
        " sistema)*"
    )

    if not df_filtrado.empty:
      for index, row in df_filtrado.iterrows():
        col_reg1, col_reg2, col_reg3 = st.columns([1.5, 4, 1])
        with col_reg1:
          st.text(row["ID_Devolucao"])
        with col_reg2:
          st.text(
              f"Ped: {row['Pedido']} | Cliente: {row['Cliente']} ({row['Status']})"
          )
        with col_reg3:
          if st.button("🗑️ Excluir", key=f"lixeira_{row['ID_Devolucao']}"):
            st.session_state.df_reversas = st.session_state.df_reversas[
                st.session_state.df_reversas["ID_Devolucao"]
                != row["ID_Devolucao"]
            ].reset_index(drop=True)
            salvar_dados(st.session_state.df_reversas)
            st.success(f"Registro {row['ID_Devolucao']} excluído com sucesso!")
            st.rerun()
    else:
      st.info("Nenhum registro para gerenciar nesta visualização.")

  else:
    st.info("Nenhuma ordem de devolução registrada no sistema.")

# ==========================================
# ABA 2: PORTAL DA TRANSPORTADORA
# ==========================================
with aba_transportadora:
  st.subheader("🚚 Portal do Operador Logístico (Transportadora José Augusto)")
  st.markdown(
      "Painel de despacho de motoristas para recolhimento nas rotas."
  )

  transp_selecionada = "Transportadora José Augusto"
  df_transp = st.session_state.df_reversas[
      st.session_state.df_reversas["Transportadora"] == transp_selecionada
  ]

  cidades_disponiveis = ["Todas"] + list(df_transp["Cidade"].unique())
  filtro_cidade_transp = st.selectbox(
      "Filtrar por Rota / Cidade:", cidades_disponiveis, key="filtro_cid_transp"
  )

  if filtro_cidade_transp != "Todas":
    pendentes = df_transp[
        (df_transp["Status"] == "🟡 Aguardando Coleta")
        & (df_transp["Cidade"] == filtro_cidade_transp)
    ]
  else:
    pendentes = df_transp[df_transp["Status"] == "🟡 Aguardando Coleta"]

  st.markdown("---")
  st.markdown(
      f"### 📦 Ordens Liberadas para Coleta ({len(pendentes)} ordens na fila)"
  )

  if not pendentes.empty:
    for index, row in pendentes.iterrows():
      with st.container(border=True):
        col_a, col_b = st.columns([2.5, 1.5])
        with col_a:
          st.markdown(
              f"**ID:** `{row['ID_Devolucao']}` | **Emissão:**"
              f" `{row['Data_Registro']}`"
          )
          st.markdown(
              f"📄 **Pedido:** `{row['Pedido']}` | **NF:** `{row['NF']}`"
          )
          st.markdown(
              f"👤 **Cliente:** **{row['Cliente']}** (📍 {row['Cidade']})"
          )
          st.markdown(f"❓ **Motivo:** `{row['Motivo']}`")
          st.markdown(f"📦 **Itens a Coletar:** {row['Itens']}")

        with col_b:
          st.markdown("##### ✍️ Check-in de Retirada")
          nome_entregador = st.text_input(
              "Motorista Responsável:",
              key=f"ent_{row['ID_Devolucao']}",
              placeholder="Nome do motorista",
          )

          if st.button(
              "✅ Confirmar Retirada", key=f"conf_{row['ID_Devolucao']}"
          ):
            if not nome_entregador.strip():
              st.error("⚠️ Informe o nome do motorista!")
            else:
              idx_real = st.session_state.df_reversas[
                  st.session_state.df_reversas["ID_Devolucao"]
                  == row["ID_Devolucao"]
              ].index[0]
              st.session_state.df_reversas.at[
                  idx_real, "Status"
              ] = "🟢 Coletado / Autorizado"
              st.session_state.df_reversas.at[
                  idx_real, "Data_Autorizacao"
              ] = datetime.now().strftime("%d/%m/%Y %H:%M")
              st.session_state.df_reversas.at[
                  idx_real, "Entregador_Responsavel"
              ] = str(nome_entregador).upper()
              salvar_dados(st.session_state.df_reversas)
              st.success("Retirada registrada com sucesso!")
              st.rerun()

          st.markdown("---")
          obs_transp = st.text_input(
              "Motivo da Ocorrência:",
              key=f"obs_{row['ID_Devolucao']}",
              placeholder="Ex: Cliente ausente",
          )
          if st.button(
              "⚠️ Registrar Ocorrência", key=f"prob_{row['ID_Devolucao']}"
          ):
            if obs_transp:
              idx_real = st.session_state.df_reversas[
                  st.session_state.df_reversas["ID_Devolucao"]
                  == row["ID_Devolucao"]
              ].index[0]
              st.session_state.df_reversas.at[
                  idx_real, "Status"
              ] = "🔴 Com Ocorrência/Problema"
              st.session_state.df_reversas.at[
                  idx_real, "Observacao_Transportadora"
              ] = str(obs_transp)
              salvar_dados(st.session_state.df_reversas)
              st.warning("Ocorrência enviada ao CD!")
              st.rerun()
            else:
              st.error("Digite o motivo ao lado.")
  else:
    st.info("🎉 Nenhuma ordem pendente nas rotas selecionadas.")

  st.markdown("---")
  st.markdown("### 📜 Histórico de Despachos Realizados")
  concluidas = df_transp[df_transp["Status"] != "🟡 Aguardando Coleta"]
  if not concluidas.empty:
    st.dataframe(concluidas, use_container_width=True)
  else:
    st.info("Nenhum histórico operacional registrado.")

# ==========================================
# ABA 3: INDICADORES & ALERTAS
# ==========================================
with aba_relatorios:
  st.subheader("📊 Dashboard de Performance Logística (KPIs)")
  df_geral = st.session_state.df_reversas

  if not df_geral.empty:
    total_geral = len(df_geral)
    total_aguardando = len(
        df_geral[df_geral["Status"] == "🟡 Aguardando Coleta"]
    )
    total_coletados = len(
        df_geral[df_geral["Status"] == "🟢 Coletado / Autorizado"]
    )
    total_resolvidos = len(
        df_geral[df_geral["Status"] == "✅ Concluído / Resolvido no CD"]
    )
    total_problemas = len(
        df_geral[df_geral["Status"] == "🔴 Com Ocorrência/Problema"]
    )

    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
    col_m1.metric("Total Ordens", total_geral)
    col_m2.metric("Aguardando Coleta", total_aguardando)
    col_m3.metric("Em Trânsito", total_coletados)
    col_m4.metric("Resolvidos no CD", total_resolvidos)
    col_m5.metric("Ocorrências", total_problemas)
  else:
    st.info("Insira dados para visualizar os indicadores do painel.")

# ==========================================
# ABA 4: GERADOR DE TERMO DE RESPONSABILIDADE (WORD)
# ==========================================
with aba_termo:
  st.subheader("📄 Geração de Documentação Oficial (Termo de Coleta)")
  st.markdown(
      "Emita o termo impresso de responsabilidade para assinatura no"
      " recebimento."
  )

  df_termo_geral = st.session_state.df_reversas

  if not df_termo_geral.empty:
    lista_ids = df_termo_geral["ID_Devolucao"].tolist()
    id_escolhido = st.selectbox(
        "Selecione o ID da Ordem para Gerar o Documento:", lista_ids
    )

    dados_linha = df_termo_geral[
        df_termo_geral["ID_Devolucao"] == id_escolhido
    ].iloc[0]

    st.markdown("---")
    st.markdown("#### 🔍 Prévia dos Dados para Impressão:")
    c_t1, c_t2, c_t3 = st.columns(3)
    c_t1.text_input("ID", value=dados_linha["ID_Devolucao"], disabled=True)
    c_t2.text_input("Pedido", value=dados_linha["Pedido"], disabled=True)
    c_t3.text_input("Nota Fiscal", value=dados_linha["NF"], disabled=True)

    st.text_input("Cliente", value=dados_linha["Cliente"], disabled=True)
    st.text_input("Motivo", value=dados_linha["Motivo"], disabled=True)
    st.text_area("Itens", value=dados_linha["Itens"], disabled=True)

    if st.button("📥 Baixar Termo em Formato Word (.docx)", use_container_width=True):
      doc = Document()

      if os.path.exists("logo.png"):
        try:
          doc.add_picture("logo.png", width=Inches(2.0))
        except:
          pass

      p_titulo = doc.add_paragraph()
      run_titulo = p_titulo.add_run(
          "GRUPO RMC MARIANO - TERMO DE RESPONSABILIDADE DE DEVOLUÇÃO"
      )
      run_titulo.bold = True
      run_titulo.font.size = Pt(14)

      doc.add_paragraph(f"Data de Emissão: {dados_linha['Data_Registro']}")
      doc.add_paragraph(f"ID da Ocorrência: {dados_linha['ID_Devolucao']}")
      doc.add_paragraph(f"Nº do Pedido: {dados_linha['Pedido']}")
      doc.add_paragraph(f"Nota Fiscal: {dados_linha['NF']}")
      doc.add_paragraph(f"Revendedora / Cliente: {dados_linha['Cliente']}")
      doc.add_paragraph(f"Cidade: {dados_linha['Cidade']}")
      doc.add_paragraph(f"Motivo da Devolução: {dados_linha['Motivo']}")

      doc.add_paragraph("\nITENS A SEREM DEVOLVIDOS / COLETADOS:")
      doc.add_paragraph(str(dados_linha["Itens"]))

      doc.add_paragraph(
          "\nDeclaro para os devidos fins que os produtos acima descritos"
          " estão sendo entregues à Transportadora José Augusto para retorno"
          " ao Centro de Distribuição do Grupo RMC Mariano."
      )

      doc.add_paragraph("\n\n__________________________________________________")
      doc.add_paragraph(f"Assinatura da Revendedora: {dados_linha['Cliente']}")

      doc.add_paragraph("__________________________________________________")
      doc.add_paragraph("Assinatura / Carimbo da Transportadora José Augusto")

      buffer = io.BytesIO()
      doc.save(buffer)
      buffer.seek(0)

      st.download_button(
          label="💾 Clique aqui para baixar o documento Word gerado",
          data=buffer,
          file_name=f"Termo_Devolucao_{dados_linha['ID_Devolucao']}.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
      st.success("Documento gerado com sucesso!")
  else:
    st.info("Insira ordens nas abas anteriores para gerar documentos.")