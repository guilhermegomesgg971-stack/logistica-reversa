from datetime import datetime, timedelta, timezone
import io
import os
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
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


# Função para pegar o horário exato de Brasília (-3 horas do UTC) sem erros de biblioteca
def obter_horario_brasilia():
  fuso_brasilia = timezone(timedelta(hours=-3))
  return datetime.now(fuso_brasilia).strftime("%d/%m/%Y %H:%M")


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
      "Data_Inicio_Coleta",
      "Entregador_Responsavel",
      "Evidencia_Anexo",
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
        <p style="color: #94a3b8; margin: 5px 0 0 0; font-size: 14px;">Controle de fluxo de coletas, rastreio de transportadora e monitoramento de status.</p>
    </div>
""",
    unsafe_allow_html=True,
)
st.markdown("<br>", unsafe_allow_html=True)

# Criação das Abas Principais
aba_gestor, aba_transportadora, aba_analise, aba_relatorios, aba_termo = st.tabs([
    "📋 1. Painel Operacional (Cadastro & Histórico)",
    "🚚 2. Portal da Transportadora",
    "🔍 3. Analisar Pedidos Finalizados",
    "📊 4. Dashboard & Indicadores",
    "📄 5. Emissão de Termos / Checklist",
])

# ==========================================
# ABA 1: PAINEL OPERACIONAL (CADASTRO E HISTÓRICO DE STATUS)
# ==========================================
with aba_gestor:
  st.subheader("📝 Abertura de Ordem de Coleta Reversa")
  st.markdown(
      "Cadastre a ordem de reversa. O pedido será enviado diretamente para a"
      " transportadora realizar a verificação e coleta."
  )

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
          ["Produto Avariado", "Troca de Caixa", "Desistência do Pedido"],
      )

    itens = st.text_area(
        "SKUs / Descrição dos Itens a Serem Coletados",
        placeholder="Ex: 1x Colônia Floratta Red, 2x Hidratante Instance...",
    )

    submitted = st.form_submit_button(
        "🚀 Emitir Ordem e Enviar para Transportadora",
        use_container_width=True,
    )

    if submitted:
      if not num_pedido or not cliente:
        st.warning(
            "⚠️ Preencha os campos obrigatórios (Nº do Pedido e Cliente)!"
        )
      else:
        fuso_brasilia = timezone(timedelta(hours=-3))
        novo_id = f"REV-{datetime.now(fuso_brasilia).strftime('%Y%m%d%H%M%S')}"
        nova_linha = {
            "ID_Devolucao": str(novo_id),
            "Data_Registro": obter_horario_brasilia(),
            "Pedido": str(num_pedido),
            "NF": str(num_nf) if num_nf else "N/A",
            "Cliente": str(cliente).upper(),
            "Cidade": str(cidade).upper(),
            "Transportadora": str(transportadora),
            "Motivo": str(motivo),
            "Itens": str(itens),
            "Status": "🟡 Aguardando Verificação",
            "Data_Inicio_Coleta": "",
            "Entregador_Responsavel": "",
            "Evidencia_Anexo": "",
            "Data_Conclusao": "",
        }

        st.session_state.df_reversas = pd.concat(
            [st.session_state.df_reversas, pd.DataFrame([nova_linha])],
            ignore_index=True,
        )
        salvar_dados(st.session_state.df_reversas)
        st.success(
            f"✅ Ordem **{novo_id}** gerada com sucesso e enviada para o portal"
            " da transportadora!"
        )

  st.markdown("---")
  st.subheader(
      "📊 Histórico Geral de Status de Cada Pedido (Acompanhamento)"
  )
  st.markdown(
      "Visualize abaixo o andamento de todas as ordens abertas e concluídas."
  )

  df_atual = st.session_state.df_reversas
  if not df_atual.empty:
    filtro_status = st.selectbox(
        "Filtrar por Status Operacional",
        [
            "Todos",
            "🟡 Aguardando Verificação",
            "🚚 Coleta Iniciada",
            "✅ Finalizado pela Transportadora",
        ],
    )
    if filtro_status != "Todos":
      df_filtrado = df_atual[df_atual["Status"] == filtro_status]
    else:
      df_filtrado = df_atual

    st.dataframe(df_filtrado, use_container_width=True)

    st.markdown("### 🗑️ Gestão de Registros (Exclusão Direta)")
    if not df_filtrado.empty:
      for index, row in df_filtrado.iterrows():
        col_reg1, col_reg2, col_reg3 = st.columns([1.5, 4, 1])
        with col_reg1:
          st.text(row["ID_Devolucao"])
        with col_reg2:
          st.text(
              f"Ped: {row['Pedido']} | Cliente: {row['Cliente']} ("
              f"{row['Status']})"
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
  st.subheader(
      "🚚 Portal do Operador Logístico (Verificação, Coleta e Envio de Caixa)"
  )
  st.markdown(
      "Gerencie os pedidos aguardando verificação, autorize e registre a"
      " data/hora da coleta pelo entregador."
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
    pendentes_transp = df_transp[
        (df_transp["Status"] != "✅ Finalizado pela Transportadora")
        & (df_transp["Cidade"] == filtro_cidade_transp)
    ]
  else:
    pendentes_transp = df_transp[
        df_transp["Status"] != "✅ Finalizado pela Transportadora"
    ]

  st.markdown("---")
  st.markdown(
      f"### 📦 Ordens Ativas no Painel ({len(pendentes_transp)} ordens)"
  )

  if not pendentes_transp.empty:
    for index, row in pendentes_transp.iterrows():
      with st.container(border=True):
        col_a, col_b = st.columns([2.5, 1.5])
        with col_a:
          st.markdown(
              f"**ID:** `{row['ID_Devolucao']}` | **Status Atual:**"
              f" **{row['Status']}**"
          )
          st.markdown(
              f"📄 **Pedido:** `{row['Pedido']}` | **NF:** `{row['NF']}`"
          )
          st.markdown(
              f"👤 **Cliente:** **{row['Cliente']}** (📍 {row['Cidade']})"
          )
          st.markdown(f"❓ **Motivo:** `{row['Motivo']}`")
          st.markdown(f"📦 **Itens a Coletar:** {row['Itens']}")
          if row["Data_Inicio_Coleta"]:
            st.markdown(
                f"⏱️ **Autorizado/Início da Coleta:**"
                f" `{row['Data_Inicio_Coleta']}` por"
                f" `{row['Entregador_Responsavel']}`"
            )

        with col_b:
          st.markdown("##### ✍️ Ações da Transportadora")

          # Ação 1: Iniciar Coleta
          if row["Status"] == "🟡 Aguardando Verificação":
            nome_entregador = st.text_input(
                "Motorista Responsável:",
                key=f"ent_{row['ID_Devolucao']}",
                placeholder="Nome do motorista",
            )
            if st.button(
                "🚀 Autorizar e Iniciar Coleta",
                key=f"btn_iniciar_{row['ID_Devolucao']}",
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
                ] = "🚚 Coleta Iniciada"
                st.session_state.df_reversas.at[
                    idx_real, "Data_Inicio_Coleta"
                ] = obter_horario_brasilia()
                st.session_state.df_reversas.at[
                    idx_real, "Entregador_Responsavel"
                ] = str(nome_entregador).upper()
                salvar_dados(st.session_state.df_reversas)
                st.success(
                    "Coleta autorizada e horário oficial registrado com"
                    " sucesso!"
                )
                st.rerun()

          # Ação 2: Finalizar Pedido com anexo de foto/vídeo
          if row["Status"] == "🚚 Coleta Iniciada":
            uploaded_file = st.file_uploader(
                "Anexar Foto ou Vídeo (Evidência):",
                type=["png", "jpg", "jpeg", "mp4", "mov"],
                key=f"file_{row['ID_Devolucao']}",
            )

            if st.button(
                "✅ Finalizar Pedido (Enviar Caixa)",
                key=f"btn_finalizar_{row['ID_Devolucao']}",
            ):
              idx_real = st.session_state.df_reversas[
                  st.session_state.df_reversas["ID_Devolucao"]
                  == row["ID_Devolucao"]
              ].index[0]

              nome_arquivo = ""
              if uploaded_file is not None:
                nome_arquivo = uploaded_file.name
                os.makedirs("uploads", exist_ok=True)
                with open(
                    os.path.join("uploads", uploaded_file.name), "wb"
                ) as f:
                  f.write(uploaded_file.getbuffer())

              st.session_state.df_reversas.at[
                  idx_real, "Status"
              ] = "✅ Finalizado pela Transportadora"
              st.session_state.df_reversas.at[
                  idx_real, "Evidencia_Anexo"
              ] = nome_arquivo
              st.session_state.df_reversas.at[
                  idx_real, "Data_Conclusao"
              ] = obter_horario_brasilia()

              salvar_dados(st.session_state.df_reversas)
              st.success(
                  "Pedido finalizado e enviado para análise no Centro de"
                  " Distribuição!"
              )
              st.rerun()
  else:
    st.info(
        "🎉 Nenhuma ordem ativa pendente na transportadora para as rotas"
        " selecionadas."
    )

# ==========================================
# ABA 3: ANALISAR PEDIDOS FINALIZADOS (COM VISUALIZAÇÃO DE ANEXOS)
# ==========================================
with aba_analise:
  st.subheader("🔍 Central de Análise de Pedidos Finalizados")
  st.markdown(
      "Aqui você visualiza os dados, confere o histórico e analisa as"
      " evidências (fotos/vídeos) enviadas pela transportadora."
  )

  df_geral = st.session_state.df_reversas
  df_finalizados = df_geral[
      df_geral["Status"] == "✅ Finalizado pela Transportadora"
  ]

  if not df_finalizados.empty:
    st.markdown(
        f"### 📥 Total de pedidos aguardando sua análise:"
        f" {len(df_finalizados)}"
    )

    for index, row in df_finalizados.iterrows():
      with st.container(border=True):
        col_an1, col_an2 = st.columns([2, 1.5])
        with col_an1:
          st.markdown(
              f"**ID:** `{row['ID_Devolucao']}` | **Conclusão:**"
              f" `{row['Data_Conclusao']}`"
          )
          st.markdown(
              f"📄 **Pedido:** `{row['Pedido']}` | **NF:** `{row['NF']}`"
          )
          st.markdown(
              f"👤 **Cliente:** **{row['Cliente']}** (📍 {row['Cidade']})"
          )
          st.markdown(f"❓ **Motivo:** `{row['Motivo']}`")
          st.markdown(f"📦 **Itens:** {row['Itens']}")
          st.markdown(
              f"⏱️ **Início/Autorização da Coleta:**"
              f" `{row['Data_Inicio_Coleta']}`"
          )
          st.markdown(
              f"🚚 **Motorista Responsável:** `{row['Entregador_Responsavel']}`"
          )

        with col_an2:
          st.markdown("##### 📎 Evidência Anexada")
          arquivo_anexado = row["Evidencia_Anexo"]

          if arquivo_anexado:
            st.success(f"Arquivo: {arquivo_anexado}")
            caminho_arquivo = os.path.join("uploads", arquivo_anexado)

            if os.path.exists(caminho_arquivo):
              extensao = arquivo_anexado.split(".")[-1].lower()
              if extensao in ["png", "jpg", "jpeg"]:
                st.image(
                    caminho_arquivo,
                    caption="Foto enviada pelo entregador",
                    use_container_width=True,
                )
              elif extensao in ["mp4", "mov"]:
                st.video(caminho_arquivo)

              with open(caminho_arquivo, "rb") as file_to_download:
                st.download_button(
                    label="📥 Baixar Evidência",
                    data=file_to_download,
                    file_name=arquivo_anexado,
                    key=f"dl_{row['ID_Devolucao']}",
                )
            else:
              st.warning(
                  "⚠️ O registro indica um anexo, mas o arquivo físico não foi"
                  " encontrado na pasta do sistema."
              )
          else:
            st.info(
                "Nenhum arquivo de foto ou vídeo foi anexado pela transportadora"
                " neste pedido."
            )
  else:
    st.info(
        "Nenhum pedido finalizado pela transportadora aguardando análise no"
        " momento."
    )

# ==========================================
# ABA 4: INDICADORES & ALERTAS
# ==========================================
with aba_relatorios:
  st.subheader("📊 Dashboard de Performance Logística (KPIs)")
  df_geral = st.session_state.df_reversas

  if not df_geral.empty:
    total_geral = len(df_geral)
    total_aguardando = len(
        df_geral[df_geral["Status"] == "🟡 Aguardando Verificação"]
    )
    total_coleta_iniciada = len(
        df_geral[df_geral["Status"] == "🚚 Coleta Iniciada"]
    )
    total_finalizados = len(
        df_geral[df_geral["Status"] == "✅ Finalizado pela Transportadora"]
    )

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.metric("Total de Ordens", total_geral)
    col_m2.metric("Aguardando Verificação", total_aguardando)
    col_m3.metric("Coleta Iniciada", total_coleta_iniciada)
    col_m4.metric("Finalizados (Para Análise)", total_finalizados)
  else:
    st.info("Insira dados para visualizar os indicadores do painel.")

# ==========================================
# ABA 5: EMISSÃO DE TERMOS / CHECKLIST (FORMATO VISUAL E SIMPLIFICADO)
# ==========================================
with aba_termo:
  st.subheader("📄 Emissão de Checklist de Conferência (Word)")
  st.markdown(
      "Gere o termo em formato de **checklist visual** para facilitar a"
      " conferência física dos produtos."
  )

  df_termo_geral = st.session_state.df_reversas

  if not df_termo_geral.empty:
    lista_ids = df_termo_geral["ID_Devolucao"].tolist()
    id_escolhido = st.selectbox(
        "Selecione o ID da Ordem para Gerar o Checklist:",
        lista_ids,
        key="sel_id_chk",
    )

    dados_linha = df_termo_geral[
        df_termo_geral["ID_Devolucao"] == id_escolhido
    ].iloc[0]

    st.markdown("---")
    st.markdown("#### 🔍 Prévia dos Dados do Checklist:")
    c_t1, c_t2, c_t3 = st.columns(3)
    c_t1.text_input("ID", value=dados_linha["ID_Devolucao"], disabled=True)
    c_t2.text_input("Pedido", value=dados_linha["Pedido"], disabled=True)
    c_t3.text_input("Nota Fiscal", value=dados_linha["NF"], disabled=True)

    st.text_input("Cliente", value=dados_linha["Cliente"], disabled=True)
    st.text_input("Motivo", value=dados_linha["Motivo"], disabled=True)
    st.text_area("Itens", value=dados_linha["Itens"], disabled=True)

    if st.button(
        "📥 Baixar Checklist em Formato Word (.docx)",
        use_container_width=True,
        key="btn_dl_chk",
    ):
      doc = Document()

      # Margens
      for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

      COR_PRINCIPAL = RGBColor(26, 82, 118)
      COR_CINZA = RGBColor(100, 100, 100)

      # Título Principal
      p_titulo = doc.add_paragraph()
      p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
      r_titulo = p_titulo.add_run(
          "CHECKLIST DE CONFERÊNCIA E RECEBIMENTO\nLOGÍSTICA REVERSA - RMC"
          " MARIANO"
      )
      r_titulo.bold = True
      r_titulo.font.size = Pt(15)
      r_titulo.font.color.rgb = COR_PRINCIPAL

      doc.add_paragraph()

      # Bloco 1: Tabela de Identificação
      tabela_id = doc.add_table(rows=2, cols=2)
      tabela_id.alignment = WD_TABLE_ALIGNMENT.CENTER

      dados_id_box = [
          [
              f"ID da Ordem: {dados_linha['ID_Devolucao']}",
              f"Data de Registro: {dados_linha['Data_Registro']}",
          ],
          [
              f"Transportadora: {dados_linha['Transportadora']}",
              f"Motorista / Entregador:"
              f" {dados_linha['Entregador_Responsavel'] or 'Não informado'}",
          ],
      ]

      for row_idx, row_data in enumerate(dados_id_box):
        for col_idx, text in enumerate(row_data):
          cell = tabela_id.cell(row_idx, col_idx)
          cell.text = text
          shading_elm = parse_xml(
              r'<w:shd {} w:fill="F2F4F4"/>'.format(nsdecls("w"))
          )
          cell._tc.get_or_add_tcPr().append(shading_elm)

      doc.add_paragraph()

      # Bloco 2: Checklist Operacional Visual
      p_sec1 = doc.add_paragraph()
      r_sec1 = p_sec1.add_run(
          "1. ITENS DE VERIFICAÇÃO VISUAL (Tique o que foi conferido)"
      )
      r_sec1.bold = True
      r_sec1.font.size = Pt(11)
      r_sec1.font.color.rgb = COR_PRINCIPAL

      itens_checklist = [
          "A embalagem externa está lacrada e sem sinais de violação?",
          "Os produtos estão sem vazamentos, amassados ou avarias físicas?",
          "A quantidade de volumes confere com o documento de coleta?",
          (
              "As fotos/vídeos de evidência foram verificados no painel do"
              " sistema?"
          ),
      ]

      tabela_chk = doc.add_table(rows=len(itens_checklist) + 1, cols=2)
      tabela_chk.alignment = WD_TABLE_ALIGNMENT.CENTER

      hdr_cells = tabela_chk.rows[0].cells
      hdr_cells[0].text = "Pergunta / Verificação"
      hdr_cells[1].text = "STATUS ([ ] SIM  /  [ ] NÃO)"
      hdr_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

      for cell in hdr_cells:
        shd = parse_xml(r'<w:shd {} w:fill="1A5276"/>'.format(nsdecls("w")))
        cell._tc.get_or_add_tcPr().append(shd)
        for paragraph in cell.paragraphs:
          for run in paragraph.runs:
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.bold = True

      for i, item in enumerate(itens_checklist):
        row_cells = tabela_chk.rows[i + 1].cells
        row_cells[0].text = item
        row_cells[1].text = "[   ] SIM      [   ] NÃO"
        row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

      doc.add_paragraph()

      # Bloco 3: Resumo por Marcas
      p_sec2 = doc.add_paragraph()
      r_sec2 = p_sec2.add_run("2. CONTAGEM DE VOLUMES POR MARCA")
      r_sec2.bold = True
      r_sec2.font.size = Pt(11)
      r_sec2.font.color.rgb = COR_PRINCIPAL

      marcas = ["O Boticário", "Eudora", "Quem Disse, Berenice?", "O.U.i."]
      tabela_marcas = doc.add_table(rows=len(marcas) + 1, cols=3)
      tabela_marcas.alignment = WD_TABLE_ALIGNMENT.CENTER

      m_hdr = tabela_marcas.rows[0].cells
      m_hdr[0].text = "Marca / Linha"
      m_hdr[1].text = "Qtd Informada"
      m_hdr[2].text = "Qtd Conferida (CD)"

      for cell in m_hdr:
        shd = parse_xml(r'<w:shd {} w:fill="1A5276"/>'.format(nsdecls("w")))
        cell._tc.get_or_add_tcPr().append(shd)
        for paragraph in cell.paragraphs:
          for run in paragraph.runs:
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.bold = True

      for idx, marca in enumerate(marcas):
        r_cells = tabela_marcas.rows[idx + 1].cells
        r_cells[0].text = marca
        r_cells[1].text = "_______"
        r_cells[2].text = "_______"
        r_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

      doc.add_paragraph()

      # Bloco 4: Assinaturas
      p_sec3 = doc.add_paragraph()
      r_sec3 = p_sec3.add_run("3. VALIDAÇÃO E ASSINATURAS")
      r_sec3.bold = True
      r_sec3.font.size = Pt(11)
      r_sec3.font.color.rgb = COR_PRINCIPAL

      p_obs = doc.add_paragraph()
      p_obs.add_run(
          f"Cliente / Revendedora: {dados_linha['Cliente']} | Motivos:"
          f" {dados_linha['Motivo']}\n\n"
      )
      p_obs.runs[0].font.size = Pt(9.5)
      p_obs.runs[0].font.color.rgb = COR_CINZA

      tabela_ass = doc.add_table(rows=2, cols=2)
      tabela_ass.alignment = WD_TABLE_ALIGNMENT.CENTER
      tabela_ass.cell(0, 0).text = (
          "____________________________________\nResponsável pela Coleta /"
          " Entregador"
      )
      tabela_ass.cell(0, 1).text = (
          "____________________________________\nConferente / Operador do CD"
      )
      tabela_ass.cell(1, 0).text = "Data: ____/____/________"
      tabela_ass.cell(1, 1).text = "Data: ____/____/________"

      buffer = io.BytesIO()
      doc.save(buffer)
      buffer.seek(0)

      st.download_button(
          label="💾 Clique aqui para baixar o Checklist em Word",
          data=buffer,
          file_name=f"Checklist_Logistica_{dados_linha['ID_Devolucao']}.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
      st.success("Checklist gerado com sucesso!")
  else:
    st.info("Insira ordens nas abas anteriores para gerar documentos.")
