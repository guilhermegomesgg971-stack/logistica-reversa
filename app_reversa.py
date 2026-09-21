import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def gerar_termo_checklist_docx(dados_pedido):
    doc = Document()
    
    # Configurar margens
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    # Cores do tema (Azul corporativo e cinza claro)
    COR_PRINCIPAL = RGBColor(26, 82, 118)
    COR_CINZA = RGBColor(100, 100, 100)
    
    # Título Principal
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_titulo = p_titulo.add_run("CHECKLIST DE CONFERÊNCIA E RECEBIMENTO\nLOGÍSTICA REVERSA")
    r_titulo.bold = True
    r_titulo.font.size = Pt(16)
    r_titulo.font.color.rgb = COR_PRINCIPAL
    
    doc.add_paragraph() # Espaço
    
    # Bloco 1: Informações Básicas (Tabela de Identificação)
    tabela_id = doc.add_table(rows=2, cols=2)
    tabela_id.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    dados_id = [
        [f"ID da Ordem: {dados_pedido.get('id', 'N/A')}", f"Data da Coleta: {dados_pedido.get('data', 'N/A')}"],
        [f"Transportadora: {dados_pedido.get('transportadora', 'N/A')}", f"Entregador: {dados_pedido.get('entregador', 'N/A')}"]
    ]
    
    for row_idx, row_data in enumerate(dados_id):
        for col_idx, text in enumerate(row_data):
            cell = tabela_id.cell(row_idx, col_idx)
            cell.text = text
            # Estilizar fundo da caixa de ID
            shading_elm = parse_xml(r'<w:shd {} w:fill="F2F4F4"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(shading_elm)
            
    doc.add_paragraph()
    
    # Bloco 2: Checklist Operacional (O que o conferente deve ticar)
    p_sec1 = doc.add_paragraph()
    r_sec1 = p_sec1.add_run("1. CONFERÊNCIA VISUAL E DE SEGURANÇA")
    r_sec1.bold = True
    r_sec1.font.size = Pt(12)
    r_sec1.font.color.rgb = COR_PRINCIPAL
    
    itens_checklist = [
        "A embalagem externa está lacrada e sem sinais de violação?",
        "Os produtos estão sem vazamentos, amassados ou avarias físicas?",
        "A quantidade de volumes confere com a nota/ordem de coleta?",
        "As fotos/vídeos de evidência foram registrados pelo entregador?"
    ]
    
    tabela_chk = doc.add_table(rows=len(itens_checklist) + 1, cols=2)
    tabela_chk.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Cabeçalho da tabela de checklist
    hdr_cells = tabela_chk.rows[0].cells
    hdr_cells[0].text = "Item de Verificação"
    hdr_cells[1].text = "STATUS ([ ] SIM  /  [ ] NÃO)"
    hdr_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for cell in hdr_cells:
        shd = parse_xml(r'<w:shd {} w:fill="1A5276"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shd)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.bold = True

    for i, item in enumerate(itens_checklist):
        row_cells = tabela_chk.rows[i + 1].cells
        row_cells[0].text = item
        row_cells[1].text = "[   ] OK      [   ] FALHOU"
        row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # Bloco 3: Resumo de Produtos / Marcas Conferidas
    p_sec2 = doc.add_paragraph()
    r_sec2 = p_sec2.add_run("2. RESUMO DE VOLUMES POR MARCA")
    r_sec2.bold = True
    r_sec2.font.size = Pt(12)
    r_sec2.font.color.rgb = COR_PRINCIPAL
    
    marcas = ["O Boticário", "Eudora", "Quem Disse, Berenice?", "O.U.i."]
    
    tabela_marcas = doc.add_table(rows=len(marcas) + 1, cols=3)
    tabela_marcas.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    m_hdr = tabela_marcas.rows[0].cells
    m_hdr[0].text = "Marca / Linha"
    m_hdr[1].text = "Qtd Informada"
    m_hdr[2].text = "Qtd Conferida (CD)"
    
    for cell in m_hdr:
        shd = parse_xml(r'<w:shd {} w:fill="1A5276"/>'.format(nsdecls('w')))
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
    doc.add_paragraph()

    # Bloco 4: Assinaturas Simplificadas
    p_sec3 = doc.add_paragraph()
    r_sec3 = p_sec3.add_run("3. VALIDAÇÃO E ASSINATURAS")
    r_sec3.bold = True
    r_sec3.font.size = Pt(12)
    r_sec3.font.color.rgb = COR_PRINCIPAL

    p_obs = doc.add_paragraph()
    p_obs.add_run("Declaro que os produtos acima foram conferidos visualmente de acordo com os itens listados.\n\n")
    p_obs.runs[0].font.size = Pt(10)
    p_obs.runs[0].font.color.rgb = COR_CINZA

    tabela_ass = doc.add_table(rows=2, cols=2)
    tabela_ass.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    tabela_ass.cell(0, 0).text = "____________________________________\nResponsável pela Coleta / Entregador"
    tabela_ass.cell(0, 1).text = "____________________________________\nConferente / Operador do CD"
    
    tabela_ass.cell(1, 0).text = "Data: ____/____/________"
    tabela_ass.cell(1, 1).text = "Data: ____/____/________"

    # Salvar em memória para o Streamlit disponibilizar o botão de Download
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
