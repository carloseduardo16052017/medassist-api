"""
MedAssist AUTO — API FastAPI
Endpoints:
  POST /gerar-acs                       → Gera minuta ACS (.docx)
  POST /gerar-dbe-ingressantes          → Gera planilha DBE dos ingressantes (.xlsx)
  POST /gerar-dbe-retirantes            → Gera planilha DBE dos retirantes (.xlsx)
  POST /gerar-declaracao-autenticidade  → Gera Declaração de Autenticidade (.docx)
  POST /converter-para-pdf              → Converte .docx para .pdf via LibreOffice
"""
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
import json
import io
import pandas as pd
from docx import Document
from scripts.build_acs import gerar_acs, sem_acento, ler_ingressantes_excel

app = FastAPI(
    title="MedAssist AUTO API",
    description="API para geração automática de minutas ACS e planilhas DBE",
    version="2.1.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Helpers ────────────────────────────────────────────────────────

def gerar_planilha_dbe(dados: list[dict]) -> bytes:
    """
    Gera planilha Excel no formato DBE.
    Colunas: CPF | Nome | Quota | CEP | Número | (vazio) | Complemento | 11 | 11111111
    """
    rows = []
    for d in dados:
        cpf_limpo = ''.join(c for c in str(d.get('cpf', '')) if c.isdigit())
        cep_limpo = ''.join(c for c in str(d.get('cep', '')) if c.isdigit())
        rows.append([
            f"'{cpf_limpo}",           # Col 1: CPF sem pontuação com apóstrofo
            d.get('nome', '').upper(), # Col 2: Nome
            '1,00',                    # Col 3: Quota
            cep_limpo,                 # Col 4: CEP sem hífen
            str(d.get('numero', '')),  # Col 5: Número
            '',                        # Col 6: vazio
            str(d.get('complemento', '') or ''),  # Col 7: Complemento
            '11',                      # Col 8: fixo
            '11111111',                # Col 9: fixo
        ])

    df = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, header=False)
    output.seek(0)
    return output.read()


def extrair_nomes_retirantes_do_docx(docx_bytes: bytes) -> list[str]:
    """
    Extrai AUTOMATICAMENTE os nomes dos retirantes da seção 1.2.

    Suporta dois formatos:
    1. TABELA após "Os sócios:" — formato das minutas registradas na QUALIFEMME
       (cada linha da tabela = um retirante)
    2. PARÁGRAFOS bold numerados (1.2.1., 1.2.2...) — formato gerado pelo sistema
       e usado na OPTIMUM
    """
    from docx.oxml.ns import qn as _qn

    doc = Document(io.BytesIO(docx_bytes))
    nomes = []

    QUALIFICACAO_KWS = (
        'médico', 'médica', 'cpf', 'crm', 'rg nº', 'cep',
        'nascido', 'nascida', 'inscrito', 'inscrita', 'portador',
        'portadora', 'residente', 'domiciliado', 'domiciliada',
        'empresária', 'empresário', 'brasileiro', 'brasileira',
    )
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    body = doc.element.body
    found_os_socios = False
    para_idx = 0  # índice paralelo para acessar doc.paragraphs

    for elem in body:
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

        # ── Elemento parágrafo ────────────────────────────────────────
        if tag == 'p':
            if para_idx >= len(doc.paragraphs):
                para_idx += 1
                continue
            p = doc.paragraphs[para_idx]
            txt = p.text.strip()
            para_idx += 1

            if not found_os_socios:
                # Localizar "Os sócios:" na seção Capital Social
                if txt in ('Os sócios:', 'Os sócios') or txt.endswith('Os sócios:'):
                    found_os_socios = True
                continue

            # Fim da seção de retirantes
            if txt.startswith('Acima qualificados') or txt.startswith('Os sócios ingressantes'):
                break

            if not txt:
                continue

            # Descartar qualificações e cessões
            txt_lower = txt.lower()
            if any(kw in txt_lower for kw in QUALIFICACAO_KWS):
                continue
            if ',' in txt:
                continue

            # Formato 2: parágrafo bold = nome de retirante (OPTIMUM e ACS geradas)
            is_bold = any(r.font.bold for r in p.runs if r.font.bold is True)
            if is_bold and len(txt) > 5:
                nomes.append(txt)

        # ── Elemento tabela ───────────────────────────────────────────
        elif tag == 'tbl' and found_os_socios:
            # Formato 1: tabela de retirantes (QUALIFEMME minutas registradas)
            # Cada linha da tabela contém o nome de um retirante
            for tr in elem.findall(f'{{{NS}}}tr'):
                # Concatenar texto de todas as células da linha
                cells_text = []
                for t_elem in tr.iter(f'{{{NS}}}t'):
                    if t_elem.text:
                        cells_text.append(t_elem.text)
                nome = ' '.join(cells_text).strip()
                # Descartar cabeçalhos e linhas com qualificação
                nome_lower = nome.lower()
                if not nome or len(nome) < 5:
                    continue
                if any(kw in nome_lower for kw in QUALIFICACAO_KWS):
                    continue
                if ',' in nome:
                    continue
                nomes.append(nome)
            # Após processar a tabela de retirantes, parar
            break

    return nomes


def extrair_dados_retirantes_docx(docx_bytes: bytes) -> list[dict]:
    """
    1. Extrai os nomes dos retirantes automaticamente da seção 1.2.
    2. Busca a qualificação completa de cada retirante no preâmbulo da minuta.
    3. Extrai CPF, CEP, número e complemento de cada qualificação.
    """
    import re
    doc = Document(io.BytesIO(docx_bytes))

    # Passo 1: obter nomes dos retirantes
    nomes_retirantes = extrair_nomes_retirantes_do_docx(docx_bytes)
    if not nomes_retirantes:
        return []

    retirantes_norm = {sem_acento(n): n for n in nomes_retirantes}
    resultado = []

    # Passo 2: buscar qualificação no preâmbulo
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt or ',' not in txt:
            continue
        nome_p = sem_acento(txt.split(',')[0].strip())
        if nome_p not in retirantes_norm:
            continue

        # Extrair campos
        cpf_match = re.search(r'CPF\D{0,10}([\d]{3}\.[\d]{3}\.[\d]{3}-[\d]{2})', txt)
        cpf = cpf_match.group(1) if cpf_match else ''

        cep_match = re.search(r'CEP\s*([\d]{5}-?[\d]{3})', txt)
        cep = cep_match.group(1).replace('-', '') if cep_match else ''

        num_match = re.search(r'n[°º]\s*(\d+)', txt)
        numero = num_match.group(1) if num_match else ''

        comp_match = re.search(r'n[°º]\s*\d+\s*,\s*([^,]+),\s*\w', txt)
        complemento = comp_match.group(1).strip() if comp_match else ''

        resultado.append({
            'nome': retirantes_norm[nome_p],
            'cpf': cpf,
            'cep': cep,
            'numero': numero,
            'complemento': complemento,
        })

    return resultado


# ── Endpoints ──────────────────────────────────────────────────────

@app.get("/")
def root():
    return {"status": "ok", "service": "MedAssist AUTO API", "version": "1.1.0"}


@app.get("/health")
def health():
    return {"status": "healthy"}


@app.post("/gerar-acs")
async def gerar_acs_endpoint(
    template: UploadFile = File(..., description="Última ACS registrada (.docx)"),
    relatorio: UploadFile = File(..., description="Relatório ClickSign (.xlsx)"),
    retirantes: str = Form(..., description='JSON array com nomes dos retirantes'),
    num_alteracao: int = Form(..., description="Número da nova alteração"),
    config: str = Form(..., description="JSON com configurações da empresa"),
):
    try:
        retirantes_list = json.loads(retirantes)
        cfg = json.loads(config)
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=400, detail=f"JSON inválido: {e}")

    template_bytes = await template.read()
    excel_bytes = await relatorio.read()
    manter_set = {sem_acento(n) for n in cfg.get('manter_preambulo', [])}

    try:
        docx_bytes = gerar_acs(
            template_bytes=template_bytes,
            excel_bytes=excel_bytes,
            retirantes=retirantes_list,
            num_alteracao=num_alteracao,
            manter_preambulo=manter_set,
            andres_texto=cfg.get('andres_texto'),
            lidya_texto=cfg.get('lidya_texto'),
            auro_q_inicio=cfg.get('majoritario_q_inicio'),
            nomes_qsa_apos=cfg.get('nomes_qsa_apos', []),
            majoritario_nome=cfg.get('majoritario_nome', 'AURO BUFFANI CLAUDINO'),
            majoritario_q_final=cfg.get('majoritario_q_final'),
            secundario_nome=cfg.get('secundario_nome'),
            secundario_q=cfg.get('secundario_q'),
            total_capital=cfg.get('total_capital', 10000),
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao gerar ACS: {str(e)}")

    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{num_alteracao}a_ACS_v1.docx"'}
    )


@app.post("/gerar-dbe-ingressantes")
async def gerar_dbe_ingressantes(
    relatorio: UploadFile = File(..., description="Relatório ClickSign (.xlsx) com dados dos ingressantes"),
):
    """
    Gera planilha DBE dos sócios INGRESSANTES.
    Lê os dados do relatório ClickSign e formata para o DBE.
    """
    try:
        excel_bytes = await relatorio.read()
        dados = ler_ingressantes_excel(excel_bytes)
        if not dados:
            raise HTTPException(status_code=400, detail="Nenhum ingressante encontrado no relatório")
        xlsx_bytes = gerar_planilha_dbe(dados)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao gerar DBE ingressantes: {str(e)}")

    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="DBE_Ingressantes.xlsx"'}
    )


@app.post("/gerar-dbe-retirantes")
async def gerar_dbe_retirantes(
    minuta: UploadFile = File(..., description="Última ACS registrada (.docx)"),
):
    """
    Gera planilha DBE dos sócios RETIRANTES.
    Formato: 2 colunas — CPF (sem pontuação) | Nome (maiúsculas).
    Aba: 'Sócios Ingressantes'.
    Extrai automaticamente da minuta sem precisar informar os nomes.
    """
    import re
    try:
        docx_bytes = await minuta.read()
        dados = extrair_dados_retirantes_docx(docx_bytes)
        if not dados:
            raise HTTPException(
                status_code=404,
                detail="Nenhum retirante encontrado na seção 'Os sócios:' da minuta."
            )

        # Montar planilha: apenas CPF limpo + Nome maiúsculas
        rows = []
        for d in dados:
            cpf_limpo = re.sub(r'[.\-]', '', str(d.get('cpf', '')))
            nome = d.get('nome', '').upper().strip()
            rows.append([cpf_limpo, nome])

        df = pd.DataFrame(rows, columns=['CPF', 'Nome'])
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False, sheet_name='Sócios Ingressantes')
        output.seek(0)
        xlsx_bytes = output.read()

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao gerar DBE retirantes: {str(e)}")

    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="DBE_Retirantes.xlsx"'}
    )


# ── Helpers — Declaração de Autenticidade ──────────────────────────

def extrair_representante_da_minuta(docx_bytes: bytes) -> str:
    """
    Extrai o nome do representante da empresa da minuta registrada.
    Busca pelos padrões:
      - "representadas por procuração por [NOME]"
      - "representados por procuração por [NOME]"
    """
    import re
    doc = Document(io.BytesIO(docx_bytes))
    for p in doc.paragraphs:
        txt = p.text.strip()
        m = re.search(
            r'representad[ao]s?\s+por\s+procura[çc][aã]o\s+por\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÈÌÒÙÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÈÌÒÙÇa-záéíóúâêîôûãõàèìòùç\s]+?)(?:,|\.|$)',
            txt, re.IGNORECASE
        )
        if m:
            nome = m.group(1).strip()
            # Capitalizar corretamente
            return nome.title()
    return "o Representante"


def extrair_nome_empresa_da_minuta(docx_bytes: bytes) -> str:
    """
    Extrai o nome da empresa da minuta registrada.
    O nome da empresa fica na 3ª linha do cabeçalho (parágrafo 2).
    """
    doc = Document(io.BytesIO(docx_bytes))
    # O nome da empresa está tipicamente no parágrafo 2 (índice 2)
    for p in doc.paragraphs[:6]:
        txt = p.text.strip()
        # Nome da empresa tem LTDA ou S/A e é todo maiúsculo
        if txt and ('LTDA' in txt or 'S/A' in txt or 'EIRELI' in txt):
            return txt
    return ""


def _make_run_elem(text: str, bold: bool = False) -> object:
    """Cria um elemento <w:r> com Calibri (minorHAnsi) 9pt e bold opcional."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:asciiTheme'), 'minorHAnsi')
    rFonts.set(qn('w:hAnsiTheme'), 'minorHAnsi')
    sz   = OxmlElement('w:sz');   sz.set(qn('w:val'), '18')    # 9pt = 18 half-points
    szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), '18')
    if bold:
        b = OxmlElement('w:b'); rPr.append(b)
    rPr.append(rFonts); rPr.append(sz); rPr.append(szCs)
    r.append(rPr)
    t = OxmlElement('w:t'); t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    return r


def _replace_text_in_para(para, old: str, new: str):
    """Substitui texto simples mantendo todos os runs inalterados."""
    full = ''.join(r.text for r in para.runs)
    if old not in full:
        return False
    new_full = full.replace(old, new, 1)
    if para.runs:
        para.runs[0].text = new_full
        for r in para.runs[1:]:
            r.text = ''
    return True


def _replace_with_bold_segment(para, full_new_text: str, bold_segment: str):
    """
    Reescreve um parágrafo com Calibri 9pt, deixando bold_segment em negrito.
    Remove todos os runs existentes e cria novos com a formatação correta.
    """
    from docx.oxml.ns import qn
    # Remover runs existentes
    for r in list(para._p.findall(qn('w:r'))):
        para._p.remove(r)
    if bold_segment and bold_segment in full_new_text:
        idx   = full_new_text.index(bold_segment)
        before = full_new_text[:idx]
        after  = full_new_text[idx + len(bold_segment):]
        if before: para._p.append(_make_run_elem(before, bold=False))
        para._p.append(_make_run_elem(bold_segment, bold=True))
        if after:  para._p.append(_make_run_elem(after,  bold=False))
    else:
        para._p.append(_make_run_elem(full_new_text, bold=False))


def gerar_declaracao_autenticidade_docx(
    nomes_ingressantes: list[str],
    nome_representante: str,
    nome_empresa: str,
) -> bytes:
    """
    Abre o template declaracao_autenticidade_template.docx e substitui:
    - Nome da empresa no item 2 da lista
    - Nome do representante no item 3
    - Linhas da tabela (mantém cabeçalho, substitui ingressantes com nome em negrito)
    - Data pela data de hoje em português
    """
    import os
    import copy
    from datetime import date
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    MONTHS_PT = {
        1:'janeiro', 2:'fevereiro', 3:'março', 4:'abril',
        5:'maio', 6:'junho', 7:'julho', 8:'agosto',
        9:'setembro', 10:'outubro', 11:'novembro', 12:'dezembro'
    }
    today = date.today()
    data_pt = f"{today.day:02d} de {MONTHS_PT[today.month]} de {today.year}"

    # Localizar o template relativo ao main.py
    base_dir = os.path.dirname(os.path.abspath(__file__))
    tpl_path = os.path.join(base_dir, 'templates', 'declaracao_autenticidade_template.docx')
    doc = Document(tpl_path)

    # ── 1. Item 2 — nome da empresa em negrito ───────────────────────
    empresa_nova = nome_empresa or 'QUALIFEMME SERVIÇOS MÉDICOS LTDA'
    for p in doc.paragraphs:
        txt = ''.join(r.text for r in p.runs)
        if 'QUALIFEMME SERVIÇOS MÉDICOS LTDA' in txt and 'Procuração outorgada' in txt:
            novo_txt = txt.replace('QUALIFEMME SERVIÇOS MÉDICOS LTDA', empresa_nova)
            _replace_with_bold_segment(p, novo_txt, empresa_nova)
            break

    # ── 2. Item 3 — nome do representante em negrito ─────────────────
    for p in doc.paragraphs:
        txt = ''.join(r.text for r in p.runs)
        if 'Marcelo Costa Moreira' in txt and 'outorgada para' in txt:
            novo_txt = txt.replace('Marcelo Costa Moreira', nome_representante)
            _replace_with_bold_segment(p, novo_txt, nome_representante)
            break

    # ── 3. Substituir data ───────────────────────────────────────────
    for p in doc.paragraphs:
        txt = ''.join(r.text for r in p.runs)
        if 'São Paulo, SP,' in txt and 'de 20' in txt:
            import re
            old_date = re.search(r'\d{1,2} de \w+ de \d{4}', txt)
            if old_date:
                _replace_text_in_para(p, old_date.group(0), data_pt)
            break

    # ── 4. Substituir linhas da tabela ───────────────────────────────
    for tbl in doc.tables:
        if len(tbl.rows) < 1:
            continue
        # Verificar se é a tabela de ingressantes (cabeçalho com NOME)
        hdr_txt = tbl.rows[0].cells[0].text.strip()
        if hdr_txt.upper() != 'NOME':
            continue

        # Remover todas as linhas exceto o cabeçalho
        for row in list(tbl.rows[1:]):
            tbl._tbl.remove(row._tr)

        # Adicionar linhas dos ingressantes com nome em negrito
        template_row = tbl.rows[0]  # usar cabeçalho como referência de formato
        for nome in nomes_ingressantes:
            # Criar nova linha copiando estrutura do cabeçalho
            new_tr = copy.deepcopy(template_row._tr)
            cells = new_tr.findall(f'{{{qn("w:tc").split("}")[0][1:]}}}tc') if False else \
                    new_tr.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')

            # Limpar e definir conteúdo das células — Calibri 9pt explícito
            def set_tc_text(tc_elem, text, bold=False):
                NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                for p_elem in tc_elem.findall(f'{{{NS}}}p'):
                    for r_elem in list(p_elem.findall(f'{{{NS}}}r')):
                        p_elem.remove(r_elem)
                    r_new = _make_run_elem(text, bold=bold)
                    p_elem.append(r_new)

            if len(cells) >= 2:
                set_tc_text(cells[0], nome.upper(), bold=True)
                set_tc_text(cells[1], 'PROCURAÇÃO E DOCUMENTOS PESSOAIS', bold=False)

            tbl._tbl.append(new_tr)
        break  # processar apenas a primeira tabela relevante

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.read()


@app.post("/gerar-declaracao-autenticidade")
async def gerar_declaracao_autenticidade(
    relatorio_clicksign: UploadFile = File(..., description="Relatório ClickSign (.xlsx)"),
    minuta_registrada: UploadFile = File(..., description="Minuta registrada (.docx)"),
):
    """
    Gera a Declaração de Autenticidade (.docx).
    Extrai ingressantes do relatório ClickSign e representante da minuta.
    """
    try:
        excel_bytes = await relatorio_clicksign.read()
        docx_bytes  = await minuta_registrada.read()

        # Extrair ingressantes
        dados_ing = ler_ingressantes_excel(excel_bytes)
        if not dados_ing:
            raise HTTPException(status_code=400, detail="Nenhum ingressante encontrado no relatório ClickSign.")
        nomes_ing = [d['nome'].upper() for d in dados_ing]

        # Extrair representante e nome da empresa da minuta
        nome_rep     = extrair_representante_da_minuta(docx_bytes)
        nome_empresa = extrair_nome_empresa_da_minuta(docx_bytes)

        # Gerar .docx
        docx_out = gerar_declaracao_autenticidade_docx(nomes_ing, nome_rep, nome_empresa)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao gerar declaração: {str(e)}")

    return Response(
        content=docx_out,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="Declaracao_Autenticidade.docx"'}
    )


def _extrair_dados_declaracao(docx_bytes: bytes) -> dict:
    """
    Extrai do .docx da Declaração de Autenticidade:
    - nome_empresa (do item 2)
    - nome_representante (do item 3)
    - data (parágrafo com 'São Paulo, SP,')
    - ingressantes (nomes da tabela, excluindo cabeçalho)
    """
    doc = Document(io.BytesIO(docx_bytes))
    dados = {
        "nome_empresa": "",
        "nome_representante": "",
        "data": "",
        "ingressantes": [],
    }

    for p in doc.paragraphs:
        txt = p.text.strip()
        if 'Procuração outorgada' in txt and 'para Raphael' in txt:
            # "Procuração outorgada EMPRESA para Raphael..."
            import re as _re
            m = _re.search(r'outorgada\s+(.+?)\s+para Raphael', txt)
            if m:
                dados["nome_empresa"] = m.group(1).strip()
        elif 'outorgada para' in txt and 'sócios ingressantes' in txt:
            # "...outorgada para REPRESENTANTE"
            import re as _re
            m = _re.search(r'outorgada para\s+(.+)$', txt)
            if m:
                dados["nome_representante"] = m.group(1).strip()
        elif 'São Paulo, SP,' in txt:
            dados["data"] = txt

    for tbl in doc.tables:
        if tbl.rows and tbl.rows[0].cells[0].text.strip().upper() == 'NOME':
            for row in tbl.rows[1:]:
                nome = row.cells[0].text.strip()
                if nome:
                    dados["ingressantes"].append(nome)

    return dados


@app.post("/converter-para-pdf")
async def converter_para_pdf(
    arquivo: UploadFile = File(..., description="Arquivo .docx da Declaração de Autenticidade"),
):
    """
    Gera PDF da Declaração de Autenticidade usando ReportLab.
    Extrai dados do .docx recebido e monta o PDF com a formatação correta.
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    )
    from reportlab.lib import colors
    from reportlab.lib.colors import black

    docx_bytes = await arquivo.read()
    nome_base  = arquivo.filename.replace(".docx", "") if arquivo.filename else "Declaracao"

    # ── Extrair dados do .docx ──────────────────────────────────────
    dados = _extrair_dados_declaracao(docx_bytes)
    empresa     = dados["nome_empresa"]    or "QUALIFEMME SERVIÇOS MÉDICOS LTDA"
    representante = dados["nome_representante"] or "Marcelo Costa Moreira"
    data_txt    = dados["data"]            or "São Paulo, SP, ___ de ___ de 2026."
    ingressantes = dados["ingressantes"]

    # ── Estilos base ────────────────────────────────────────────────
    FONT      = "Helvetica"
    FONT_BOLD = "Helvetica-Bold"
    SZ        = 9

    def st(name, **kw):
        base = dict(fontName=FONT, fontSize=SZ, leading=12, spaceAfter=4)
        base.update(kw)
        return ParagraphStyle(name, **base)

    s_justify  = st("justify",  alignment=TA_JUSTIFY)
    s_center   = st("center",   alignment=TA_CENTER)
    s_left     = st("left",     alignment=TA_LEFT)

    # ── Construir conteúdo ──────────────────────────────────────────
    story = []
    W, H  = A4

    # Título em caixa com borda
    titulo_style = ParagraphStyle(
        "titulo", fontName=FONT_BOLD, fontSize=SZ, alignment=TA_CENTER, leading=14
    )
    titulo_tbl = Table(
        [[Paragraph("DECLARAÇÃO DE AUTENTICIDADE", titulo_style)]],
        colWidths=[W - 4*cm],
    )
    titulo_tbl.setStyle(TableStyle([
        ("BOX",         (0,0), (-1,-1), 1, black),
        ("TOPPADDING",  (0,0), (-1,-1), 6),
        ("BOTTOMPADDING",(0,0),(-1,-1), 6),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("RIGHTPADDING",(0,0), (-1,-1), 8),
    ]))
    story.append(titulo_tbl)
    story.append(Spacer(1, 0.4*cm))

    # Corpo — RAPHAEL e DECLARO em negrito
    corpo = (
        f'Eu <b>RAPHAEL ALVES ANTUNES</b>, com inscrição ativa na OAB/SP sob o nº 286.717, '
        f'expedida em 24.10.2019, inscrito no CPF nº 340.541.598-50, <b>DECLARO</b>, '
        f'sob as penas da Lei penal e, sem prejuízo das sanções administrativas e cíveis, '
        f'que estes documentos são autênticos e condizem com os originais respectivos.'
    )
    story.append(Paragraph(corpo, s_justify))
    story.append(Spacer(1, 0.3*cm))

    # "Documentos apresentados:" sublinhado
    story.append(Paragraph('<u>Documentos apresentados:</u>', s_left))
    story.append(Spacer(1, 0.2*cm))

    # Item 1 — "Carteira OAB/SP" sublinhado
    story.append(Paragraph(
        '- <u>Carteira OAB/SP</u> de Raphael Alves Antunes (Qtde. Folhas: 1);',
        s_left
    ))
    # Item 2 — empresa em negrito
    story.append(Paragraph(
        f'- Procuração outorgada <b>{empresa}</b> para Raphael Alves Antunes (Qtde. Folhas: 2)',
        s_left
    ))
    # Item 3 — representante em negrito
    story.append(Paragraph(
        f'- Procurações e documentos dos sócios ingressantes, outorgada para <b>{representante}</b>',
        s_left
    ))
    story.append(Spacer(1, 0.4*cm))

    # Tabela de ingressantes
    hdr_style = ParagraphStyle("hdr", fontName=FONT, fontSize=SZ, alignment=TA_CENTER, leading=12)
    row_name  = ParagraphStyle("rn",  fontName=FONT_BOLD, fontSize=SZ, alignment=TA_CENTER, leading=12)
    row_proc  = ParagraphStyle("rp",  fontName=FONT, fontSize=SZ, alignment=TA_CENTER, leading=12)

    tbl_data = [[
        Paragraph("NOME", hdr_style),
        Paragraph("PROCURAÇÃO E DOCUMENTOS PESSOAIS", hdr_style),
    ]]
    for nome in ingressantes:
        tbl_data.append([
            Paragraph(nome, row_name),
            Paragraph("PROCURAÇÃO E DOCUMENTOS PESSOAIS", row_proc),
        ])

    col1_w = (W - 4*cm) * 0.55
    col2_w = (W - 4*cm) * 0.45
    tabela = Table(tbl_data, colWidths=[col1_w, col2_w], repeatRows=1)
    tabela.setStyle(TableStyle([
        ("GRID",         (0,0), (-1,-1), 0.5, colors.grey),
        ("BACKGROUND",   (0,0), (-1,0),  colors.whitesmoke),
        ("TOPPADDING",   (0,0), (-1,-1), 4),
        ("BOTTOMPADDING",(0,0), (-1,-1), 4),
        ("LEFTPADDING",  (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(tabela)
    story.append(Spacer(1, 0.5*cm))

    # Data
    story.append(Paragraph(data_txt, s_center))
    story.append(Spacer(1, 0.8*cm))

    # Assinatura em negrito
    story.append(Paragraph(
        '<b>RAPHAEL ALVES ANTUNES</b>',
        ParagraphStyle("sig", fontName=FONT_BOLD, fontSize=SZ, alignment=TA_CENTER, leading=12)
    ))

    # ── Gerar PDF ───────────────────────────────────────────────────
    buf = io.BytesIO()
    doc_pdf = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2.5*cm, rightMargin=2.5*cm,
        topMargin=2.5*cm,  bottomMargin=2.5*cm,
    )
    doc_pdf.build(story)
    buf.seek(0)
    pdf_bytes = buf.read()

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{nome_base}.pdf"'}
    )



# ── /gerar-vre ──────────────────────────────────────────────
import re as _re, shutil as _shutil, tempfile as _tempfile, os as _os
from openpyxl import load_workbook as _load_workbook

def _parse_ec(raw):
    if not raw or str(raw).strip() in ('', 'nan'): return 'Outros', None
    r = str(raw).lower()
    if 'casado' in r:
        m = _re.search(r'sob o regime d[ao]? (.+)', r)
        reg = m.group(1).strip().capitalize() if m else None
        if reg and 'separa' in reg.lower(): reg = 'Separação de bens'
        return 'Casado(a)', reg
    if 'solteiro' in r: return 'Solteiro(a)', None
    if 'divorciado' in r: return 'Divorciado(a)', None
    if 'viú' in r or 'viuv' in r: return 'Viuvo(a)', None
    if 'desquitado' in r: return 'Desquitado(a)', None
    if 'separado' in r and 'judicial' in r: return 'Separado(a) Judicialmente', None
    return 'Outros', None

def _fc(v):
    if not v or str(v).lower()=='nan': return None
    return _re.sub(r'[.\-]','',str(v)).zfill(11)

def _fd(v):
    if not v or str(v).lower()=='nan': return None
    p = _re.split(r'[/\-.]', str(v).strip())
    if len(p)==3: return f"{p[0].zfill(2)}{p[1].zfill(2)}{p[2]}"
    c = _re.sub(r'[/\-.]','',str(v))
    return c.zfill(8) if c.isdigit() else v

def _fr(v):
    if not v or str(v).lower()=='nan': return None
    return _re.sub(r'[.\-/\s]','',str(v))

def _retirantes(docx_bytes):
        import io as _io; from docx import Document as _D; from docx.oxml.ns import qn as _qn
    doc = _D(_io.BytesIO(docx_bytes))
    retirante_names = []; found_socios = False; found_retirando = False
    for _child in doc.element.body:
        _tag = _child.tag.split('}')[-1] if '}' in _child.tag else _child.tag
        if _tag == 'p':
            _txt = ''.join(_t.text or '' for _t in _child.findall('.//' + _qn('w:t')))
            if _txt.strip().endswith('cios:') and not found_socios: found_socios = True
            elif found_socios and not found_retirando and 'retirando-se' in _txt.lower(): found_retirando = True; break
        elif _tag == 'tbl' and found_socios and not found_retirando:
            for _cell in _child.findall('.//' + _qn('w:tc')):
                _ct = ''.join(_t.text or '' for _t in _cell.findall('.//' + _qn('w:t'))).strip()
                if _ct and len(_ct) > 5: retirante_names.append(_ct)
    result = []
    for _nm in retirante_names:
        for _p in doc.paragraphs:
            if _p.text.strip().startswith(_nm):
                _m = _re.search('CPF sob n[^ ]* *([0-9]{3}[.][0-9]{3}[.][0-9]{3}-[0-9]{2})', _p.text)
                if _m: result.append({'nome': _nm, 'cpf': _re.sub('[.-]','',_m.group(1)).zfill(11)}); break
    return result

@app.post("/gerar-vre")
async def gerar_vre(relatorio: UploadFile = File(...), minuta: UploadFile = File(...)):
    tmp = _tempfile.mkdtemp()
    try:
        rp = _os.path.join(tmp,"r.xlsx"); mp = _os.path.join(tmp,"m.docx"); op = _os.path.join(tmp,"VRE.xlsx")
        with open(rp,"wb") as f: f.write(await relatorio.read())
        mb = await minuta.read()
        with open(mp,"wb") as f: f.write(mb)
        df = pd.read_excel(rp, sheet_name='MeusDados')
        df = df[df['Status do documento']=='Finalizado'].copy()
        rets = _retirantes(mb)
        from openpyxl import Workbook as _WBB; _wbb=_WBB(); _wbb.active.title='Dados'; _wbb.create_sheet('Retirantes'); _wbb.save(op)
        wb = _load_workbook(op)
        ws = wb['Dados']
        for row in ws.iter_rows():
            for cell in row: cell.value = None
        for i,(_, r) in enumerate(df.iterrows(), start=1):
            ec, reg = _parse_ec(r.get('Formulário 1 Qual o seu estado civil?'))
            nac = r.get('Formulário 1 Qual a sua nacionalidade?')
            ws.cell(i,1).value = r.get('Formulário 1 Qual o seu nome completo?')
            ws.cell(i,2).value = _fc(r.get('Formulário 1 Qual o seu CPF?'))
            ws.cell(i,3).value = _fd(r.get('Formulário 1 Qual a sua data de nascimento?'))
            ws.cell(i,4).value = _fr(r.get('Formulário 1 Qual o número do seu RG ou RNE (caso estrangeiro)?'))
            ws.cell(i,5).value = r.get('Formulário 1 Qual o órgão emissor do seu RG ou RNE (caso estrangeiro)?')
            ws.cell(i,6).value = r.get('Formulário 1 Qual o Estado (UF) de Emissão do seu RG ou RNE (caso estrangeiro)?')
            ws.cell(i,7).value = (str(nac)[0].upper()+str(nac)[1:]) if nac and str(nac).lower()!='nan' else None
            ws.cell(i,8).value = 'Médico(a)'
            ws.cell(i,9).value = ec
            ws.cell(i,10).value = reg
            ws.cell(i,11).value = 'Não Declarada'
            ws.cell(i,12).value = r.get('Formulário 1 CEP')
            ws.cell(i,13).value = r.get('Formulário 1 Número')
            ws.cell(i,14).value = r.get('Formulário 1 Complemento de endereço')
        ws_r = wb['Retirantes']
        ws_r.cell(1,1).value='CPF'; ws_r.cell(1,2).value='Nome'
        for row in ws_r.iter_rows(min_row=2):
            for cell in row: cell.value=None
        for i,ret in enumerate(rets,start=2):
            ws_r.cell(i,1).value=ret['cpf']; ws_r.cell(i,2).value=ret['nome']
        wb.save(op)
        with open(op,"rb") as f: content=f.read()
        return Response(content=content, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition":"attachment; filename=VRE_JUCESP.xlsx"})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        _shutil.rmtree(tmp, ignore_errors=True)
