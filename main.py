"""
MedAssist AUTO — API FastAPI
Endpoints:
  POST /gerar-acs                       → Gera minuta ACS (.docx)
  POST /gerar-dbe-ingressantes          → Gera planilha DBE dos ingressantes (.xlsx)
  POST /gerar-dbe-retirantes            → Gera planilha DBE dos retirantes (.xlsx)
  POST /gerar-declaracao-autenticidade  → Gera Declaração de Autenticidade (.docx)
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
    version="1.9.0"
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
