"""
MedAssist AUTO — Gerador de ACS
Lógica central: usar a última ACS como template e alterar apenas o QSA.
"""
import unicodedata
import copy
import re
import io
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pandas as pd


def sem_acento(s: str) -> str:
    return unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode().upper()


def sort_key(t: str) -> str:
    return sem_acento(t.split(',')[0].strip())


def del_p(p):
    p._p.getparent().remove(p._p)


def clone_para(template_para, runs_data):
    """Clona parágrafo substituindo runs. runs_data = [(texto, bold), ...]"""
    new_p = copy.deepcopy(template_para._p)
    for r in list(new_p.findall(qn('w:r'))):
        new_p.remove(r)

    bold_rPr = None
    normal_rPr = None
    for r in template_para._p.findall(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        is_bold = rPr is not None and rPr.find(qn('w:b')) is not None
        if is_bold and bold_rPr is None:
            bold_rPr = copy.deepcopy(rPr)
        if not is_bold and normal_rPr is None:
            normal_rPr = copy.deepcopy(rPr) if rPr is not None else None

    for texto, bold in runs_data:
        if not texto:
            continue
        r = OxmlElement('w:r')
        rPr_use = copy.deepcopy(bold_rPr if bold else normal_rPr)
        if rPr_use is None:
            rPr_use = OxmlElement('w:rPr')
            if bold:
                b = OxmlElement('w:b')
                rPr_use.append(b)
        r.append(rPr_use)
        t = OxmlElement('w:t')
        t.text = texto
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t)
        new_p.append(r)
    return new_p


def ler_ingressantes_excel(excel_bytes: bytes) -> list[dict]:
    """
    Lê o relatório ClickSign (.xlsx) e retorna lista de ingressantes.
    Suporta aba 'QUALIFICAÇÕES' ou fallback para 'MeusDados'.
    """
    df = None
    try:
        df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name='QUALIFICAÇÕES')
    except Exception:
        try:
            df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name='MeusDados')
        except Exception:
            df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=0)

    col_map = {
        'nome':          ['Formulário 1 Qual o seu nome completo?', 'NOME'],
        'nascimento':    ['Formulário 1 Qual a sua data de nascimento?', 'DT NASCIMENTO'],
        'nacionalidade': ['Formulário 1 Qual a sua nacionalidade?', 'NACIONALIDADE'],
        'estado_civil':  ['Formulário 1 Qual o seu estado civil?', 'ESTADO CIVIL'],
        'crm':           ['Formulário 1 Qual o número de seu CRM?', 'CRM'],
        'crm_uf':        ['Formulário 1 Qual o Estado de Emissão do CRM?', 'UF CRM'],
        'rg':            ['Formulário 1 Qual o número do seu RG ou RNE (caso estrangeiro)?', 'RG/RNE'],
        'rg_orgao':      ['Formulário 1 Qual o órgão emissor do seu RG ou RNE (caso estrangeiro)?', 'ÓRGÃO RG/RNE'],
        'rg_uf':         ['Formulário 1 Qual o Estado (UF) de Emissão do seu RG ou RNE (caso estrangeiro)?', 'UF RG/RNE'],
        'cpf':           ['Formulário 1 Qual o seu CPF?', 'CPF'],
        'cep':           ['Formulário 1 CEP', 'CEP'],
        'logradouro':    ['Formulário 1 Logradouro', 'LOGRADOURO'],
        'numero':        ['Formulário 1 Número', 'Nº'],
        'complemento':   ['Formulário 1 Complemento de endereço', 'COMPLEMENTO'],
        'bairro':        ['Formulário 1 Bairro', 'BAIRRO'],
        'cidade':        ['Formulário 1 Cidade', 'CIDADE'],
        'estado':        ['Formulário 1 Estado', 'ESTADO'],
    }

    def get_col(row, candidates):
        for c in candidates:
            if c in df.columns and pd.notna(row.get(c)):
                return str(row[c]).strip()
        return ''

    ingressantes = []
    for _, row in df.iterrows():
        row = row.to_dict()
        d = {k: get_col(row, v) for k, v in col_map.items()}
        if not d['nome']:
            continue
        ingressantes.append(d)

    return sorted(ingressantes, key=lambda x: sem_acento(x['nome']))


def formatar_qualificacao(d: dict) -> str:
    """Monta o texto de qualificação completo a partir dos dados do ClickSign."""
    nome = d['nome'].upper()
    nac = d.get('nacionalidade', 'brasileiro(a)').lower()
    ec = d.get('estado_civil', 'solteiro(a)').lower()
    nasc = d.get('nascimento', '')
    crm = d.get('crm', '')
    crm_uf = d.get('crm_uf', 'SP')
    rg = d.get('rg', '')
    rg_orgao = d.get('rg_orgao', 'SSP')
    rg_uf = d.get('rg_uf', 'SP')
    cpf = d.get('cpf', '')
    logr = d.get('logradouro', '')
    num = d.get('numero', '')
    comp = d.get('complemento', '')
    bairro = d.get('bairro', '')
    cidade = d.get('cidade', '')
    estado = d.get('estado', 'SP')
    cep = d.get('cep', '')

    # Género baseado na nacionalidade/nome
    feminino = any(x in nac for x in ['a)', 'eira', 'essa'])

    med = 'médica' if feminino else 'médico'
    inscr = 'inscrita' if feminino else 'inscrito'
    port = 'portadora' if feminino else 'portador'
    nasc_v = 'nascida' if feminino else 'nascido'
    res = 'residente e domiciliada' if feminino else 'residente e domiciliado'

    comp_str = f', {comp}' if comp and str(comp).lower() not in ('nan', '') else ''
    end = f'{logr}, nº {num}{comp_str}, {bairro}, {cidade}/{estado}, CEP {cep}'

    return (
        f'{nome}, {nac}, {ec}, {nasc_v} em {nasc}, {med} inscrito(a) no CRM/{crm_uf} '
        f'sob o nº {crm} e no CPF sob nº {cpf}, {port} do RG nº {rg} {rg_orgao}/{rg_uf}, '
        f'{res} na {end};'
    )


def gerar_acs(
    template_bytes: bytes,
    excel_bytes: bytes,
    retirantes: list[str],
    num_alteracao: int,
    manter_preambulo: set[str],
    andres_texto: str = None,
    lidya_texto: str = None,
    auro_q_inicio: int = None,
    nomes_qsa_apos: list[str] = None,
    majoritario_nome: str = 'AURO BUFFANI CLAUDINO',
    majoritario_q_final: int = None,
    secundario_nome: str = None,
    secundario_q: int = None,
    total_capital: int = 10000,
) -> bytes:
    """
    Gera nova ACS editando o template.
    Retorna bytes do .docx gerado.
    """
    # Ler ingressantes do Excel
    dados_ing = ler_ingressantes_excel(excel_bytes)
    ingressantes_texto = [formatar_qualificacao(d) for d in dados_ing]

    doc = Document(io.BytesIO(template_bytes))
    p_list = list(doc.paragraphs)

    # Identificar templates por posição
    tpl_socio = None
    tpl_cessao = None
    tpl_ossocios_p = None

    for p in p_list:
        pPr = p._p.find(qn('w:pPr'))
        numPr = pPr.find(qn('w:numPr')) if pPr is not None else None
        if numPr is None:
            continue
        ilvl = numPr.find(qn('w:ilvl'))
        numId_el = numPr.find(qn('w:numId'))
        il = ilvl.get(qn('w:val')) if ilvl is not None else None
        ni = numId_el.get(qn('w:val')) if numId_el is not None else None

        txt = p.text.strip()
        if il == '0' and tpl_socio is None and len(txt) > 30:
            if any(kw in txt.lower() for kw in ['médico', 'médica', 'empresária', 'inscrito']):
                tpl_socio = p
        if il == '2' and '01 (uma) quota' in txt and tpl_cessao is None:
            tpl_cessao = p
        if il == '1' and txt == 'Os sócios:' and tpl_ossocios_p is None:
            tpl_ossocios_p = p

    # Passo 1: Título
    titulo_p = p_list[0]
    old_num = f'{num_alteracao - 1}ª'
    new_num = f'{num_alteracao}ª'
    full = ''.join(r.text or '' for r in titulo_p.runs)
    if titulo_p.runs:
        titulo_p.runs[0].text = full.replace(old_num, new_num)
        for r in titulo_p.runs[1:]:
            r.text = ''

    # Passo 2: Deletar sócios do preâmbulo não presentes no QSA atual
    in_preambulo = False
    paras_deletar = []
    for p in p_list:
        txt = p.text.strip()
        if 'Pelo presente instrumento' in txt:
            in_preambulo = True
            continue
        if in_preambulo and ('e, ainda' in txt or 'e ainda' in txt):
            break
        if in_preambulo:
            n = sem_acento(txt.split(',')[0].strip())
            is_socio = any(kw in txt.lower() for kw in ['médico', 'médica', 'empresária', 'inscrito', 'inscrita'])
            if is_socio and n and n not in manter_preambulo:
                paras_deletar.append(p)
    for p in paras_deletar:
        del_p(p)

    # Passo 3: Inserir ANDRES/LIDYA se necessário
    if andres_texto and tpl_socio:
        andrea_p = None
        for p in doc.paragraphs:
            if sem_acento(p.text.split(',')[0].strip()) == sem_acento('ANDRÉA CARLA OLIVEIRA GOMES'):
                andrea_p = p
                break
        if andrea_p:
            nome, resto = andres_texto.split(',', 1)
            new_p = clone_para(tpl_socio, [(nome.strip(), True), (', ' + resto.strip(), False)])
            andrea_p._p.addnext(new_p)

    if lidya_texto and tpl_socio:
        karyn_p = None
        for p in doc.paragraphs:
            if sem_acento(p.text.split(',')[0].strip()) == sem_acento('KARYN NEMETH'):
                karyn_p = p
                break
        if karyn_p:
            nome, resto = lidya_texto.split(',', 1)
            new_p = clone_para(tpl_socio, [(nome.strip(), True), (', ' + resto.strip(), False)])
            karyn_p._p.addnext(new_p)

    # Passo 4: Atualizar "itens (3) a (X)"
    total_itens = 2 + len([p for p in doc.paragraphs
                           if any(kw in p.text.lower() for kw in ['médico', 'médica'])]) + len(ingressantes_texto)
    for p in doc.paragraphs:
        if 'itens (3) a (' in p.text and 'procuração' in p.text:
            full = ''.join(r.text or '' for r in p.runs)
            new_full = re.sub(r'\(3\) a \(\d+\)', f'(3) a ({total_itens})', full)
            if p.runs:
                p.runs[0].text = new_full
                for r in p.runs[1:]:
                    r.text = ''
            break

    # Passo 5: "e, ainda" — substituir placeholder pelos ingressantes
    cc_para = None
    for p in doc.paragraphs:
        if p.text.strip() == 'cc':
            cc_para = p
            break
    if cc_para and tpl_socio:
        for texto in reversed(ingressantes_texto):
            nome, resto = texto.split(',', 1)
            new_p = clone_para(tpl_socio, [(nome.strip(), True), (', ' + resto.strip(), False)])
            cc_para._p.addprevious(new_p)
        del_p(cc_para)

    # Passo 6: Atualizar quotas do majoritário
    if auro_q_inicio:
        ext_map = {
            6951: 'seis mil novecentas e cinquenta e uma',
            6943: 'seis mil novecentas e quarenta e três',
            49875: 'quarenta e nove mil oitocentos e setenta e cinco',
            49881: 'quarenta e nove mil oitocentos e oitenta e uma',
        }
        ext = ext_map.get(auro_q_inicio, str(auro_q_inicio))
        for p in doc.paragraphs:
            if majoritario_nome in p.text and 'detentor de' in p.text:
                full = ''.join(r.text or '' for r in p.runs)
                new_full = re.sub(
                    r'detentor de [\d\.]+ \([^)]+\) quotas',
                    f'detentor de {auro_q_inicio:,} ({ext}) quotas'.replace(',', '.'),
                    full
                )
                if p.runs:
                    p.runs[0].text = new_full
                    for r in p.runs[1:]:
                        r.text = ''
                break

    # Passo 7: Substituir cessões antigas pelos novos ingressantes
    cessao_paras = []
    for p in doc.paragraphs:
        if '01 (uma) quota' in p.text and 'ingressa' in p.text:
            cessao_paras.append(p)

    if cessao_paras and tpl_cessao:
        ref_elem = cessao_paras[0]._p
        for i, texto_ing in enumerate(ingressantes_texto):
            nome = texto_ing.split(',')[0].strip()
            suf = '.' if i == len(ingressantes_texto) - 1 else ('; e' if i == len(ingressantes_texto) - 2 else ';')
            new_p = clone_para(tpl_cessao, [
                ('01 (uma) quota que detém do capital social da Sociedade para ', False),
                (nome, True),
                (f', que ingressa e passa a compor o quadro de sócios da Sociedade{suf}', False),
            ])
            ref_elem.addprevious(new_p)
        for cp in cessao_paras:
            del_p(cp)

    # Passo 8: Retirantes — remover tabela antiga + inserir novos
    acima_p = None
    for p in doc.paragraphs:
        if p.text.strip().startswith('Acima qualificados'):
            acima_p = p
            break

    body = doc.element.body
    if tpl_ossocios_p and acima_p:
        idx_os = list(body).index(tpl_ossocios_p._p)
        idx_ac = list(body).index(acima_p._p)
        for elem in list(body)[idx_os + 1:idx_ac]:
            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            if tag == 'tbl':
                body.remove(elem)
            elif tag == 'p':
                txt = ''.join(r.text or '' for r in
                              elem.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
                if not txt.strip():
                    body.remove(elem)

    if tpl_cessao and tpl_ossocios_p:
        ref_elem = tpl_ossocios_p._p
        for nome_ret in retirantes:
            new_p = clone_para(tpl_cessao, [(nome_ret, True)])
            ref_elem.addnext(new_p)
            ref_elem = new_p
        # Linha em branco após último retirante
        blank = copy.deepcopy(tpl_cessao._p)
        for r in list(blank.findall(qn('w:r'))):
            blank.remove(r)
        pPr_blank = blank.find(qn('w:pPr'))
        if pPr_blank is not None:
            np_el = pPr_blank.find(qn('w:numPr'))
            if np_el is not None:
                pPr_blank.remove(np_el)
        ref_elem.addnext(blank)

    # Passo 9: Atualizar tabelas QSA
    if nomes_qsa_apos and majoritario_q_final:
        def update_table(table):
            for row in list(table.rows[1:]):
                table._tbl.remove(row._tr)
            W = [432, 6480, 1224, 1269]

            def add_row(vals, bold=False):
                from docx.shared import Pt
                tr = OxmlElement('w:tr')
                for i, v in enumerate(vals):
                    tc = OxmlElement('w:tc')
                    tcPr = OxmlElement('w:tcPr')
                    tcW = OxmlElement('w:tcW')
                    tcW.set(qn('w:w'), str(W[i]))
                    tcW.set(qn('w:type'), 'dxa')
                    tcPr.append(tcW)
                    tc.append(tcPr)
                    p_elem = OxmlElement('w:p')
                    r_elem = OxmlElement('w:r')
                    rPr = OxmlElement('w:rPr')
                    rFonts = OxmlElement('w:rFonts')
                    rFonts.set(qn('w:ascii'), 'Calibri')
                    rFonts.set(qn('w:hAnsi'), 'Calibri')
                    sz = OxmlElement('w:sz')
                    sz.set(qn('w:val'), '18')
                    szCs = OxmlElement('w:szCs')
                    szCs.set(qn('w:val'), '18')
                    rPr.append(rFonts)
                    rPr.append(sz)
                    rPr.append(szCs)
                    if bold:
                        b = OxmlElement('w:b')
                        rPr.insert(0, b)
                    r_elem.append(rPr)
                    t = OxmlElement('w:t')
                    t.text = str(v)
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    r_elem.append(t)
                    p_elem.append(r_elem)
                    tc.append(p_elem)
                    tr.append(tc)
                table._tbl.append(tr)

            add_row(['', majoritario_nome, str(majoritario_q_final),
                     f'{majoritario_q_final:,}'.replace(',', '.')])
            if secundario_nome and secundario_q:
                add_row(['', secundario_nome, str(secundario_q),
                         f'{secundario_q:,}'.replace(',', '.')])
            for n, nome in enumerate(nomes_qsa_apos, 1):
                add_row([str(n), nome, '1', '1,00'])
            total_str = f'{total_capital:,}'.replace(',', '.')
            add_row(['', 'Total', f'{total_capital:,}'.replace(',', '.'), f'{total_capital:,},00'.replace(',', '.')],
                    bold=True)

        for tbl in doc.tables:
            if len(tbl.rows) > 5:
                update_table(tbl)

    # Passo 10: Data
    for p in doc.paragraphs:
        if 'São Paulo, SP,' in p.text:
            for r in p.runs:
                if any(str(y) in r.text for y in range(2020, 2030)):
                    r.text = '____ de __________ de 2026.'
            break

    # Salvar em bytes
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.read()
