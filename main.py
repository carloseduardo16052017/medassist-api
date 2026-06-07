"""
MedAssist AUTO — API FastAPI
Endpoints:
  POST /gerar-acs                  → Gera minuta ACS (.docx)
  POST /gerar-dbe-ingressantes     → Gera planilha DBE dos ingressantes (.xlsx)
  POST /gerar-dbe-retirantes       → Gera planilha DBE dos retirantes (.xlsx)
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
    version="1.3.0"
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
    Extrai AUTOMATICAMENTE os nomes dos retirantes da minuta.

    Padrão Word: os retirantes ficam entre 'Os sócios:' (ilvl=N) e
    'Acima qualificados' como parágrafos numerados (ilvl=N+1) em negrito,
    contendo apenas o nome — sem qualificação (sem CPF, CRM, endereço).

    Funciona tanto para QUALIFEMME (1.2.1./1.2.2.) quanto OPTIMUM e outros.
    """
    from docx.oxml.ns import qn as _qn
    doc = Document(io.BytesIO(docx_bytes))
    nomes = []
    in_retirantes = False
    os_socios_ilvl = None

    QUALIFICACAO_KWS = ('médico', 'médica', 'cpf', 'crm', 'rg nº', 'cep',
                        'nascido', 'nascida', 'inscrito', 'inscrita', 'portador')

    for p in doc.paragraphs:
        txt = p.text.strip()

        # ── Detectar cabeçalho "Os sócios:" ──────────────────────────
        if txt in ('Os sócios:', 'Os sócios') or txt.endswith('Os sócios:'):
            in_retirantes = True
            pPr = p._p.find(_qn('w:pPr'))
            if pPr is not None:
                numPr = pPr.find(_qn('w:numPr'))
                if numPr is not None:
                    il = numPr.find(_qn('w:ilvl'))
                    if il is not None:
                        try:
                            os_socios_ilvl = int(il.get(_qn('w:val')))
                        except (TypeError, ValueError):
                            os_socios_ilvl = None
            continue

        # ── Fim da seção ──────────────────────────────────────────────
        if in_retirantes and txt.startswith('Acima qualificados'):
            break

        # ── Dentro da seção: identificar nomes ───────────────────────
        if not in_retirantes or not txt:
            continue

        # Verificar se parece qualificação (não é um nome de retirante)
        txt_lower = txt.lower()
        has_qualification = any(kw in txt_lower for kw in QUALIFICACAO_KWS)
        if has_qualification:
            continue

        # Verificar numPr: retirantes têm ilvl = os_socios_ilvl + 1
        pPr = p._p.find(_qn('w:pPr'))
        numPr = pPr.find(_qn('w:numPr')) if pPr is not None else None
        is_bold = any(r.font.bold for r in p.runs if r.font.bold is True)

        if numPr is not None:
            il = numPr.find(_qn('w:ilvl'))
            if il is not None:
                try:
                    ilvl = int(il.get(_qn('w:val')))
                    # ilvl deve ser 1 nível abaixo de "Os sócios:"
                    expected = (os_socios_ilvl + 1) if os_socios_ilvl is not None else ilvl
                    if ilvl == expected and is_bold:
                        nomes.append(txt)
                        continue
                except (TypeError, ValueError):
                    pass

        # Fallback: parágrafo bold, sem vírgula, entre as marcas certas
        if is_bold and ',' not in txt and len(txt) > 5:
            nomes.append(txt)

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
        cpf_match = re.search(r'CPF sob n[°º]\s*([\d]{3}\.[\d]{3}\.[\d]{3}-[\d]{2})', txt)
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
    Extrai automaticamente os nomes e dados dos retirantes da minuta (.docx).
    Não é necessário informar os nomes separadamente.
    """
    try:
        docx_bytes = await minuta.read()
        dados = extrair_dados_retirantes_docx(docx_bytes)
        if not dados:
            raise HTTPException(
                status_code=404,
                detail="Nenhum retirante encontrado na seção '1.2. Os sócios:' da minuta."
            )
        xlsx_bytes = gerar_planilha_dbe(dados)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao gerar DBE retirantes: {str(e)}")

    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="DBE_Retirantes.xlsx"'}
    )
