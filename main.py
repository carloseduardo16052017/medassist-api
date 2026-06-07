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
    Extrai AUTOMATICAMENTE os nomes dos retirantes da seção 1.2.

    Regra: parágrafos **bold**, sem texto de qualificação (CPF, CRM, etc.)
    e sem vírgula (só o nome), localizados entre 'Os sócios:' e
    'Acima qualificados' no Capital Social.

    Funciona para QUALIFEMME (ilvl=0) e OPTIMUM (ilvl=2) e qualquer empresa.
    """
    doc = Document(io.BytesIO(docx_bytes))
    nomes = []
    in_retirantes = False

    # Palavras que indicam qualificação — parágrafo com qualquer uma dessas NÃO é retirante
    QUALIFICACAO_KWS = (
        'médico', 'médica', 'cpf', 'crm', 'rg nº', 'cep',
        'nascido', 'nascida', 'inscrito', 'inscrita', 'portador',
        'portadora', 'residente', 'domiciliado', 'domiciliada',
        'empresária', 'empresário', 'brasileiro', 'brasileira',
    )

    for p in doc.paragraphs:
        txt = p.text.strip()

        # ── Início: localizar cabeçalho "Os sócios:" na seção 1.2 ────
        # Pode vir como texto puro ou como item numerado automaticamente pelo Word
        if not in_retirantes:
            if txt in ('Os sócios:', 'Os sócios') or txt.endswith('Os sócios:'):
                in_retirantes = True
            continue

        # ── Fim: parar ao encontrar "Acima qualificados" ──────────────
        if txt.startswith('Acima qualificados'):
            break

        # ── Pular parágrafos vazios ────────────────────────────────────
        if not txt:
            continue

        # ── Descartar se contém texto de qualificação ─────────────────
        txt_lower = txt.lower()
        if any(kw in txt_lower for kw in QUALIFICACAO_KWS):
            continue

        # ── Descartar se contém vírgula (qualificação ou cessão) ──────
        # Nomes de retirantes são sem vírgula; cessões (1.1.x) têm vírgulas
        if ',' in txt:
            continue

        # ── Aceitar: parágrafo bold com apenas o nome ─────────────────
        is_bold = any(r.font.bold for r in p.runs if r.font.bold is True)
        if is_bold and len(txt) > 5:
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
