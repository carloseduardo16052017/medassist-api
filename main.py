"""
MedAssist AUTO — API FastAPI
Endpoint principal: POST /gerar-acs
"""
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
import json
from scripts.build_acs import gerar_acs, sem_acento

app = FastAPI(
    title="MedAssist AUTO API",
    description="API para geração automática de minutas ACS",
    version="1.0.0"
)

# CORS — permitir chamadas do Lovable/Supabase
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def root():
    return {"status": "ok", "service": "MedAssist AUTO API"}


@app.get("/health")
def health():
    return {"status": "healthy"}


@app.post("/gerar-acs")
async def gerar_acs_endpoint(
    template: UploadFile = File(..., description="Última ACS registrada (.docx)"),
    relatorio: UploadFile = File(..., description="Relatório ClickSign (.xlsx)"),
    retirantes: str = Form(..., description='JSON array com nomes dos retirantes. Ex: ["NOME 1","NOME 2"]'),
    num_alteracao: int = Form(..., description="Número da nova alteração. Ex: 6"),
    config: str = Form(..., description="""JSON com configurações da empresa:
    {
      "majoritario_nome": "AURO BUFFANI CLAUDINO",
      "majoritario_q_inicio": 6951,
      "majoritario_q_final": 6950,
      "secundario_nome": "YARA FERRAZ CALDARONE",
      "secundario_q": 2975,
      "total_capital": 10000,
      "manter_preambulo": ["NOME 1", "NOME 2", ...],
      "nomes_qsa_apos": ["NOME 1", "NOME 2", ...],
      "andres_texto": "ANDRES ORTIZ, brasileiro...",
      "lidya_texto": "LIDYA CAROLINA DE MELO, brasileira..."
    }
    """),
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

    nome_arquivo = f"{num_alteracao}a_ACS_v1.docx"
    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{nome_arquivo}"'}
    )
