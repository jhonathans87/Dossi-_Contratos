
import io
import os
import re
from datetime import date
from typing import Any

import streamlit as st
from pypdf import PdfReader
from docx import Document
from openai import OpenAI
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Dossiê Imobiliário", page_icon="🏢", layout="wide")

MASTER_PROMPT = """
Crie um Dossiê Executivo de Valorização Patrimonial e Análise Contratual,
em português do Brasil, usando exclusivamente os documentos enviados,
o valor atual de tabela e a data-base.

Regras:
- Não invente dados, cláusulas, páginas, índices ou taxas.
- Informe a origem de cada dado relevante.
- Aplique a hierarquia: aditivo mais recente, aditivos anteriores,
  contrato original, extrato financeiro e informação do usuário.
- O aditivo prevalece somente sobre o que alterar expressamente.
- Diferencie valor original, valor ajustado, total pago nominal,
  total pago corrigido, recebimento líquido e saldo devedor.
- Destaque todas as divergências.
- Não apresente valorização bruta como lucro líquido.
- Não afirme patrimônio de afetação sem previsão expressa.
- Não estime CDI. Sem série oficial, informe que a simulação não foi concluída.
- Indique cláusula/item e página em temas jurídicos quando localizados.

Estrutura:
1. Identificação da unidade
2. Resumo executivo
3. Documentos analisados e ausentes
4. Dados contratuais originais
5. Alterações por aditivos
6. Conferência financeira
7. Inconsistências
8. Evolução patrimonial
9. Rentabilidades
10. Multiplicador patrimonial
11. Patrimônio líquido estimado
12. CDI e comparativo
13. Prazo de entrega
14. Multa por atraso
15. Inadimplência do adquirente
16. Patrimônio de afetação
17. Quadro executivo
18. Ranking
19. Conclusão executiva
"""

def clean(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def extract_pdf(raw: bytes) -> tuple[str, int]:
    reader = PdfReader(io.BytesIO(raw))
    pages = []
    for number, page in enumerate(reader.pages, 1):
        content = page.extract_text() or ""
        pages.append(f"\n--- PÁGINA {number} ---\n{content}")
    return clean("\n".join(pages)), len(reader.pages)

def extract_docx(raw: bytes) -> str:
    doc = Document(io.BytesIO(raw))
    output = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            output.append(" | ".join(cell.text.strip() for cell in row.cells))
    return clean("\n".join(output))

def extract_file(uploaded) -> dict[str, Any]:
    raw = uploaded.getvalue()
    ext = uploaded.name.lower().rsplit(".", 1)[-1]
    pages = None
    if ext == "pdf":
        text, pages = extract_pdf(raw)
    elif ext == "docx":
        text = extract_docx(raw)
    elif ext in {"txt", "md", "csv"}:
        text = clean(raw.decode("utf-8", errors="ignore"))
    else:
        text = ""
    return {"name": uploaded.name, "pages": pages, "text": text, "chars": len(text)}

def response_text(response) -> str:
    direct = getattr(response, "output_text", None)
    if direct and str(direct).strip():
        return str(direct).strip()

    parts = []
    for item in getattr(response, "output", []) or []:
        for content in getattr(item, "content", []) or []:
            value = getattr(content, "text", None)
            if isinstance(value, str):
                parts.append(value)
            elif value is not None:
                nested = getattr(value, "value", None)
                if nested:
                    parts.append(str(nested))
    return "\n".join(parts).strip()

def generate_report(doc_text: str, table_value: float, base_date: date, notes: str):
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    model = (os.getenv("OPENAI_MODEL") or "gpt-5-mini").strip()

    if not key:
        raise RuntimeError("OPENAI_API_KEY não configurada no Render.")
    if key.lower() in {"sua chave", "sua_chave", "sua_chave_aqui"}:
        raise RuntimeError("OPENAI_API_KEY ainda contém um valor de exemplo.")

    client = OpenAI(api_key=key, timeout=180.0, max_retries=2)
    user_input = f"""
VALOR ATUAL DE TABELA: R$ {table_value:,.2f}
DATA-BASE: {base_date.strftime('%d/%m/%Y')}
OBSERVAÇÕES: {notes or 'Nenhuma.'}

DOCUMENTOS:
{doc_text[:90000]}
"""
    response = client.responses.create(
        model=model,
        instructions=MASTER_PROMPT,
        input=user_input,
        max_output_tokens=10000,
    )
    report = response_text(response)
    diagnostics = {
        "modelo": model,
        "response_id": getattr(response, "id", "não informado"),
        "status": getattr(response, "status", "não informado"),
        "caracteres_enviados": len(user_input),
        "caracteres_recebidos": len(report),
    }
    if not report:
        diagnostics["incomplete_details"] = str(
            getattr(response, "incomplete_details", "não informado")
        )
        raise RuntimeError(f"A API respondeu sem texto. Diagnóstico: {diagnostics}")
    return report, diagnostics

def make_pdf(text: str) -> bytes:
    buffer = io.BytesIO()
    styles = getSampleStyleSheet()
    story = [Paragraph("Dossiê Executivo Imobiliário", styles["Title"]), Spacer(1, 12)]
    for line in text.splitlines():
        line = line.strip()
        if not line:
            story.append(Spacer(1, 6))
        elif line.startswith("#"):
            story.append(Paragraph(line.lstrip("# ").strip(), styles["Heading2"]))
        else:
            safe = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(safe, styles["BodyText"]))
    SimpleDocTemplate(buffer, pagesize=A4).build(story)
    return buffer.getvalue()

st.title("🏢 Dossiê Executivo Imobiliário")
st.caption("Versão corrigida para Render, com diagnóstico de extração e resposta da API.")

uploads = st.file_uploader(
    "Contrato, aditivos, extrato e demais documentos",
    type=["pdf", "docx", "txt", "md", "csv"],
    accept_multiple_files=True
)

col1, col2 = st.columns(2)
with col1:
    table_value = st.number_input(
        "Valor atual de tabela (R$)",
        min_value=0.0,
        step=1000.0,
        format="%.2f"
    )
with col2:
    base_date = st.date_input("Data-base", value=date.today())

notes = st.text_area("Observações complementares")

documents = []
if uploads:
    for uploaded in uploads:
        try:
            documents.append(extract_file(uploaded))
        except Exception as exc:
            st.error(f"Erro ao ler {uploaded.name}: {exc}")

if documents:
    st.subheader("Diagnóstico dos arquivos")
    st.dataframe(
        [{
            "Arquivo": d["name"],
            "Páginas": d["pages"] or "-",
            "Caracteres extraídos": d["chars"],
            "Situação": "Texto localizado" if d["chars"] >= 100 else "Provável PDF escaneado"
        } for d in documents],
        use_container_width=True,
        hide_index=True
    )

button = st.button(
    "Processar e gerar dossiê",
    type="primary",
    use_container_width=True,
    disabled=not documents or table_value <= 0
)

if button:
    try:
        usable = [d for d in documents if d["chars"] >= 100]
        if not usable:
            raise RuntimeError(
                "Nenhum documento possui texto suficiente. "
                "Os PDFs provavelmente são imagens escaneadas e precisam de OCR."
            )

        combined = "\n\n".join(
            f"===== {d['name']} =====\n{d['text']}" for d in usable
        )
        if len(combined) < 500:
            raise RuntimeError("Texto documental insuficiente para análise.")

        with st.spinner("Analisando documentos e gerando dossiê..."):
            report, diagnostics = generate_report(
                combined, table_value, base_date, notes
            )
        st.session_state["report"] = report
        st.session_state["diagnostics"] = diagnostics
    except Exception as exc:
        st.error(str(exc))
        st.info(
            "No Render, abra Logs para conferir o erro completo. "
            "Depois de alterar variáveis ou arquivos, use "
            "Manual Deploy → Clear build cache & deploy."
        )

if "report" in st.session_state:
    st.success("Dossiê gerado com sucesso.")
    st.markdown(st.session_state["report"])

    with st.expander("Diagnóstico técnico"):
        st.json(st.session_state.get("diagnostics", {}))

    st.download_button(
        "Baixar PDF",
        data=make_pdf(st.session_state["report"]),
        file_name="dossie_executivo_imobiliario.pdf",
        mime="application/pdf",
        use_container_width=True
    )
