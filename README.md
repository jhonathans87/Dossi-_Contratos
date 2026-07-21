# Dossiê Executivo Imobiliário

Aplicação web em Streamlit para:

- anexar contrato original;
- anexar múltiplos aditivos;
- anexar extrato financeiro;
- informar o valor atual de tabela;
- definir a data-base;
- gerar um dossiê patrimonial, financeiro e contratual;
- baixar o relatório em PDF e Markdown.

## Execução local

1. Instale Python 3.11 ou superior.
2. Crie um ambiente virtual.
3. Instale as dependências:

```bash
pip install -r requirements.txt
```

4. Configure a chave:

Linux/macOS:

```bash
export OPENAI_API_KEY="sua_chave"
export OPENAI_MODEL="gpt-5-mini"
```

Windows PowerShell:

```powershell
$env:OPENAI_API_KEY="sua_chave"
$env:OPENAI_MODEL="gpt-5-mini"
```

5. Execute:

```bash
streamlit run app.py
```

## Hospedagem no Render

1. Crie um repositório privado no GitHub.
2. Envie estes arquivos para o repositório.
3. No Render, crie um novo Web Service conectado ao repositório.
4. O arquivo `render.yaml` e o `Dockerfile` já estão preparados.
5. Cadastre a variável secreta `OPENAI_API_KEY`.
6. Confirme o modelo em `OPENAI_MODEL`.

## Hospedagem no Streamlit Community Cloud

1. Envie o projeto para um repositório privado ou público compatível com sua conta.
2. Crie um novo app e indique `app.py`.
3. Em Secrets, informe:

```toml
OPENAI_API_KEY = "sua_chave"
OPENAI_MODEL = "gpt-5-mini"
```

Observação: esta versão lê PDFs com camada de texto. PDFs integralmente digitalizados podem exigir um módulo de OCR.

## Segurança recomendada antes de uso real

Para documentos contratuais e dados pessoais, implemente:

- autenticação de usuários;
- controle por empresa e perfil;
- criptografia de arquivos;
- exclusão automática de documentos;
- registro de auditoria;
- política de retenção;
- aceite de termos;
- revisão jurídica e adequação à LGPD;
- armazenamento privado;
- backups;
- mascaramento de CPF e demais dados sensíveis.

## Limites do MVP

- A análise depende da qualidade do texto extraído.
- Tabelas complexas podem exigir conferência manual.
- O PDF gerado é um relatório inicial.
- O cálculo exato de CDI ainda exige integração com a série histórica oficial e um motor de capitalização por pagamento.
- Nenhuma conclusão deve substituir revisão jurídica, contábil ou financeira.
