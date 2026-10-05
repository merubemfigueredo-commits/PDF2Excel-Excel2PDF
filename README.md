# Conversor PDF e Excel

Aplicativo feito com Python e Streamlit para:

- Extrair tabelas e texto selecionável de arquivos PDF para uma planilha `.xlsx`.
- Converter as abas de um arquivo `.xlsx` em um único PDF.

## Requisitos

- Python 3.10 ou superior
- pip

## Executar no computador

Abra um terminal nesta pasta e execute:

### Windows

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

### macOS ou Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

O Streamlit mostrará o endereço local para abrir no navegador.

## Limitações conhecidas

- PDFs escaneados ou compostos apenas por imagens precisam passar por OCR antes de
  serem convertidos. O aplicativo não executa OCR.
- A extração de tabelas depende da estrutura do PDF; documentos sem tabelas detectáveis
  são exportados como texto por página, quando houver texto selecionável.
- Na conversão de Excel para PDF, os valores das células são exportados; cores, fontes,
  gráficos e demais estilos da planilha não são preservados. Fórmulas são exibidas
  como texto, sem cálculo.
- Arquivos de entrada estão limitados a 50 MB.

## Privacidade

O aplicativo converte os arquivos em memória durante a sessão. Não grava os arquivos
enviados em disco nem os envia para serviços externos.
