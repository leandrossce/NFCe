DANFE NFC-e PDF
Principais instalações (dependências)
Recomendado: Python 3.8+ (testado com 3.10+).
Via pip (Windows / Linux / macOS)

# Núcleo
pip install lxml reportlab qrcode[pil] pillow

# Excel (opcional, só se for exportar planilha)
pip install pandas openpyxl

Como usar
Gere PDF(s) do DANFE NFC-e a partir de XML e, opcionalmente, exporte os itens para Excel. Funciona via GUI ou CLI.
GUI (interface gráfica)
python danfe_nfce_pdf.py --gui
1. Selecione o diretório com XMLs.
2. Escolha o diretório de saída dos PDFs.
3. Defina opções (A4 ou 80mm, padrão glob, busca recursiva).
4. (Opcional) Ative e escolha o caminho do Excel para exportar itens.
5. Clique em Converter.
CLI
1) Diretório → vários PDFs (+ Excel opcional)

python danfe_nfce_pdf.py "F:\DIRETORIO_XML" "F:\SAIDA"

# Com Excel dos itens:
python danfe_nfce_pdf.py "F:\DIRETORIO_XML" "F:\SAIDA" --excel "F:\SAIDA\NFCe_itens.xlsx"

2) Arquivo único → um PDF (+ Excel opcional)

python danfe_nfce_pdf.py "F:\um.xml" "F:\saida.pdf"

# Com Excel:
python danfe_nfce_pdf.py "F:\um.xml" "F:\saida.pdf" --excel "F:\itens.xlsx"

python danfe_nfce_pdf.py "F:\um.xml" "F:\SAIDA\" --use-chave

Opções (CLI)

--paper {A4|80mm}: tamanho do papel (padrão: A4).
--glob "<padrão>": padrão de busca quando a entrada é diretório (ex.: --glob "*2025*.xml").
--recursive: busca também em subpastas (quando a entrada é diretório).
--excel <caminho.xlsx>: exporta itens para Excel (requer pandas + openpyxl).
--use-chave: ao salvar em diretório com arquivo único, nomeia o PDF pela chave de acesso (se disponível).
--gui: abre a interface gráfica.

Saídas

PDF(s): um por XML processado (A4 ou 80mm).
Excel (opcional): colunas DATA EMISSÃO, CHAVE ELETRÔNICA, CÓD, DESCRIÇÃO, QTD, UN, V.UNIT, V.TOTAL.

Dicas:
O script tenta extrair a chave de infNFe/@Id ou protNFe/infProt/chNFe.
Sem pandas/openpyxl, apenas os PDFs são gerados.
Para impressoras térmicas, prefira --paper 80mm.

