GERAÇÃO DO CADASTRO CANÔNICO V3

OBJETIVO
Criar duas versões claramente separadas do cadastro canônico:

1. cadastro_canonico_v3_COM_URL_SHAREPOINT.json
   - contém Item ID, List ID, URL do item e URL do arquivo SharePoint;
   - indicado para administração, auditoria e atualização do banco.

2. cadastro_canonico_v3_SEM_URL_SHAREPOINT.json
   - remove deliberadamente todas as URLs, List IDs e Item IDs do SharePoint;
   - indicado para análise, compartilhamento interno restrito e processamento que não precise acessar o SharePoint.

Também cria:
- diagnostico_cadastro_canonico_v3.xlsx

ENTRADAS PADRÃO
- C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\02_BANCO_ATUAL\Baixar cadastro_sharepoint_completo_corrigido.json
- C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\02_BANCO_ATUAL\cadastro_canonico_v2.json
- Planilha Valid-Documents mais recente encontrada em 02_BANCO_ATUAL ou em link_normas
- C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\05_DOCUMENTS

COMO EXECUTAR
1. Extraia os três arquivos deste pacote.
2. Execute rodar_geracao_cadastro_canonico_v3.bat.
3. Os resultados serão gravados em 02_BANCO_ATUAL.

REGRAS APLICADAS
- prefere inglês quando existem versões inglesa e alemã;
- mantém alemão somente quando não existe inglês;
- registra versões alemãs descartadas em language_variants;
- gera document_id UUID5 único e determinístico;
- usa Document No como identidade atual;
- usa Previous Document No como alias histórico, removendo apenas a versão entre parênteses;
- Revision No não participa da identidade;
- preserva o arquivo físico exato retornado pelo SharePoint;
- mantém W-BE-0885-en como exceção original vigente fora da lista oficial.

IMPORTANTE SOBRE A PLANILHA VALID DOCUMENTS
A ferramenta procura automaticamente a planilha mais recente cujo nome começa por Valid-Documents.
Se a planilha não estiver dentro de link_normas, execute manualmente:

py -3 .\gerar_cadastro_canonico_v3.py --valid-documents "C:\caminho\Valid-Documents-AAAA-MM-DD.xlsx"

DIFERENÇA ENTRE AS DUAS VERSÕES
O arquivo SEM_URL_SHAREPOINT não contém:
- sharepoint_item_url
- sharepoint_file_url
- sharepoint_list_id
- sharepoint_item_id
- caminhos de origem que contenham URL

Os dois arquivos mantêm os mesmos documentos, códigos, títulos, idiomas, revisões, nomes anteriores e vínculos com arquivos locais.
