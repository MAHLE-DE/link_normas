COMPARAÇÃO DO CADASTRO SHAREPOINT COM 05_DOCUMENTS

Entrada padrão:
C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\02_BANCO_ATUAL\Baixar cadastro_sharepoint_completo_corrigido.json

Fallback aceito automaticamente:
C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\02_BANCO_ATUAL\cadastro_sharepoint_completo_corrigido.json

Pasta física:
C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\05_DOCUMENTS

COMO EXECUTAR
1. Extraia os três arquivos do ZIP em qualquer pasta.
2. Execute rodar_comparacao_sharepoint_vs_05_documents.bat.
3. Aguarde a conclusão.

SAÍDAS EM 02_BANCO_ATUAL
- cadastro_canonico_v2.json
- diagnostico_sharepoint_vs_05_documents.xlsx
- comparacao_sharepoint_vs_05_documents.json

CLASSIFICAÇÕES
- CORRESPONDENCIA_EXATA: mesmo nome exato retornado pelo SharePoint.
- CORRESPONDENCIA_NORMALIZADA: diferença apenas de caixa, Unicode ou espaços.
- CANDIDATO_RELAXADO_REVISAR: diferenças de separadores; exige revisão manual.
- ARQUIVO_OFICIAL_AUSENTE_LOCALMENTE: SharePoint retornou o nome, mas o arquivo não existe em 05_DOCUMENTS.
- DUPLICADO_LOCAL_EXATO: o mesmo nome existe em mais de um caminho local.
- AMBIGUO_NORMALIZADO / AMBIGUO_RELAXADO: múltiplos candidatos.

ABAS DO EXCEL
- Resumo
- Comparacao_completa
- Ausentes_localmente
- Revisar
- Arquivos_locais_extras
- Conflitos_arquivo
- Duplicados_locais
- Erros

REGRAS IMPORTANTES
- O nome exato devolvido pelo SharePoint é a autoridade para associação física.
- Correspondência relaxada nunca é aceita automaticamente.
- A ferramenta não move, renomeia ou exclui arquivos.
- São analisados PDF, DOC, DOCX, XLS, XLSX, PPT e PPTX, recursivamente.

HASH OPCIONAL
Para calcular SHA-256 dos arquivos locais, execute no PowerShell:
py -3 .\comparar_sharepoint_com_05_documents.py --hash

O hash deixa a execução mais lenta e não é necessário na primeira comparação.
