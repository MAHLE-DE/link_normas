V7.3

1. Extraia o pacote na mesma pasta da V7.2.
2. Execute rodar_piloto_v7_3.bat.
3. Envie resultado_piloto_v7_3.xlsx e avaliacao_piloto_v7_3.xlsx.

A V7.3:
- carrega o grafo de cadeias de todos os JSONs disponíveis na pasta;
- resolve o ground truth e a saída antes da comparação;
- elimina relações históricas de substituição;
- suprime famílias genéricas quando partes específicas existem;
- aplica regra de auditoria a exemplos de DSGN-348;
- recupera context_rejected somente com evidência normativa forte;
- deduplica por norma-pai atual + norma citada atual;
- registra linhas promovidas e removidas em abas próprias.
