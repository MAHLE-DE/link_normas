FERRAMENTA DE UNIAO - CAMINHOS DETERMINADOS PELO MAPA ATUAL

Mapa usado:
C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\estrutura_link_normas.json

A ferramenta nao possui caminhos fixos para a amostra nem para o banco fisico.
Ela le o mapa e exige exatamente:
- uma pasta chamada 05_DOCUMENTS;
- um arquivo chamado amostra_validacao_40.json.

No mapa de 24/11/2026, os caminhos encontrados sao:
- Banco fisico: C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\05_DOCUMENTS
- Amostra: C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\04_NOVO_TESTE_40\amostra_validacao_40.json
- Saidas: C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\04_NOVO_TESTE_40

Como executar:
1. Mantenha estrutura_link_normas.json na raiz link_normas.
2. Extraia este pacote.
3. Execute rodar_uniao_segundo_piloto_40_mapa_atual.bat.

Saidas:
- 40_normas_validacao_unificadas.pdf
- 40_normas_validacao_unificadas.manifesto.json
- 40_normas_validacao_unificadas.log.txt

Se o mapa futuro mudar os caminhos, basta regenerar estrutura_link_normas.json. O script passa a usar os novos caminhos sem edicao.
