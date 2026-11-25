PILOTO V5

COMO EXECUTAR
1. Extraia os arquivos deste ZIP dentro da pasta original do piloto V3.
2. A pasta deve conter também:
   - piloto_40_normas.json
   - equivalences_pilot.json
   - base_mestra_cadeias_consolidada.xlsx
3. Execute rodar_piloto_v5.bat.
4. Envie:
   - resultado_piloto_v5.xlsx
   - avaliacao_piloto_v5.xlsx

ALTERACOES DO EXTRATOR
- Reativa leitura do corpo técnico para MATL, TEST, MEAS e MNFG.
- Aceita tabelas e linhas com termos de método, ensaio, medição e requisito.
- Mantém DSGN mais restritivo para evitar códigos de exemplos e índices.
- Fecha Referenced Documents quando uma nova seção começa.
- Exclui Lessons Learned, histórico, keywords, clientes e métodos proprietários.
- Preserva evidência, página, seção e referência bruta.
- Reconhece referências compostas e qualificadores com maior extensão.
- Aplica a cadeia oficial até MEAS-579.

ALTERACOES DO AVALIADOR
- Usa somente as 40 normas e relações Nível 1.
- Compara códigos vigentes.
- Separa referências compostas da planilha manual.
- Reconhece equivalência hierárquica entre família e parte, sem apagar a referência específica da saída.
- Normaliza DIN EN ISO / DIN ISO / ISO para comparação.
