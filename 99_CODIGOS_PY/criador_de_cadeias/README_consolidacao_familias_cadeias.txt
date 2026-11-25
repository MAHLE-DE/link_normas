CONSOLIDAÇÃO DAS FAMÍLIAS DE CADEIAS PARA V7.5

Objetivo
Concluir tudo que for comprovável antes do formato X-X-XX-NNNN, especialmente WNA, WNC, WNF, WNE, WNL e WNV, sem inferir equivalência pela semelhança numérica.

Entradas padrão
- 02_BANCO_ATUAL\cadastro_canonico_v3_SEM_URL_SHAREPOINT.json
- 05_DOCUMENTS\ (PDF e DOCX)
- decisões e grafos aprovados ainda existentes dentro de link_normas

Saídas em 02_BANCO_ATUAL
1. familia_cadeias_completa_v7_5.xlsx
   Única planilha de trabalho e revisão, com as abas:
   - Resumo
   - Familias_cadeias
   - Elos_comprovados
   - Pendencias_legado
   - Ocorrencias_auditoria
   - Legados_finais
   - Exclusoes
   - Conflitos
   - Regras

2. grafo_cadeias_v7_5_completo.json
   Arquivo operacional para a V7.5. Contém alias_to_current e elos comprovados.

3. pendencias_cadeias_v7_5.json
   Pendências e conflitos para correção posterior.

Como executar
1. Extraia este pacote.
2. Execute rodar_consolidacao_familias_cadeias.bat.
3. Envie os três arquivos gerados.

Critérios
- Previous Document No -> Document No é elo oficial.
- WNA/WNF/WNC/WNE/WNL/WNV só entra automaticamente se houver evidência explícita ou decisão aprovada.
- O programa não converte por coincidência numérica.
- Códigos sem comprovação ficam em Pendencias_legado.
- Códigos excluídos permanecem em Exclusoes e não serão usados pela V7.5.
- WNA 687 permanece legado final válido.

Decisões já incorporadas
- WNA 255 -> S-F-PT-0255-en
- WNA 685 -> S-F-PT-0685-en
- WNA 690 -> S-F-PT-0690-en
- WNA 693 -> S-F-PT-0693-en
- WNA 70 -> S-G-PT-0070-en
- WNA 5072 -> S-H-PT-5072-de
- WNA 5105 -> S-L-PT-5105-en
- WNA 5115 -> S-L-PT-5115-en
- WNF 102810 -> S-F-ML-2810-en
- WNC 100002 -> S-C-TG-0002-en
- WNA 687 como legado final válido
- WNA 280, WNA 5046 e WNA 60521 como não considerar
