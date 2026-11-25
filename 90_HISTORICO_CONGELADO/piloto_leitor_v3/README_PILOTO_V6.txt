PILOTO V6

1. Extraia os arquivos deste ZIP dentro da pasta original do piloto V5.
2. Execute rodar_piloto_v6.bat.
3. Envie resultado_piloto_v6.xlsx e avaliacao_piloto_v6.xlsx.

A pasta deve manter piloto_40_normas.json, equivalences_pilot.json e base_mestra_cadeias_consolidada.xlsx do piloto anterior.
A busca dos documentos usa primeiro:
C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\Valid-Documents
Se necessário, usa validstandards na unidade X:.

MELHORIAS
- tabelas extraídas com PyMuPDF page.find_tables();
- cabeçalho de coluna usado como contexto técnico;
- listas densas com 3 ou mais códigos aceitas fora de histórico e exemplos;
- sobreposição DIN EN ISO / ISO eliminada na saída;
- REACH e GADSL somente em caixa alta;
- autorreferência normalizada sem idioma;
- códigos após former/formerly/instead of/replaced by classificados como históricos;
- ISO 11357-1/-2 expandido;
- ISO 9227 NSS e ISO 16232 Parts 1-10 normalizados;
- referência bruta, página, seção e evidência preservadas.
