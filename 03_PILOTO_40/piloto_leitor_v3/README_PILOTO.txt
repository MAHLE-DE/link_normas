PILOTO DO LEITOR - 40 NORMAS

CORRECAO DA VERSAO 2.1
A primeira versao nao instalava automaticamente o PyMuPDF, biblioteca que fornece o modulo "fitz".

EXECUCAO RECOMENDADA
1. Extraia todo o ZIP em uma unica pasta local.
2. De duplo clique em rodar_piloto.bat.
3. O BAT verifica e instala automaticamente:
   - PyMuPDF, modulo fitz, para leitura de PDF
   - python-docx, para leitura de Word
   - openpyxl, para Excel
   - pypdf, apoio para PDF
4. Depois, executa o piloto e gera:
   - resultado_piloto.xlsx
   - avaliacao_piloto.xlsx

SE A INSTALACAO AUTOMATICA FOR BLOQUEADA
Execute separadamente 00_instalar_dependencias.bat. Isso normalmente exige acesso ao repositório Python pela rede corporativa.

A pasta de origem usada pelo piloto e:
X:\15. Robustez DE-CAD\00 Dashboard - KPI\source_app\validstandards
