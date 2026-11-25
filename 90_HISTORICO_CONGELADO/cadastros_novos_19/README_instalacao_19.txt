INSTALAÇÃO DOS 19 NOVOS CADASTROS

1. Extraia este pacote em uma pasta local.
2. Antes de instalar, revise principalmente os cinco registros com categoria/família "process":
   SOP-0000603, P-B-0796-en, P-B-132-en, W-B-720-en, W-BE-1002-en e W-BE-7200-en.
   Observação: são seis documentos. A categoria "process" ainda não existia nos 523 cadastros anteriores; confirme que a aplicação aceita essa nova família.
3. Execute no PowerShell:
   py -3 .\04_instalar_cadastros_novos.py
4. O instalador interrompe sem copiar nada se:
   - algum document_no já existir;
   - algum JSON já existir pelo mesmo nome;
   - algum arquivo físico não estiver em validstandards;
   - a quantidade não for exatamente 19.
5. Após a instalação, gere novamente os dois mapas:
   py -3 .\ferramentas_mapeamento\01_compilar_banco_cadastros.py
   py -3 .\ferramentas_mapeamento\02_mapear_validstandards.py

O arquivo cadeias_novas_grupos_1_a_7.json NÃO é instalado no banco de cadastros. Ele será usado na etapa posterior de construção das cadeias.
