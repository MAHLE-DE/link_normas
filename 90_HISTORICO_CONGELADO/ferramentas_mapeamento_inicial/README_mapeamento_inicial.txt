MAPEAMENTO INICIAL DAS NORMAS

1. Copie estes quatro arquivos para uma pasta local:
   - 01_compilar_banco_cadastros.py
   - 02_mapear_validstandards.py
   - rodar_mapeamento.bat
   - README_mapeamento_inicial.txt

2. Dê duplo clique em rodar_mapeamento.bat.

3. A ferramenta usa automaticamente:
   Cadastros:
   C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\Banco_de_cadastros

   Documentos reais:
   X:\15. Robustez DE-CAD\00 Dashboard - KPI\source_app\validstandards

4. Serão gerados:
   C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\banco_cadastros_compilado.json
   C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\inventario_validstandards.json

5. O primeiro arquivo contém todos os cadastros completos, além de um resumo e verificações de qualidade.
6. O segundo contém todos os nomes e caminhos de PDF/DOC/DOCX existentes em validstandards.
7. Por padrão, o inventário não calcula hash para ser rápido. Para detectar arquivos binariamente iguais, execute:
   py -3 02_mapear_validstandards.py --hash

Observação: o mapeador percorre também eventuais subpastas.
