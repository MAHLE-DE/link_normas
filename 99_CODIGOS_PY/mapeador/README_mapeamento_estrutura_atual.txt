MAPEAMENTO DA ESTRUTURA ATUAL DE link_normas

Arquivos deste pacote:
- mapear_estrutura_atual.py
- rodar_mapeamento_estrutura_atual.bat
- README_mapeamento_estrutura_atual.txt

OBJETIVO
Mapear recursivamente todos os arquivos e subpastas existentes em:
C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas

O resultado será:
C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\estrutura_atual_link_normas.json

COMO EXECUTAR
1. Extraia os três arquivos deste ZIP em qualquer pasta.
2. Dê duplo clique em rodar_mapeamento_estrutura_atual.bat.
3. Aguarde a mensagem "Mapeamento concluído".
4. Envie o arquivo estrutura_atual_link_normas.json.

O QUE O JSON CONTÉM
- pasta raiz;
- resumo de quantidade de pastas e arquivos;
- contagem por extensão;
- tamanho total;
- lista completa de subpastas;
- lista de arquivos com nome, extensão, caminho relativo, caminho absoluto, tamanho e data de modificação;
- erros de leitura, se houver.

SEGURANÇA
- A ferramenta é somente leitura em relação ao conteúdo existente.
- Nenhum arquivo ou pasta é movido, renomeado ou excluído.
- O próprio arquivo estrutura_atual_link_normas.json é ignorado no inventário para não se inventariar a si mesmo.
- Pastas técnicas como __pycache__ e .git são ignoradas.

OPCIONAL: HASH SHA-256
Para calcular o SHA-256 de todos os arquivos, execute no PowerShell:
py -3 .\mapear_estrutura_atual.py --hash

Isso pode demorar bastante devido ao tamanho dos PDFs. Para o primeiro diagnóstico, não é necessário usar --hash.

CAMINHO DIFERENTE
Se a pasta raiz mudar, execute:
py -3 .\mapear_estrutura_atual.py --root "C:\novo\caminho"
