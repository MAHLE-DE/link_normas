# Extrator único de referências normativas

`extrator_normas.py` lê PDF (PyMuPDF, texto unido por página) e DOCX
(parágrafos e tabelas na ordem original). Produz pares únicos de norma-pai e
norma citada, preservando código bruto, páginas, seção, arquivo e evidência.
O arquivo XLSX tem `Correlacoes`, `Descartes`, `Ignorados` e, com gabarito,
`Avaliacao`.
Um JSON irmão contém os mesmos registros e métricas.

Requer Python **3.10 ou mais recente**, PyMuPDF **1.23.0 ou mais recente**,
python-docx e openpyxl. Os testes usam pytest. O import aceita tanto
`pymupdf` quanto o nome antigo `fitz`. O mínimo de PyMuPDF cobre também o
dump do corpus, que usa `find_tables`, introduzido na
[versão 1.23.0](https://pymupdf.io/changelog).

## Como usar no PC antigo

Baixe e extraia o repositório inteiro em qualquer pasta, preservando sua
estrutura. Abra `10_EXTRATOR/`. Os lançadores usam `py -3`, ajustam o diretório
para o local do próprio `.bat` e ativam o console UTF-8. Não há caminho fixo
de instalação do projeto ou do interpretador.

Execute primeiro `00`, `01` e `02`. Rode `04` para conferir os testes antes
de executar `03` sobre o banco completo.

| Lançador | O que faz e onde fica o resultado |
| --- | --- |
| `00_instalar_dependencias.bat` | Instala PyMuPDF, python-docx, openpyxl e pytest no Python de `py -3`. |
| `01_gerar_mapa.bat` | Gera `02_BANCO_ATUAL/mapa_normas_v1.json` e `mapa_normas_v1_revisao.xlsx`. Na primeira execução, cria o dump do corpus se ele faltar; essa etapa pode demorar. |
| `02_validar_40_normas.bat` | Imprime precisão/recall/F1 e gera `saida/resultado_40.xlsx` e `.json`. A referência atual é 100%/100%, com 558 pares. |
| `03_extrair_05_DOCUMENTS.bat` | Percorre `05_DOCUMENTS/` e todas as subpastas; gera `saida/resultado_05_DOCUMENTS.xlsx` e `.json`. Se a pasta faltar, mostra uma orientação e encerra. |
| `04_rodar_testes.bat` | Executa pytest em `02_BANCO_ATUAL/` e `10_EXTRATOR/`. |

Todas as saídas de extração ficam em `10_EXTRATOR/saida/`, ignorada pelo Git.
Confira `Ignorados`: variantes de idioma não operacionais, arquivos sem
identidade segura e erros de leitura ficam registrados com motivo e caminho
relativo à entrada. O lote continua com os demais arquivos. Com `--manifesto`,
erros de identidade ou leitura interrompem a execução para proteger a validação.

Quando sair um **Valid-Documents novo**, use a ferramenta antiga em
`99_CODIGOS_PY/cadastro_v3/`, seguindo as instruções dela, para regenerar o
cadastro. Confirme que o novo
`cadastro_canonico_v3_SEM_URL_SHAREPOINT.json` está em `02_BANCO_ATUAL/` deste
repositório. Depois execute `01_gerar_mapa.bat`, `02_validar_40_normas.bat` e
`04_rodar_testes.bat`; só então repita a extração completa. As ferramentas
antigas não foram modificadas.

## Execução

Na raiz do repositório:

```powershell
py -3 10_EXTRATOR/extrator_normas.py --entrada 03_PILOTO_40/corpus/normas --manifesto 03_PILOTO_40/corpus/corpus_40.json --saida 10_EXTRATOR/saida/resultado_40.xlsx --gabarito 03_PILOTO_40/corpus/gabarito_40_v1.json
py -3 -m pytest 10_EXTRATOR -q
```

Dentro de `10_EXTRATOR/`, os caminhos relativos à raiz também funcionam:

```powershell
py -3 extrator_normas.py --entrada 03_PILOTO_40/corpus/normas --manifesto 03_PILOTO_40/corpus/corpus_40.json --saida saida/resultado_40.xlsx --gabarito 03_PILOTO_40/corpus/gabarito_40_v1.json
```

`--entrada` aceita um arquivo ou os PDF/DOCX diretamente em uma pasta.
Use `--recursivo` para incluir subpastas, por exemplo:

```powershell
py -3 extrator_normas.py --entrada ..\05_DOCUMENTS --recursivo --saida saida\resultado_05_DOCUMENTS.xlsx
```

`--manifesto` e `--gabarito` são opcionais. `--mapa` permite selecionar outro
mapa; o padrão é `02_BANCO_ATUAL/mapa_normas_v1.json`.

Sem manifesto, a identidade do pai vem dos nomes de arquivo do cadastro v3,
das associações físicas aprovadas de F2 ou do `Document-No.` da capa resolvido
pelo mapa. Ausência de identidade vai para `Ignorados`; com `--manifesto`,
gera erro explícito. Variantes listadas em `language_variants` do cadastro
não são processadas quando há versão operacional equivalente.

## Regras

- Importa a normalização interna do gerador do mapa. Consulta as cadeias
  prontas; não infere sucessores nem incorpora decisões por documento.
- Vigentes e resolvidos usam o código vigente. Retirados e excluídos por
  decisão são descartados. Os demais estados preservam o código citado
  normalizado e exibem o status do mapa.
- Inclui normas públicas, famílias e partes, com identidade canônica sem
  edição, idioma ou qualificadores de método. Expande partes compactas e
  mantém separadas normas equivalentes citadas com barra.
- Regulatórios exigem contexto de requisito. Definições por extenso e listas
  de abreviaturas não bastam. A forma canônica é maiúscula, inclusive `ROHS`.
- Exclui clientes, fornecedores, processos e documentos regionais; preserva
  a exceção expressa `W-B-0718` e processos cadastrados como vigentes.
- Exclui autorreferências, histórico de revisão, códigos antigos em frases
  de substituição, metadados de Keywords, marcas `[WITHDRAWN]`/`Not in use`
  da mesma entrada, caminhos, curingas e classes de óleo SAE.
- Códigos atuais desconhecidos pelo mapa são falsos positivos. Uma pequena
  correção genérica trata `DIN 228`; outras trocas DIN/EN exigem a forma EN
  presente no mesmo documento. Não há condicionais por norma-pai ou página.

Cada descarte conserva o motivo e a evidência. A avaliação normaliza ambos
os lados pela mesma função, resolve internas, remove retiradas, exclusões e
autorreferências e compara pares exatos deduplicados. FP/FN incluem evidência
aceita, descartada ou a entrada do gabarito quando não houve candidato.

## Validação e limites

A regressão executa os 40 PDFs e exige precisão e recall de 100% contra o
gabarito **1.1**, recebido do orquestrador e copiado sem edição. Ele contém
576 inclusões brutas; após mapa, filtros de estado e deduplicação, são 558
pares. O teste não usa o gabarito para orientar a extração.

Não executa OCR nem `find_tables`. PDF sem texto extraível precisa de OCR
prévio. DOCX usa a página lógica 1, sem reproduzir a paginação do Word.
Cabeçalhos, entradas de tabela e Keywords são reconhecidos por heurísticas;
layouts diferentes do corpus podem exigir revisão dos descartes. Códigos
não confirmados no mapa continuam visíveis, sem promoção por semelhança.
Os resultados medidos se referem ao corpus de 40 documentos, não ao banco
inteiro. `saida/` é artefato local ignorado pelo Git.
