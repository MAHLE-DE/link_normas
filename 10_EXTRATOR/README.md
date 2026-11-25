# Extrator único de referências normativas

`extrator_normas.py` lê PDF (PyMuPDF, texto unido por página) e DOCX
(parágrafos e tabelas na ordem original). Produz pares únicos de norma-pai e
norma citada, preservando código bruto, páginas, seção, arquivo e evidência.
O arquivo XLSX tem `Correlacoes`, `Descartes`, `Ignorados`, `Avisos` e, com gabarito,
`Avaliacao`.
Um JSON irmão contém os mesmos registros e métricas.

Listas imediatamente regidas por verbos ativos de substituição (`replaces`,
`supersedes`, `ersetzt`, `Ersatz für` etc.) são históricas até o fim da frase
ou da entrada. Formas passivas preservam o código sucessor. No mapa, uma
substituição só cria identidade quando há um único documento declarante;
múltiplos declarantes ficam em Avisos. Legados sem sucessor com cadeia usam
o destino final, tanto na extração quanto na avaliação, para deduplicar aliases.

Sem manifesto, a identidade prioriza o arquivo do cadastro Valid-Documents,
depois identifica variantes não operacionais, consulta associações aprovadas
de arquivo e, por fim, a capa. `Avisos` informa associações aprovadas que
colidem com outro documento do cadastro e normas processadas a partir de
mais de um arquivo; estes arquivos continuam sendo agregados.

Para atualizar a base completa no PC antigo, execute `01_gerar_mapa.bat`
e depois `03_extrair_05_DOCUMENTS.bat`. O primeiro lê `05_DOCUMENTS/`
recursivamente quando ela existe, incluindo variantes de idioma como evidência,
e grava o mapa local em `02_BANCO_ATUAL/mapa_base/`. Sem essa pasta de documentos,
avisa e usa somente fontes do repositório. O extrator prefere esse mapa local
quando disponível e informa o mapa escolhido; `--mapa` tem prioridade explícita.
O mapa versionado permanece gerado sem base (`--sem-base`).

O gerador aceita `--base-documentos <pasta>` e `--saida <pasta>`; impede
sobrescrever o mapa versionado quando usa base. `--sem-cache` desliga o cache
de texto por SHA256 em `02_BANCO_ATUAL/cache_texto/`. Cache e `mapa_base/`
são ignorados pelo Git por dependerem dos arquivos locais do usuário.

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
| `01_gerar_mapa.bat` | Gera JSON e XLSX em `02_BANCO_ATUAL/mapa_base/`, usando `05_DOCUMENTS/` quando existir. Preserva o mapa versionado. Na primeira execução, cria o dump do corpus se faltar. |
| `02_validar_40_normas.bat` | Imprime precisão/recall/F1 e gera `saida/resultado_40.xlsx` e `.json`. A referência atual é 100%/100%, com 551 pares após as decisões F9. |
| `03_extrair_05_DOCUMENTS.bat` | Percorre `05_DOCUMENTS/` e todas as subpastas; gera `saida/resultado_05_DOCUMENTS.xlsx` e `.json`. Se a pasta faltar, mostra uma orientação e encerra. |
| `04_rodar_testes.bat` | Executa pytest em `02_BANCO_ATUAL/` e `10_EXTRATOR/`. |
| `05_validar_segunda_amostra.bat` | Valida a segunda amostra contra o gabarito v2.1; gera `saida/resultado_v2.xlsx` e `.json`. |
| `06_validar_terceira_amostra.bat` | Valida a terceira amostra contra o gabarito v3.1; gera `saida/resultado_v3.xlsx` e `.json`. |
| `07_gerar_pendencias.bat` | Cria `02_BANCO_ATUAL/pendencias_codigos_sem_cadeia.xlsx` a partir do resultado da base completa. |
| `08_aplicar_pendencias.bat` | Aplica as decisões preenchidas na planilha padrão, acrescentando entradas F9 e preservando decisões já existentes. |

As avaliações cegas, antes de corrigir regras gerais, mediram precisão
90,80% / recall 76,70% na segunda amostra e precisão 96,23% / recall 98,55%
na terceira. As três amostras têm testes de regressão exigindo 100%/100%.

## Revisar códigos sem cadeia

Depois de extrair a base completa, execute `07_gerar_pendencias.bat`.
A aba `Pendências` ordena os códigos por número de relações e mostra
normas-pai/títulos, até três contextos, dados BU2 e candidatos com indícios
e similaridade. Preencha somente `Decisão`, `Código destino` e `Observação`.
As ações são `ligar`, `sem_sucessor`, `nao_considerar` e `retirado`.
Para `ligar`, o destino precisa existir no mapa ou ser cadastrado também
como decisão na planilha; as outras ações deixam o destino vazio.

Guarde uma cópia preenchida antes de gerar novamente: a geração recria as
colunas de decisão vazias. Execute `08_aplicar_pendencias.bat` para a planilha
padrão, ou `py -3 ..\02_BANCO_ATUAL\aplicar_pendencias.py --planilha <copia.xlsx>`.
O aplicador valida tudo antes de gravar, ignora decisões repetidas e avisa
quando uma decisão existente diferente foi preservada. Campos vazios adiam
a decisão. Em seguida, rode `01_gerar_mapa.bat` e `03_extrair_05_DOCUMENTS.bat`.
Valid-Documents/F1 sempre prevalece.

O gerador aceita `--resultado`, `--mapa` e `--saida`; por padrão usa o resultado
da base completa e `mapa_base/`, com fallback ao mapa versionado. Decisões F9
mais recentes são conferidas no mapa versionado atualizado para retirar da
fila casos já decididos, como W-B-0718. O aplicador aceita `--planilha`,
`--mapa` e `--decisoes` para usar cópias em testes ou revisão.

O dump usa por padrão texto PDF sem `find_tables`, que pode encerrar o
processo nativamente. O texto das tabelas permanece no texto extraído;
tabelas DOCX continuam sendo emitidas. O extrator não chama `find_tables`.

Todas as saídas de extração ficam em `10_EXTRATOR/saida/`, ignorada pelo Git.
Confira `Ignorados`: variantes de idioma não operacionais, arquivos sem
identidade segura e erros de leitura ficam registrados com motivo e caminho
relativo à entrada. O lote continua com os demais arquivos. Com `--manifesto`,
erros de identidade ou leitura interrompem a execução para proteger a validação.
Com manifesto, todos os arquivos explicitamente listados são processados,
inclusive variantes de idioma. O filtro de variantes não operacionais vale
somente para lotes sem manifesto.

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
mapa; o padrão prefere `02_BANCO_ATUAL/mapa_base/mapa_normas_v1.json`, se existir,
e usa `02_BANCO_ATUAL/mapa_normas_v1.json` caso contrário.

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
576 inclusões brutas; após mapa, filtros de estado e deduplicação, são 551
pares com as decisões F9 atuais. Esse total depende do mapa vigente.
O teste não usa o gabarito para orientar a extração.

Não executa OCR nem `find_tables`. PDF sem texto extraível precisa de OCR
prévio. DOCX usa a página lógica 1, sem reproduzir a paginação do Word.
Cabeçalhos, entradas de tabela e Keywords são reconhecidos por heurísticas;
layouts diferentes do corpus podem exigir revisão dos descartes. Códigos
não confirmados no mapa continuam visíveis, sem promoção por semelhança.
Os resultados medidos se referem ao corpus de 40 documentos, não ao banco
inteiro. `saida/` é artefato local ignorado pelo Git.
