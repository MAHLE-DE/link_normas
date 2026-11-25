# Extrator único de referências normativas

`extrator_normas.py` lê PDF (PyMuPDF, texto unido por página) e DOCX
(parágrafos e tabelas na ordem original). Produz pares únicos de norma-pai e
norma citada, preservando código bruto, páginas, seção, arquivo e evidência.
O arquivo XLSX tem `Correlacoes`, `Descartes` e, com gabarito, `Avaliacao`.
Um JSON irmão contém os mesmos registros e métricas.

Requer Python 3.13 com PyMuPDF, python-docx e openpyxl. Os testes usam pytest.

## Execução

Na raiz do repositório:

```powershell
python 10_EXTRATOR/extrator_normas.py --entrada 03_PILOTO_40/corpus/normas --manifesto 03_PILOTO_40/corpus/corpus_40.json --saida 10_EXTRATOR/saida/resultado_40.xlsx --gabarito 03_PILOTO_40/corpus/gabarito_40_v1.json
python -m pytest 10_EXTRATOR -q
```

Dentro de `10_EXTRATOR/`, os caminhos relativos à raiz também funcionam:

```powershell
python extrator_normas.py --entrada 03_PILOTO_40/corpus/normas --manifesto 03_PILOTO_40/corpus/corpus_40.json --saida saida/resultado_40.xlsx --gabarito 03_PILOTO_40/corpus/gabarito_40_v1.json
```

`--entrada` aceita um arquivo ou os PDF/DOCX diretamente em uma pasta.
`--manifesto` e `--gabarito` são opcionais. `--mapa` permite selecionar outro
mapa; o padrão é `02_BANCO_ATUAL/mapa_normas_v1.json`.

Sem manifesto, a identidade do pai vem dos nomes de arquivo do cadastro v3,
das associações físicas aprovadas de F2 ou do `Document-No.` da capa resolvido
pelo mapa. Ausência de identidade gera erro explícito.

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
