# Segunda amostra — holdout

`corpus_40_v2.json` identifica os 40 PDFs mantidos em `../normas_v2/`.
A autoridade de identidade é o cadastro canônico v3; `idioma` descreve o
arquivo da amostra, inclusive variantes alemãs. A referência de seleção
original é `../amostra_validacao_40.json`.

Para gerar somente os dumps, a partir da raiz do repositório:

```powershell
py -3 03_PILOTO_40/corpus/dump_corpus.py --manifesto 04_NOVO_TESTE_40/corpus/corpus_40_v2.json --saida 04_NOVO_TESTE_40/corpus/dump
```

Os caminhos dos PDFs são relativos ao manifesto. O dump inclui todas as
línguas; não aplica o filtro operacional do extrator. `dump/` não é versionado.

Esta amostra está reservada para avaliação cega. Nesta preparação, executou-se
apenas o dump de estudo. Não executar o extrator nem ajustar suas regras
usando estes documentos antes da elaboração independente do gabarito v2.
