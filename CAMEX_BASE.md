# Base CAMEX/Gecex por NCM

Base gerada em 2026-06-08 para alertar NCMs de XML NF-e que aparecam em listas oficiais CAMEX/Gecex e Siscomex.

## Arquivos gerados

- `data/camex_ncm_database.csv`: base consolidada usada pelo app.
- `data/camex_ncm_summary.json`: resumo de contagens e notas de geracao.
- `data/camex_sources/source_manifest.json`: URLs, hashes e tamanhos dos arquivos oficiais baixados.

## Fontes oficiais usadas

- Anexos II, III, IV, V, VI, VIII, IX e X da Resolucao Gecex 272/2021, arquivo de 01/06/2026.
- Lessin consolidada (`lessin.xlsx`).
- Cotas CAMEX vigentes, acompanhamento de cotas e modelos LPCO Decex Cotas, publicados no portal Siscomex.
- Ex-tarifarios BK/BIT vigentes publicados pelo MDIC. Em 2026-06-08, todos os registros desse arquivo tinham fim de vigencia em 2025-12-31; por isso ficam na base historica, mas nao disparam alerta vigente.
- Ex-tarifarios BK Autopropulsado vigentes.

## Exclusao intencional

O Anexo I - TEC foi excluido do alerta porque e a tarifa geral. Inclui-lo faria praticamente qualquer NCM valido aparecer como alerta CAMEX.

## Como atualizar

```powershell
python tools\build_camex_database.py
```

Ao abrir o app, `database.py` compara o hash do CSV e recarrega a tabela `camex_ncm_base` no SQLite quando houver mudanca.

## Regra de alerta

- NCM de 8 digitos: comparacao exata com o XML.
- NCM de 6 digitos: tratado como prefixo para NCMs de 8 digitos.
- EX: o app alerta pelo NCM e mostra o EX aplicavel quando existir; a conferencia final do enquadramento do EX deve ser feita pelo usuario.
