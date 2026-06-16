# Alerta grave TTD409

O app cruza o NCM de cada item do XML com o Anexo Unico do Decreto SC 2.128/2009, que lista mercadorias importadas nao alcancadas por beneficios fiscais.

Quando houver correspondencia, o sistema exibe `ALERTA GRAVE TTD409`. Esses itens nao devem ser incluidos no TTD409 sem revisao fiscal da descricao legal, NCM e excecoes aplicaveis.

## Arquivos

- `data/ttd409_exclusions.csv`: base local de NCMs/prefixos do Anexo Unico.
- `data/ttd409_summary.json`: resumo da geracao.
- `tools/build_ttd409_database.py`: atualiza a base a partir da pagina oficial.

## Atualizacao

```powershell
python tools\build_ttd409_database.py
```

Ao abrir o app, `database.py` compara o hash do CSV e recarrega automaticamente a tabela `ttd409_exclusoes` no SQLite quando houver mudanca.

## Observacoes

- A regra automatica cruza NCM e prefixos de NCM. A decisao final depende tambem da descricao legal e da mercadoria importada.
- Itens 62 a 76 possuem regra condicionada ao uso na agricultura ou pecuaria, conforme alteracao de 2026.
- Exemplo validado: NCM `7009.10.00` cai no item 3, "Espelhos, classificados no codigo NCM 7009".
