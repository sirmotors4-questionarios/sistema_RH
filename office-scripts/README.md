# Office Scripts — gravação em lote (rascunho)

`LancamentoNovoMes.ts` grava as linhas de um novo mês em `tbl_lançamento` numa única chamada,
em substituição do `Apply_to_each → Adicionar_Lancamento` (uma chamada por linha) do fluxo
`RH_API_Actualizar`. **Ainda não está ligado a nenhum fluxo.**

## Parâmetros (acção "Run script from SharePoint library")

| Parâmetro | Tipo | Exemplo |
|---|---|---|
| `ano` | número | `2026` |
| `mes` | texto (igual à coluna `Mês`) | `Março` |
| `linhas` | texto JSON: `[{"Empresa":"…","Codigo":"…", …}]` | `@{string(triggerBody()?['linhas'])}` |
| `modo` | `validar` ou `gravar` | `validar` |

Devolve `{ok, modo, recebidas, inseridas, duplicadasNoPedido, jaExistentes, erros, colunasFormula, ms}`.

## Regras garantidas

- **Fórmulas:** as colunas com fórmula são detectadas na 1.ª linha de dados; nas linhas novas é
  reposta a mesma fórmula (R1C1). Um pedido que tente escrever numa delas é rejeitado.
- **Chave:** `Empresa | Código (4 dígitos) | Ano | Mês`, normalizada como no site (`chaveRH`).
  Duplicados no pedido ou já existentes na tabela → nada é escrito.
- **Tudo ou nada na validação:** qualquer erro impede a escrita de todas as linhas.
- **Concorrência:** o script não tem bloqueios; o fluxo que o chama deve ter
  *Concurrency control = 1* no gatilho, e o script volta a verificar as chaves no momento da escrita.

## Plano de testes (antes de produção)

1. Fazer uma **cópia** de `Gestão de Colaboradores.xlsx` (ex.: `… - TESTE.xlsx`) na mesma biblioteca.
2. Guardar o script na biblioteca SharePoint da equipa (não no OneDrive pessoal).
3. Fluxo de teste → `modo = validar` com o mês seguinte; confirmar `ok` e `colunasFormula` (11 esperadas).
4. `modo = gravar` na cópia; confirmar contagem de linhas, fórmulas nas novas linhas e totais na folha.
5. Repetir o mesmo pedido → deve devolver `jaExistentes` e não escrever.
6. Medir o tempo (`ms`) e comparar com o Apply-to-each.
7. Só depois, com cópia de recuperação do `RH_API_Actualizar`, trocar o ramo `Apply_to_each`.

A lógica foi verificada fora do Excel com uma tabela simulada (validar, gravar, repetir, duplicados,
escrita em coluna com fórmula, coluna desconhecida). Falta o teste real na cópia do ficheiro.
