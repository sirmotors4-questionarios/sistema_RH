# Office Scripts — lançamentos salariais

## Desempenho

O Power Automate dá 120 s a cada execução do script. Cada leitura ao Excel é uma ida ao servidor (as escritas
seguem em lote), por isso o script não faz leituras por colaborador: as linhas do mês são lidas num só bloco,
os escalões são calculados de uma vez, o livro é recalculado uma única vez e linhas seguidas são apagadas num só
pedido. O teste `recalcular 200 Pendentes faz um número fixo de leituras` garante que isto se mantém
(versão anterior: ~4 leituras por Pendente, 805 para 200; actual: 9, independentemente do número).

## Activar retirar/introduzir no «Carregar mês» (v3) — OBRIGATÓRIO antes de usar

1. No Excel/Microsoft 365, abra o script **LancamentoNovoMes** usado pelo fluxo **RH_API_Actualizar**,
   substitua o conteúdo pelo ficheiro [LancamentoNovoMes.ts](LancamentoNovoMes.ts) desta versão e guarde.
2. Não é preciso mexer no fluxo: as operações novas (`retirar_pendentes`, `introduzir_lancamentos`) passam
   pelo ramo **lancamento_editar** já existente, como as da v2.
3. Valide numa cópia do workbook: um mês com um Pendente sem contrato (deve ser apagado), um Alterado sem
   contrato (não deve ser apagado) e um colaborador activo sem lançamento (deve ser inserido como Pendente,
   com as fórmulas nas colunas calculadas).

Enquanto o script não for actualizado, o site mostra «O script do Excel ainda não sabe retirar/introduzir
lançamentos» e não grava nada dessas operações. «Só recalcular Pendentes» continua a funcionar.

## Activar recálculo de Pendentes e reposição (v2)

1. No Excel/Microsoft 365, abra o script **LancamentoNovoMes** efectivamente usado pela
   acção Run script do fluxo **RH_API_Actualizar**. Substitua o conteúdo pelo ficheiro
   [LancamentoNovoMes.ts](LancamentoNovoMes.ts) desta versão e guarde.
2. No ramo **lancamento_editar**, confirme que o script recebe `modo = editar`,
   `ano = triggerBody()?['ano']`, `mes = triggerBody()?['mesNome']` e
   `linhas = string(triggerBody()?['linhas'])`, sem filtrar propriedades das linhas.
   A resposta deve preservar **ok, erros e linha** do resultado do script.
   Não são necessárias acções HTTP adicionais, nem novos parâmetros do script.
3. Valide numa cópia do workbook: carregar um mês com Pendente, Alterado e Finalizado;
   só o Pendente deve mudar de condições, mantendo Pendente. Edite-o, grave, reabra
   e reponha: deve recuperar as entradas anteriores. Tente repor um Finalizado e um
   Alterado sem referência: ambos devem ser recusados.
4. Depois da validação, integre a interface. A versão antiga do script rejeita os novos
   envelopes antes de escrever. A interface exige a confirmação v2 e não apresenta
   sucesso com uma resposta antiga/incompleta.

A nova tabela **tbl_lancamentos_iniciais_RH** é criada pelo script numa folha adicional,
com `Chave`, `Referencia_JSON` e `Guardado_Em`. Não a apague: contém as referências para
reposição. As permissões são as do workbook; não é exposta pela API de consulta.
As cópias de recuperação do workbook devem incluir esta tabela.

O recálculo mantém Pendente, actualiza apenas condições de pagamento e auditoria, e
preserva faltas, horas, subsídios, bónus, descontos e observações. A reposição recupera
as entradas e condições da referência, marca Pendente / Revisto Não e actualiza auditoria.
Não escreve nas colunas de fórmulas. Os valores calculados reflectem as fórmulas actuais.
Lançamentos já Alterados antes desta instalação não têm referência recuperável.

O fluxo deve continuar a autorizar apenas Superadministrador/Administrador RH e a
serializar escritas ao workbook (Concurrency control = 1). O script revalida o estado
no Excel; uma linha que passou de Pendente para Alterado/Finalizado é ignorada no lote.
A validação de todas as linhas precede as escritas; falhas de serviço durante a escrita
não têm garantia transaccional e exigem reler o mês antes de repetir.

## Referência anterior de criação de mês

`LancamentoNovoMes.ts` grava as linhas de um novo mês em `tbl_lançamento` numa única chamada,
em substituição do `Apply_to_each → Adicionar_Lancamento` (uma chamada por linha) do fluxo
`RH_API_Actualizar`. A ligação efectiva deve ser confirmada no ambiente Microsoft.

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
