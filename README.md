# Sistema RH — SIR Grupo

Aplicação Web para gestão de colaboradores, contratos, adendas, lançamentos salariais, férias e relatórios.

## Estado
**Protótipo de interface sem dados reais.** Não está conectado ao Power Automate, Excel no SharePoint ou autenticação. O modo demonstração contém apenas registos fictícios. Nenhuma informação salarial ou pessoal deve ser publicada neste repositório.

## Arquitectura escolhida
- Interface: HTML, CSS, JavaScript (GitHub Pages).
- Armazenamento: ficheiro Excel `Gestão de Colaoradores.xlsm` existente no SharePoint.
- Camada de integração: Power Automate; os fluxos e respectivas URLs ainda têm de ser configurados.
- Autorização: requer mecanismo autenticado e um intermediário seguro. **Não publicar URLs de fluxos com segredo, tokens ou credenciais no código do GitHub Pages.**
- Lógica: migração progressiva das regras VBA, validando equivalência de resultados antes de activar gravações.

## Tabelas oficiais
`tbl_funcionarios`, `tbl_lançamento`, `tbl_adendas`, `tbl_ferias`, `tbl_parametros_pagamento`, `tbl_arredondamento`, `tbl_feriados`, `Table1` (IRPS).

As tabelas `tbl_lançamento9`, `tbl_lançamento12` e `Table9` estão fora do âmbito.

## Primeira fase
- Navegação responsiva, dashboard e módulos de RH em modo demonstração.
- Validações UX, pesquisa e filtros sem gravação.
- Contrato de API documentado para futura ligação segura.
- Para entrar em produção: validar localização e permissões do Excel, concorrência de escrita, gestão de erros, autenticação, regras de acesso e parâmetros dos fluxos.

## Executar
Sirva a pasta por HTTP (por exemplo, `python -m http.server 8000`) e abra `http://localhost:8000`, ou utilize GitHub Pages na raiz de `main`. A página de Configurações utiliza módulos JavaScript; abrir por `file://` não é suportado.


## Motor ligado às Configurações
`configuracoes.html` carrega `configuracoes.js` como módulo. O simulador usa `rh-engine.js`; `rh-data.js` liga os registos das quatro tabelas à selecção de regras e à simulação.

- Sem importação: simulação demonstrativa com os campos editáveis.
- Com importação JSON: regras seleccionadas por Colaborador → Categoria → Empresa → Geral, estado e data de referência. Os parâmetros resolvidos podem ser consultados na página. Parâmetros importados prevalecem sobre entradas manuais com o mesmo código.
- Fórmulas não validadas não ficam disponíveis. O estado Validada do ficheiro não certifica equivalência com VBA. Cada fórmula é simulada individualmente; dependências entre fórmulas não são calculadas automaticamente.
- Auditoria importada é apenas consultada. Propostas, aprovação, gravação, processamento salarial completo e leitura directa do Excel continuam pendentes.
- IF avalia apenas o ramo escolhido. NETWORKDAYS usa segunda a sexta e feriados explicitamente fornecidos; não suporta outros calendários de trabalho nesta versão.

### Contrato das quatro tabelas
A importação local aceita um objecto JSON com estas quatro propriedades, cada uma contendo um array de objectos com os nomes de colunas do Excel:

```json
{
  "tbl_configuracoes_RH": [],
  "tbl_escaloes_RH": [],
  "tbl_formulas_RH": [],
  "tbl_auditoria_RH": []
}
```

As datas podem ser ISO ou números de série Excel. Valores de taxa como `0.03` são usados directamente. As variáveis de fórmula devem conservar os códigos das tabelas; variáveis adicionais são introduzidas no simulador, sem nomes inventados ou conversões implícitas. BASE e LIQUIDO são entradas manuais identificadas como simulação. Não existe selecção automática de escalões IRPS nem certificação fiscal nesta ligação.

O ficheiro JSON é lido em memória no navegador, até recarregar a página. Não publique dados reais no GitHub. A futura API autenticada poderá entregar o mesmo objecto a `readTables`; nenhum endpoint Power Automate foi configurado nesta alteração.

### Verificação
Executar `node --test tests/rh-integration.test.mjs`. Os testes cobrem âmbito, vigência, fórmulas não validadas, parâmetros efectivos, IF, arredondamentos, calendário e bloqueio de escalões sobrepostos.
