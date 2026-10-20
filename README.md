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
Abra `index.html` localmente, ou active GitHub Pages apontando para a raiz do ramo `main`. Todo o conteúdo é demonstrativo até haver ligação segura aprovada.
