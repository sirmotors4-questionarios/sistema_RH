# Sistema RH — SIR Grupo

Aplicação Web para gestão de colaboradores, contratos, adendas, lançamentos salariais, férias e relatórios.

## Estado
**Configurações ligadas ao Excel no SharePoint através do Power Automate, com autenticação Microsoft.** Os restantes módulos (colaboradores, salários, férias…) continuam em modo demonstração com registos fictícios. Nenhuma informação salarial ou pessoal deve ser publicada neste repositório.

## Arquitectura escolhida
- Interface: HTML, CSS, JavaScript (GitHub Pages).
- Armazenamento: `Gestão de Colaboradores.xlsx` no SharePoint (SIR MOTORS INTERNATIONAL › Documentos › General › Recursos Humanos › Sistema › Ficheiro). O `.xlsm` com VBA continua a ser o sistema actual; o conector Excel do Power Automate só lê `.xlsx`/`.xlsb`.
- Camada de integração: Power Automate (fluxo `RH_GET_Configuracoes`, accionador HTTP restrito a "Any user in my tenant").
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

### Fluxo de consulta dos módulos (RH_API_Consultar)
`POST {"action": "funcionarios" | "contratos" | "lancamentos" | "ferias"}` com o mesmo token. O fluxo confirma o utilizador em `tbl_utilizadores_RH`, lê as tabelas do módulo (paginação até 5000 linhas) e responde `200 {ok, utilizador, action, tabelas:[{tabela, linhas}]}`, `403` (conta não autorizada) ou `400` (acção inválida). Só leitura.

Tabelas confirmadas no ficheiro Gestão de Colaboradores.xlsx (teste de 22/10/2026; contagens indicativas):

| Módulo | Tabela | Linhas | Colunas principais |
|---|---|---|---|
| funcionarios | tbl_funcionarios | 461 | Codigo, Nome, Departamento, Empresa, Estado, Profissao, Categoria, TipoContrato, Data Admissao, Vencimento Base, RegimePagamento, … (51 colunas) |
| contratos | tbl_adendas | 286 | ID_Adenda, Codigo, Nome, Numero_Adenda, Data_Adenda, Data_Inicio_Vigencia, Data_Fim_Vigencia, Campo_Alterado, Valor_Anterior, Valor_Novo, Estado_Adenda, Aplicada, … |
| contratos | tbl_parametros_pagamento | 4 | ID_Parametro, RegimePagamento, Categoria, Valor_Unitario, Unidade, Vigencia_Inicio, Vigencia_Fim, Activo |
| lancamentos | tbl_lançamento | 1805 | ID_Lancamento, Codigo, Nome, Ano, Mês, Dias Trabalhados, Horas Extras, Subsídios, Total de Acréscimos, Total de Descontos, IRPS, Salario Liquido, Estado do Lançamento, … |
| lancamentos | Table1 (escalões IRPS) | 10 | Limite Inferior, Limite Superior, 0, 1, 2, 3, 4 ou mais, Coeficiente adicional |
| lancamentos | tbl_arredondamento | 4 | Empresa, Casas_INSS, Casas_IRPS |
| ferias | tbl_ferias | 0 | (tabela vazia; colunas por confirmar quando houver registos) |
| ferias | tbl_feriados | 9 | Dia, Mes, Feriado |

Os nomes chegam exactamente como os cabeçalhos do Excel, **incluindo espaços e acentos** (ex.: `Vencimento Base`, `Data de Saída`, `Mês`). Cada linha traz também `ItemInternalId` e `@odata.etag`, que o site deve ignorar.

### Páginas dos módulos
`index.html` usa `rh-modulos.js` para mostrar os dados devolvidos pelo fluxo (só leitura): Colaboradores (pesquisa e filtro por estado), Contratos (parâmetros de pagamento vigentes e distribuição por tipo de contrato/regime), Adendas, Lançamento salarial (por mês e ano, com totais) e Férias/feriados. Sem sessão Microsoft, a página mantém os dados fictícios. Os botões de criação continuam desactivados até existir o fluxo de gravação com aprovação.

### Acesso por código pessoal (RH_API_Acesso) — sem login Microsoft
Fluxo **RH_API_Acesso** (gatilho HTTP "Anyone"): `POST {"action", "email", "chave"}`. Autoriza só se a linha de `tbl_utilizadores_RH` com esse e-mail estiver `Activo` e `Codigo_Hash` for igual a `chave`; devolve o mesmo formato do RH_API_Consultar (403 caso contrário, 400 para acção inválida).

- `chave = SHA-256(email em minúsculas + ":" + código)`, calculada no navegador; o código nunca é enviado.
- **O URL do fluxo é secreto** (contém `sig=`): não o coloque no repositório. Cada utilizador cola-o uma vez no diálogo «Entrar»; fica no `localStorage` desse navegador. Se for divulgado, regenere o URL no gatilho do fluxo.
- Registar um utilizador: a pessoa abre «Entrar» → «Primeiro acesso: gerar a chave», escreve o e-mail e o código (8+ caracteres) e envia a chave ao administrador, que a cola em `Codigo_Hash`. Para revogar, apague a chave ou mude `Estado`.
- Tempo de resposta medido (22/10/2026): férias ~5 s, colaboradores/contratos ~12 s, lançamentos ~50 s (1805 linhas).

### Verificação
Executar `node --test tests/rh-integration.test.mjs`. Os testes cobrem âmbito, vigência, fórmulas não validadas, parâmetros efectivos, IF, arredondamentos, calendário e bloqueio de escalões sobrepostos.


## Ligação ao Power Automate e autenticação
```
Browser (GitHub Pages) --MSAL / Microsoft Entra--> token (aud https://service.flow.microsoft.com/)
   └─ POST {"action":"configuracoes"} + Bearer token --> fluxo RH_GET_Configuracoes
        ├─ Utilizador  = e-mail do token (cabeçalhos x-ms-user-*-encoded)
        ├─ Ler_Utilizadores (tbl_utilizadores_RH, filtro Email eq …)
        ├─ Ler_Configuracoes / Ler_Escaloes / Ler_Formulas / Ler_Auditoria
        └─ Response: 200 {ok, utilizador, 4 tabelas} | 403 {ok:false, erro, email}
```
- **Sem token o fluxo responde 401** (`DirectApiAuthorizationRequired`). O URL em `rh-config.js` não contém assinatura nem segredo.
- **Permissões:** o fluxo só devolve dados a contas registadas e `Activo` em `tbl_utilizadores_RH` (folha UTILIZADORES_RH: ID_Utilizador, Nome, Email, Perfil, Empresa, Estado). Perfis: Superadministrador, Administrador RH, Técnico RH, Gestor. A interface usa o perfil apenas para orientar (ex.: desactivar "Propor"); as futuras gravações terão de validar o perfil no fluxo.
- **Registo no Entra:** aplicação "Sistema RH Web" (single-tenant, plataforma SPA) com os redirect URIs das páginas (`…/sistema_RH/`, `…/sistema_RH/index.html`, `…/sistema_RH/configuracoes.html` e `http://localhost:8000/configuracoes.html` para testes) e a permissão delegada **Power Automate › User**. O `clientId` é público e vai em `rh-config.js`.
- MSAL 4.30.0 (LTS) é carregado de jsDelivr com `integrity` (SRI). A sessão fica em `sessionStorage`.

### Verificação
`node --test tests/rh-integration.test.mjs tests/rh-api.test.mjs tests/rh-modulos.test.mjs tests/rh-acesso.test.mjs`
