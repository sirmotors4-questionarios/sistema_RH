import {configValues, resolveConfig, resolveFormula, evaluateFormula} from './rh-engine.js';

export const TABLES = ['tbl_configuracoes_RH', 'tbl_escaloes_RH', 'tbl_formulas_RH', 'tbl_auditoria_RH'];

// Contrato partilhado pela importação local e pela futura resposta autenticada da API.
export function readTables(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw Error('Resposta de tabelas inválida.');
  const tables = {};
  for (const key of TABLES) {
    if (!Array.isArray(payload[key]) || payload[key].some(row => !row || typeof row !== 'object' || Array.isArray(row))) {
      throw Error('A tabela ' + key + ' deve ser uma lista de registos.');
    }
    tables[key] = payload[key].map(row => ({...row}));
  }
  return tables;
}

export function resolveTables(tables, context, at) {
  const configs = tables.tbl_configuracoes_RH;
  const parameters = [...new Set(configs.map(r => r.Codigo_Parametro).filter(Boolean))]
    .map(key => resolveConfig(configs, key, context, at)).filter(Boolean);
  const formulas = [], unavailable = [];
  for (const key of new Set(tables.tbl_formulas_RH.map(r => r.Codigo_Formula).filter(Boolean))) {
    try {
      const row = resolveFormula(tables.tbl_formulas_RH, key, context, at);
      if (row) formulas.push(row);
    } catch (error) {
      if (error.code !== 'RH_UNVALIDATED') throw error;
      unavailable.push(key);
    }
  }
  return {parameters, formulas, unavailable, values: configValues(configs, context, at)};
}

export function simulateRecord(tables, code, context, at, inputs) {
  const formula = resolveFormula(tables.tbl_formulas_RH, code, context, at);
  if (!formula) throw Error('Nenhuma fórmula vigente para este contexto.');
  const values = configValues(tables.tbl_configuracoes_RH, context, at);
  // Os parâmetros aprovados prevalecem sobre valores digitados na simulação.
  const variables = Object.fromEntries(Object.entries(inputs).map(([key, value]) => [key.toUpperCase(), value]));
  Object.assign(variables, Object.fromEntries(Object.entries(values).map(([key, value]) => [key.toUpperCase(), value])));
  const result = evaluateFormula(formula.Expressao, variables);
  if (typeof result !== 'number' || !Number.isFinite(result)) throw Error('A fórmula salarial deve devolver um número finito.');
  return {result, formula};
}
