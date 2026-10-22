// Configuração pública da ligação ao Power Automate.
// Não contém segredos: o fluxo só aceita pedidos com um token Microsoft Entra
// de um utilizador do tenant SIR, e o próprio fluxo confirma o perfil em tbl_utilizadores_RH.
export const RH_CONFIG = Object.freeze({
  tenantId: '5fc7bb0c-4c3f-4008-ae03-9efd62fc96c9',
  // ID da aplicação "Sistema RH Web" registada no Microsoft Entra (preencher após o registo).
  clientId: '',
  // Fluxo RH_GET_Configuracoes ("Any user in my tenant"). O URL não dá acesso sem token.
  apiUrl: 'https://default5fc7bb0c4c3f4008ae039efd62fc96.c9.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/23/workflows/fe9c2afd0e9b4fb9957296931ba4a6df/triggers/manual/paths/invoke?api-version=1',
  // Audiência do serviço Power Automate (a barra dupla antes de .default é obrigatória).
  scopes: ['https://service.flow.microsoft.com//.default']
});
