// Percentagem de etapas confirmadas; nunca estima o avanço dentro do Excel.
export function criarProgresso(host, titulo, total) {
  const box = document.createElement('div');
  box.className = 'rh-progresso';
  const estado = document.createElement('p');
  estado.setAttribute('role', 'status');
  estado.setAttribute('aria-live', 'polite');
  const barra = document.createElement('progress');
  barra.max = 100;
  barra.setAttribute('aria-label', titulo);
  const ajuda = document.createElement('small');
  ajuda.textContent = 'Progresso por etapas concluídas. Durante o processamento no Excel, aguarde a confirmação.';
  box.append(estado, barra, ajuda);
  host.replaceChildren(box);
  let encerrado = false;
  function actualizar(feitas, mensagem) {
    if (encerrado) return;
    const percentagem = Math.min(100, Math.max(0, Math.floor(feitas / total * 100)));
    barra.value = percentagem;
    estado.textContent = `${titulo} · ${percentagem}% — ${mensagem}`;
  }
  actualizar(0, 'A iniciar…');
  return {
    actualizar,
    concluir(mensagem = 'Concluído.') { actualizar(total, mensagem); encerrado = true; },
    falhar(mensagem) { encerrado = true; box.classList.add('rh-progresso-erro'); estado.textContent = `${titulo} · ${barra.value}% — ${mensagem}`; }
  };
}
