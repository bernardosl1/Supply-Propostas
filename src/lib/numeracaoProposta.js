function resolverNumeroDocumento(numeroSolicitado, propostas = [], idAtual = '') {
  const solicitado = String(numeroSolicitado || '').trim();
  if (!solicitado) return '';

  const propostaAtual = (Array.isArray(propostas) ? propostas : [])
    .find((proposta) => idAtual && String(proposta?.id || '') === String(idAtual));
  const numeroAtual = obterNumeroDocumento(propostaAtual);
  if (numeroAtual && normalizarNumero(numeroAtual) === normalizarNumero(solicitado)) {
    return solicitado;
  }

  const solicitadoSeparado = separarNumeroDocumento(solicitado);
  const revisoesExistentes = (Array.isArray(propostas) ? propostas : [])
    .filter((proposta) => !idAtual || String(proposta?.id || '') !== String(idAtual))
    .map((proposta) => separarNumeroDocumento(obterNumeroDocumento(proposta)))
    .filter((numero) => numero.baseNormalizada === solicitadoSeparado.baseNormalizada)
    .map((numero) => numero.revisao || 0);

  if (!revisoesExistentes.length) {
    return solicitadoSeparado.numeroSemCopia;
  }

  const proximaRevisao = Math.max(...revisoesExistentes) + 1;
  return `${solicitadoSeparado.base}-R${String(proximaRevisao).padStart(2, '0')}`;
}

function separarNumeroDocumento(value) {
  const original = String(value || '').trim();
  const numeroSemCopia = original.replace(/(?:-COPIA)+$/i, '').trim();
  const match = numeroSemCopia.match(/^(.*?)(?:-R(\d+))?$/i);
  const base = String(match?.[1] || numeroSemCopia).trim();
  return {
    base,
    baseNormalizada: normalizarNumero(base),
    revisao: match?.[2] ? Number.parseInt(match[2], 10) : null,
    numeroSemCopia
  };
}

function obterNumeroDocumento(proposta) {
  return String(proposta?.numero_documento || proposta?.data?.numero_documento || '').trim();
}

function normalizarNumero(value) {
  return String(value || '').trim().toLocaleUpperCase('pt-BR');
}

module.exports = {
  resolverNumeroDocumento,
  separarNumeroDocumento
};
