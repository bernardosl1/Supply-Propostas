const VALID_BLOCK_TYPES = new Set(['texto', 'lista', 'preco', 'tabela', 'quebra_pagina']);
const VALID_FIXED_SECTIONS = new Set(['dados_comerciais', 'objeto', 'escopo']);

function normalizarEstruturaModelo(input = {}) {
  const blocks = Array.isArray(input.blocos_adicionais)
    ? input.blocos_adicionais.map(normalizeBlock).filter(Boolean)
    : [];
  const validIds = new Set([
    ...VALID_FIXED_SECTIONS,
    ...blocks.map((block) => `flex:${block.id}`)
  ]);
  const order = [];
  (Array.isArray(input.ordem_secoes) ? input.ordem_secoes : []).forEach((sectionId) => {
    const normalized = cleanText(sectionId, 180);
    if (validIds.has(normalized) && !order.includes(normalized)) order.push(normalized);
  });
  ['dados_comerciais', 'objeto', 'escopo'].forEach((sectionId) => {
    if (!order.includes(sectionId)) order.push(sectionId);
  });
  blocks.forEach((block) => {
    const sectionId = `flex:${block.id}`;
    if (!order.includes(sectionId)) order.push(sectionId);
  });

  return {
    versao: 1,
    secoes_excluidas: (Array.isArray(input.secoes_excluidas) ? input.secoes_excluidas : [])
      .map((value) => cleanText(value, 40))
      .filter((value, index, values) => ['objeto', 'escopo'].includes(value) && values.indexOf(value) === index),
    ordem_secoes: order,
    blocos_adicionais: blocks
  };
}

function normalizeBlock(block, index) {
  const type = cleanText(block?.tipo, 30);
  if (!VALID_BLOCK_TYPES.has(type)) return null;
  const id = cleanText(block?.id, 100) || `modelo-${type}-${index + 1}`;
  const normalized = { id, tipo: type };

  if (type === 'quebra_pagina') return normalized;
  if (type === 'texto' || type === 'tabela') normalized.titulo = cleanText(block.titulo, 180);

  if (type === 'texto') {
    normalized.observacoes = [];
    normalized.subtopicos = (Array.isArray(block.subtopicos) ? block.subtopicos : [])
      .map(normalizeSubtopic)
      .filter(Boolean);
  }
  if (type === 'lista') normalized.linhas = [];
  if (type === 'tabela') {
    normalized.colunas = (Array.isArray(block.colunas) ? block.colunas : [])
      .map((column, columnIndex) => ({
        id: cleanText(column?.id, 100) || `${id}-coluna-${columnIndex + 1}`,
        nome: cleanText(column?.nome, 120)
      }));
    normalized.linhas = [];
  }
  if (type === 'preco') {
    normalized.topicos_preco = (Array.isArray(block.topicos_preco) ? block.topicos_preco : [])
      .map((topic) => ({
        titulo: cleanText(topic?.titulo, 180),
        tipo: cleanText(topic?.tipo, 60) || 'personalizado',
        itens: []
      }));
  }
  return normalized;
}

function normalizeSubtopic(subtopic, index) {
  if (!subtopic || typeof subtopic !== 'object') return null;
  return {
    id: cleanText(subtopic.id, 100) || `modelo-subtopico-${index + 1}`,
    titulo: cleanText(subtopic.titulo, 180),
    observacoes: [],
    subtopicos: (Array.isArray(subtopic.subtopicos) ? subtopic.subtopicos : [])
      .map(normalizeSubtopic)
      .filter(Boolean)
  };
}

function cleanText(value, maxLength) {
  return String(value || '').trim().slice(0, maxLength);
}

module.exports = { normalizarEstruturaModelo };
