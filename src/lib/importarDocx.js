const fs = require('node:fs');
const PizZip = require('pizzip');
const { DOMParser } = require('@xmldom/xmldom');

const OPENAI_RESPONSES_URL = 'https://api.openai.com/v1/responses';
const OPENAI_MODEL = 'gpt-5.6-luna';
const MAX_DOCX_BYTES = 10 * 1024 * 1024;
const MAX_EXTRACTED_CHARS = 900_000;

const ITEM_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['item', 'descricao', 'ncm', 'quant', 'un', 'valor_unit', 'valor_total'],
  properties: {
    item: { type: 'string' },
    descricao: { type: 'string' },
    ncm: { type: 'string' },
    quant: { type: 'number' },
    un: { type: 'string' },
    valor_unit: { type: 'number' },
    valor_total: { type: 'number' }
  }
};

const PRICE_TOPIC_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['tipo', 'titulo', 'itens'],
  properties: {
    tipo: { type: 'string', enum: ['servico', 'consumivel', 'personalizado'] },
    titulo: { type: 'string' },
    itens: { type: 'array', items: ITEM_SCHEMA }
  }
};

const LEAF_SUBTOPIC_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['titulo', 'observacoes'],
  properties: {
    titulo: { type: 'string' },
    observacoes: { type: 'array', items: { type: 'string' } }
  }
};

const SUBTOPIC_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['titulo', 'observacoes', 'subtopicos'],
  properties: {
    titulo: { type: 'string' },
    observacoes: { type: 'array', items: { type: 'string' } },
    subtopicos: { type: 'array', items: LEAF_SUBTOPIC_SCHEMA }
  }
};

const ADDITIONAL_BLOCK_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: [
    'tipo', 'titulo', 'observacoes', 'subtopicos', 'linhas_documentacao',
    'topicos_preco', 'preco_total_numero', 'preco_total_extenso', 'moeda',
    'validade_proposta', 'pagamento', 'prazo_entrega', 'frete', 'impostos',
    'colunas_tabela', 'linhas_tabela'
  ],
  properties: {
    tipo: { type: 'string', enum: ['texto', 'lista', 'preco', 'tabela', 'quebra_pagina'] },
    titulo: { type: 'string' },
    observacoes: { type: 'array', items: { type: 'string' } },
    subtopicos: { type: 'array', items: SUBTOPIC_SCHEMA },
    linhas_documentacao: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['descricao', 'numero_documento', 'data'],
        properties: {
          descricao: { type: 'string' },
          numero_documento: { type: 'string' },
          data: { type: 'string' }
        }
      }
    },
    topicos_preco: { type: 'array', items: PRICE_TOPIC_SCHEMA },
    preco_total_numero: { type: 'number' },
    preco_total_extenso: { type: 'string' },
    moeda: { type: 'string' },
    validade_proposta: { type: 'string' },
    pagamento: { type: 'string' },
    prazo_entrega: { type: 'string' },
    frete: { type: 'string' },
    impostos: { type: 'string' },
    colunas_tabela: { type: 'array', items: { type: 'string' } },
    linhas_tabela: {
      type: 'array',
      items: { type: 'array', items: { type: 'string' } }
    }
  }
};

const RESPONSE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: [
    'empresa_cliente', 'unidade', 'processo_fluig', 'solicitante_nome_cargo',
    'contato_email', 'contato_telefone', 'numero_documento', 'data_documento',
    'responsavel_nome', 'responsavel_email', 'responsavel_telefone', 'objeto',
    'servicos_descricao', 'equipe_tecnica_itens', 'local_servico', 'data_servico',
    'prazo_execucao_dias', 'blocos_adicionais', 'campos_duvidosos'
  ],
  properties: {
    empresa_cliente: { type: 'string' },
    unidade: { type: 'string' },
    processo_fluig: { type: 'string' },
    solicitante_nome_cargo: { type: 'string' },
    contato_email: { type: 'string' },
    contato_telefone: { type: 'string' },
    numero_documento: { type: 'string' },
    data_documento: { type: 'string' },
    responsavel_nome: { type: 'string' },
    responsavel_email: { type: 'string' },
    responsavel_telefone: { type: 'string' },
    objeto: { type: 'string' },
    servicos_descricao: { type: 'array', items: { type: 'string' } },
    equipe_tecnica_itens: { type: 'array', items: { type: 'string' } },
    local_servico: { type: 'string' },
    data_servico: { type: 'string' },
    prazo_execucao_dias: { type: 'string' },
    blocos_adicionais: { type: 'array', items: ADDITIONAL_BLOCK_SCHEMA },
    campos_duvidosos: { type: 'array', items: { type: 'string' } }
  }
};

async function importarDocx(filePath, options = {}) {
  const apiKey = String(options.apiKey || '').trim();
  if (!apiKey) {
    throw new Error('Configure a chave da API OpenAI antes de importar uma proposta.');
  }

  const document = extrairConteudoDocx(filePath);
  const requestBody = criarRequisicao(document);
  const fetchImpl = options.fetchImpl || globalThis.fetch;
  if (typeof fetchImpl !== 'function') {
    throw new Error('Este ambiente não oferece suporte à conexão com a API OpenAI.');
  }

  let response;
  try {
    response = await fetchImpl(OPENAI_RESPONSES_URL, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(requestBody),
      signal: options.signal || AbortSignal.timeout(180_000)
    });
  } catch (error) {
    if (error?.name === 'TimeoutError' || error?.name === 'AbortError') {
      throw new Error('A leitura pelo GPT-5.6 Luna excedeu o limite de 3 minutos. Tente novamente.');
    }
    throw new Error('Não foi possível conectar à API OpenAI. Verifique a internet e tente novamente.');
  }

  const payload = await readJsonResponse(response);
  if (!response.ok) {
    throw new Error(formatOpenAiError(response.status, payload));
  }

  const outputText = extractOutputText(payload);
  if (!outputText) {
    const reason = payload.incomplete_details?.reason;
    throw new Error(reason
      ? `A OpenAI não concluiu a leitura do documento: ${reason}.`
      : 'A OpenAI não retornou os dados estruturados da proposta.');
  }

  let extracted;
  try {
    extracted = JSON.parse(outputText);
  } catch {
    throw new Error('A resposta da OpenAI não pôde ser convertida para o formulário.');
  }

  return normalizarResultado(extracted, filePath, payload);
}

function extrairConteudoDocx(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Arquivo DOCX não encontrado: ${filePath || ''}`);
  }
  const stats = fs.statSync(filePath);
  if (!stats.isFile() || stats.size > MAX_DOCX_BYTES) {
    throw new Error('O arquivo deve ser um DOCX válido com no máximo 10 MB.');
  }

  let zip;
  try {
    zip = new PizZip(fs.readFileSync(filePath));
  } catch {
    throw new Error('O arquivo selecionado não é um documento Word válido.');
  }
  const documentFile = zip.file('word/document.xml');
  if (!documentFile) {
    throw new Error('Documento Word inválido ou sem conteúdo principal.');
  }

  const doc = new DOMParser().parseFromString(documentFile.asText(), 'text/xml');
  const blocks = extractOrderedBlocks(doc);
  const serialized = JSON.stringify({ blocks });
  if (serialized.length > MAX_EXTRACTED_CHARS) {
    throw new Error('O texto extraído do documento é grande demais para importação automática.');
  }
  return { blocks };
}

function criarRequisicao(document) {
  return {
    model: OPENAI_MODEL,
    store: false,
    reasoning: { effort: 'low' },
    instructions: [
      'Você extrai propostas comerciais em português para o formulário da Supply Marine.',
      'Copie o conteúdo fielmente: não resuma, não corrija e não invente informações.',
      'Use string vazia, zero ou lista vazia quando um dado não existir.',
      'Datas devem permanecer no formato apresentado, preferencialmente dd/mm/aaaa.',
      'Números monetários devem ser números JSON, interpretando o padrão brasileiro.',
      'Preencha os campos principais apenas com DADOS COMERCIAIS, OBJETO e ESCOPO DE FORNECIMENTO.',
      'Não repita em blocos_adicionais nenhum conteúdo já colocado nos campos principais.',
      'Depois do escopo, reconstrua TODO o conteúdo relevante como blocos_adicionais, mantendo rigorosamente a ordem do Word.',
      'Escolha exatamente um tipo para cada bloco: texto, lista, preco, tabela ou quebra_pagina.',
      'Use texto para tópicos narrativos, cláusulas, listas de observações e seus subtópicos. Remova apenas a numeração automática do título.',
      'Use lista somente para tabelas de documentação com descrição, número do documento e data.',
      'Use preco para tabelas de serviços, consumíveis, locação, materiais ou outros itens monetários e inclua as condições comerciais.',
      'Dentro de preco, use tipo servico para mão de obra/serviços, consumivel para materiais/consumíveis e personalizado nos demais casos.',
      'Use tabela para qualquer grade que não seja documentação nem preço, preservando colunas e células.',
      'Use quebra_pagina para cada marcador page_break recebido do DOCX.',
      'Nos campos que não pertencem ao tipo escolhido, devolva os valores vazios exigidos pelo schema.',
      'Não omita conteúdo por não saber classificá-lo: nesse caso, use um bloco texto e sinalize a dúvida.',
      'Ignore somente cabeçalhos e rodapés repetidos, paginação, logotipo e assinatura visual fixa do modelo.',
      'Liste em campos_duvidosos os nomes dos campos cuja leitura seja ambígua.'
    ].join('\n'),
    input: [{
      role: 'user',
      content: [{
        type: 'input_text',
        text: `Extraia esta proposta Word para o formulário.\n\nCONTEÚDO EXTRAÍDO DO DOCX:\n${JSON.stringify(document)}`
      }]
    }],
    text: {
      verbosity: 'low',
      format: {
        type: 'json_schema',
        name: 'proposta_supply_marine',
        strict: true,
        schema: RESPONSE_SCHEMA
      }
    },
    max_output_tokens: 40_000
  };
}

function normalizarResultado(extracted, filePath, responsePayload) {
  const additionalBlocks = criarBlocosAdicionais(extracted.blocos_adicionais || []);
  const priceBlocks = additionalBlocks.filter((block) => block.tipo === 'preco');
  const priceTopics = priceBlocks.flatMap((block) => block.topicos_preco || []);

  const formData = { ...extracted };
  delete formData.campos_duvidosos;
  formData.blocos_adicionais = additionalBlocks;
  formData.ordem_secoes = [
    'dados_comerciais', 'objeto', 'escopo',
    ...additionalBlocks.map((block) => `flex:${block.id}`)
  ];
  formData.equipe_tecnica = (extracted.equipe_tecnica_itens || []).join(' | ');
  formData.topicos_preco = priceTopics;
  formData.preco_total_numero = priceBlocks.reduce((total, block) => total + numberOrZero(block.preco_total_numero), 0);

  const serviceItems = priceTopics
    .filter((topic) => topic.tipo === 'servico')
    .flatMap((topic) => topic.itens || []);
  const consumableItems = priceTopics
    .filter((topic) => topic.tipo === 'consumivel')
    .flatMap((topic) => topic.itens || []);

  return {
    data: formData,
    itens_servico: serviceItems,
    itens_consumiveis: consumableItems,
    servicos_descricao: extracted.servicos_descricao || [],
    sourcePath: filePath,
    campos_duvidosos: extracted.campos_duvidosos || [],
    ai: {
      model: responsePayload.model || OPENAI_MODEL,
      input_tokens: responsePayload.usage?.input_tokens || 0,
      output_tokens: responsePayload.usage?.output_tokens || 0,
      total_tokens: responsePayload.usage?.total_tokens || 0
    }
  };
}

function criarBlocosAdicionais(sections) {
  return sections.map((section, index) => {
    const type = ['texto', 'lista', 'preco', 'tabela', 'quebra_pagina'].includes(section.tipo)
      ? section.tipo
      : 'texto';
    const base = {
      id: `${type}-ia-${index + 1}`,
      tipo: type,
      titulo: String(section.titulo || '').trim()
    };

    if (type === 'texto') {
      return {
        ...base,
        observacoes: cleanStrings(section.observacoes),
        subtopicos: normalizarSubtopicos(section.subtopicos)
      };
    }
    if (type === 'lista') {
      return {
        ...base,
        linhas: (section.linhas_documentacao || []).map((row) => ({
          descricao: String(row.descricao || ''),
          numero_documento: String(row.numero_documento || ''),
          data: String(row.data || '')
        })).filter((row) => row.descricao || row.numero_documento || row.data)
      };
    }
    if (type === 'preco') {
      return {
        ...base,
        topicos_preco: section.topicos_preco || [],
        preco_total_numero: numberOrZero(section.preco_total_numero),
        preco_total_extenso: section.preco_total_extenso || '',
        moeda: section.moeda || 'Real R$',
        validade_proposta: section.validade_proposta || '',
        pagamento: section.pagamento || '',
        prazo_entrega: section.prazo_entrega || '',
        frete: section.frete || '',
        impostos: section.impostos || ''
      };
    }
    if (type === 'tabela') {
      const columns = (section.colunas_tabela || []).map((name, columnIndex) => ({
        id: `coluna_${columnIndex + 1}`,
        nome: name || `Coluna ${columnIndex + 1}`
      }));
      return {
        ...base,
        colunas: columns,
        linhas: (section.linhas_tabela || []).map((row) => ({
          valores: Object.fromEntries(columns.map((column, columnIndex) => [
            column.id,
            String(row?.[columnIndex] ?? '')
          ]))
        }))
      };
    }
    return base;
  });
}

function normalizarSubtopicos(subtopics) {
  return (subtopics || []).map((subtopic, index) => ({
    id: `subtopico-ia-${index + 1}`,
    titulo: String(subtopic.titulo || ''),
    observacoes: cleanStrings(subtopic.observacoes),
    subtopicos: (subtopic.subtopicos || []).map((nested, nestedIndex) => ({
      id: `subtopico-ia-${index + 1}-${nestedIndex + 1}`,
      titulo: String(nested.titulo || ''),
      observacoes: cleanStrings(nested.observacoes),
      subtopicos: []
    }))
  }));
}

function cleanStrings(values) {
  return (Array.isArray(values) ? values : []).map((value) => String(value || '').trim()).filter(Boolean);
}

async function readJsonResponse(response) {
  try {
    return await response.json();
  } catch {
    return {};
  }
}

function extractOutputText(payload) {
  if (typeof payload.output_text === 'string') return payload.output_text;
  for (const item of payload.output || []) {
    if (item.type !== 'message') continue;
    for (const content of item.content || []) {
      if (content.type === 'output_text' && typeof content.text === 'string') return content.text;
    }
  }
  return '';
}

function formatOpenAiError(status, payload) {
  const message = payload.error?.message || '';
  if (status === 401) return 'A chave da API OpenAI é inválida ou foi revogada.';
  if (status === 429) return 'A API OpenAI está sem saldo, atingiu o limite ou recebeu solicitações demais. Verifique o faturamento.';
  if (status === 413) return 'A proposta é grande demais para a API OpenAI.';
  return `Falha ao importar com a OpenAI${status ? ` (${status})` : ''}${message ? `: ${message}` : '.'}`;
}

function extractOrderedBlocks(doc) {
  const body = getElements(doc, 'w:body')[0];
  if (!body) return [];
  const blocks = [];
  visitDocument(body, blocks, false);
  return blocks;
}

function visitDocument(node, blocks, insideTable) {
  for (let child = node.firstChild; child; child = child.nextSibling) {
    if (child.nodeName === 'w:tbl') {
      const rows = getElements(child, 'w:tr').map((row) => (
      getElements(row, 'w:tc').map((cell) => textContent(cell).replace(/\s+/g, ' ').trim())
      )).filter((row) => row.some(Boolean));
      if (rows.length) blocks.push({ type: 'table', rows });
      continue;
    }
    if (child.nodeName === 'w:p' && !insideTable) {
      const text = textContent(child).replace(/[ \t]+/g, ' ').replace(/\s*\n\s*/g, '\n').trim();
      if (text) blocks.push({ type: 'paragraph', text });
      if (hasPageBreak(child)) blocks.push({ type: 'page_break' });
      continue;
    }
    visitDocument(child, blocks, insideTable || child.nodeName === 'w:tbl');
  }
}

function hasPageBreak(node) {
  let found = false;
  walk(node, (child) => {
    if (child.nodeName === 'w:lastRenderedPageBreak') found = true;
    if (child.nodeName === 'w:br' && (child.getAttribute?.('w:type') || child.getAttribute?.('type')) === 'page') {
      found = true;
    }
  });
  return found;
}

function textContent(node) {
  const parts = [];
  walk(node, (child) => {
    if (child.nodeName === 'w:t') parts.push(child.textContent);
    if (child.nodeName === 'w:tab') parts.push('\t');
    if (child.nodeName === 'w:br' || child.nodeName === 'w:cr') parts.push('\n');
  });
  return parts.join('');
}

function getElements(node, name) {
  const elements = [];
  walk(node, (child) => {
    if (child.nodeName === name) elements.push(child);
  });
  return elements;
}

function walk(node, callback) {
  callback(node);
  for (let child = node.firstChild; child; child = child.nextSibling) walk(child, callback);
}

function numberOrZero(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

module.exports = {
  OPENAI_MODEL,
  RESPONSE_SCHEMA,
  importarDocx,
  extrairConteudoDocx,
  criarRequisicao,
  normalizarResultado
};
