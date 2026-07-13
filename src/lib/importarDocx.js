const fs = require('node:fs');
const PizZip = require('pizzip');
const { DOMParser } = require('@xmldom/xmldom');

function importarDocx(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Arquivo DOCX não encontrado: ${filePath || ''}`);
  }

  const zip = new PizZip(fs.readFileSync(filePath));
  const documentFile = zip.file('word/document.xml');
  if (!documentFile) {
    throw new Error('Documento Word inválido ou sem conteúdo principal.');
  }

  const doc = new DOMParser().parseFromString(documentFile.asText(), 'text/xml');
  const paragraphs = extractParagraphs(doc);
  const tables = extractTables(doc);
  return importarConteudoProposta(paragraphs, tables, filePath);
}

function importarConteudoProposta(paragraphs, tables = [], sourcePath = '') {
  const data = extractFormData(paragraphs, tables);
  const { itens_servico, itens_consumiveis } = extractItems(tables);

  return {
    data,
    itens_servico,
    itens_consumiveis,
    servicos_descricao: data.servicos_descricao || [],
    sourcePath
  };
}

function extractFormData(paragraphs, tables) {
  const data = {};

  const empresaIndex = findParagraphIndex(paragraphs, /^EMPRESA\s*\|?$/i);
  if (empresaIndex >= 0) {
    Object.assign(data, {
      empresa_cliente: paragraphs[empresaIndex + 6] || '',
      unidade: paragraphs[empresaIndex + 7] || '',
      processo_fluig: paragraphs[empresaIndex + 8] || '',
      solicitante_nome_cargo: paragraphs[empresaIndex + 9] || '',
      contato_email: paragraphs[empresaIndex + 10] || '',
      contato_telefone: paragraphs[empresaIndex + 11] || ''
    });
  }

  const documentoIndex = findParagraphIndex(paragraphs, /^N[ºO]\s*DO\s*DOCUMENTO\|?$/i);
  if (documentoIndex >= 0) {
    Object.assign(data, {
      numero_documento: paragraphs[documentoIndex + 5] || '',
      data_documento: paragraphs[documentoIndex + 6] || '',
      responsavel_nome: paragraphs[documentoIndex + 7] || '',
      responsavel_email: paragraphs[documentoIndex + 8] || '',
      responsavel_telefone: paragraphs[documentoIndex + 9] || ''
    });
  }

  data.objeto = valueAfter(paragraphs, 'OBJETO');
  data.equipe_tecnica_itens = valuesBetween(paragraphs, 'Equipe Técnica', 'Local e Data dos Serviços');
  data.equipe_tecnica = data.equipe_tecnica_itens.join(' | ')
    || valueAfter(paragraphs, 'Equipe Técnica');
  applyLocalAndDate(data, valueAfter(paragraphs, 'Local e Data dos Serviços'));
  data.servicos_descricao = valuesBetween(paragraphs, 'Descrição dos Serviços', 'Equipe Técnica');
  data.prazo_execucao_dias = extractPrazoExecucao(paragraphs);
  data.informacoes_adicionais = valuesBetween(
    paragraphs,
    'Os Equipamentos e Ferramentas de propriedade da SUPPLY MARINE deverão ser devolvidos no prazo máximo de 03 (três) dias após a conclusão dos serviços. Caso contrário, a SUPPLY MARINE cobrará pelos custos de cessão dos mesmos conforme tabela abaixo',
    'Atenciosamente,'
  );

  const summaryTable = tables.find((table) => table.some((row) => normalize(row[0]) === 'preco total'));
  if (summaryTable) {
    applySummaryTable(data, summaryTable);
  }

  return removeEmpty(data);
}

function extractItems(tables) {
  const priceTable = tables.find((table) => (
    table.some((row) => normalize(row[0]) === 'servicos')
    && table.some((row) => normalize(row[0]) === 'consumiveis')
  ));

  if (!priceTable) {
    return { itens_servico: [], itens_consumiveis: [] };
  }

  return {
    itens_servico: collectSectionItems(priceTable, 'servicos', ['total servicos', 'consumiveis']),
    itens_consumiveis: collectSectionItems(priceTable, 'consumiveis', ['total consumiveis'])
  };
}

function collectSectionItems(rows, startLabel, endLabels) {
  const startIndex = rows.findIndex((row) => normalize(row[0]) === startLabel);
  if (startIndex < 0) return [];

  const items = [];
  for (let index = startIndex + 1; index < rows.length; index += 1) {
    const row = rows[index];
    const label = normalize(row[0]);
    if (endLabels.includes(label)) break;
    if (label === 'item' || row.length < 7) continue;

    items.push({
      item: row[0] || '',
      descricao: row[1] || '',
      ncm: row[2] || '',
      quant: parseNumber(row[3]),
      un: row[4] || '',
      valor_unit: parseNumber(row[5]),
      valor_total: parseNumber(row[6])
    });
  }

  return items.filter((item) => item.descricao || item.valor_unit || item.quant);
}

function applySummaryTable(data, table) {
  table.forEach((row) => {
    const label = normalize(row[0]);
    if (label === 'preco total') {
      const value = row[1] || '';
      data.preco_total_numero = parseNumber(value);
      const match = value.match(/\((.*)\)/);
      if (match) data.preco_total_extenso = match[1].trim();
      return;
    }

    for (let index = 0; index < row.length; index += 2) {
      const key = normalize(row[index]);
      const value = row[index + 1] || '';
      if (key === 'moeda') data.moeda = value;
      if (key === 'validade da proposta') data.validade_proposta = value;
      if (key === 'pagamento') data.pagamento = value;
      if (key === 'prazo de entrega') data.prazo_entrega = value;
      if (key === 'frete') data.frete = value;
      if (key === 'impostos') data.impostos = value;
    }
  });
}

function applyLocalAndDate(data, value) {
  if (!value) return;
  const parts = value.replace(/\.$/, '').split('|').map((part) => part.trim()).filter(Boolean);
  if (parts.length >= 3) {
    data.data_servico = parts.pop();
    data.local_servico = parts.join(' | ');
    return;
  }
  data.local_servico = value;
}

function extractPrazoExecucao(paragraphs) {
  const line = paragraphs.find((item) => /^Serviços:\s*Estimados/i.test(item));
  if (!line) return '';
  const match = line.match(/Serviços:\s*Estimados\s+em\s+(.+?)(?:,|$)/i);
  return match ? match[1].trim() : '';
}

function valueAfter(paragraphs, label) {
  const index = paragraphs.findIndex((item) => normalize(item) === normalize(label));
  return index >= 0 ? paragraphs[index + 1] || '' : '';
}

function valuesBetween(paragraphs, startLabel, endLabel) {
  const start = paragraphs.findIndex((item) => normalize(item) === normalize(startLabel));
  const end = paragraphs.findIndex((item, index) => index > start && normalize(item) === normalize(endLabel));
  if (start < 0 || end < 0) return [];
  return paragraphs.slice(start + 1, end).filter(Boolean);
}

function findParagraphIndex(paragraphs, pattern) {
  return paragraphs.findIndex((item) => pattern.test(item));
}

function extractParagraphs(doc) {
  return getElements(doc, 'w:p')
    .map((paragraph) => textContent(paragraph).replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

function extractTables(doc) {
  return getElements(doc, 'w:tbl').map((table) => (
    getElements(table, 'w:tr').map((row) => (
      getElements(row, 'w:tc').map((cell) => textContent(cell).replace(/\s+/g, ' ').trim())
    )).filter((row) => row.some(Boolean))
  )).filter((table) => table.length);
}

function textContent(node) {
  const parts = [];
  walk(node, (child) => {
    if (child.nodeName === 'w:t') parts.push(child.textContent);
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
  for (let child = node.firstChild; child; child = child.nextSibling) {
    walk(child, callback);
  }
}

function parseNumber(value) {
  const clean = String(value || '')
    .replace(/[^\d,.-]/g, '')
    .replace(/\.(?=\d{3}(?:\D|$))/g, '')
    .replace(',', '.');
  const number = Number(clean);
  return Number.isFinite(number) ? number : 0;
}

function normalize(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[|:]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function removeEmpty(data) {
  return Object.fromEntries(Object.entries(data).filter(([, value]) => (
    Array.isArray(value) ? value.length : String(value || '').trim() !== ''
  )));
}

module.exports = {
  importarDocx,
  importarConteudoProposta
};
