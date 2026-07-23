const fs = require('node:fs');
const path = require('node:path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { getPropostasDir } = require('./paths');

const templatePath = path.resolve(__dirname, '..', 'templates', 'modelo_proposta.docx');
function formatCurrency(value) {
  const number = Number(value || 0);
  return new Intl.NumberFormat('pt-BR', {
    style: 'currency',
    currency: 'BRL'
  }).format(number).replace(/\u00a0/g, ' ');
}

function formatNumber(value) {
  const number = Number(value || 0);
  return new Intl.NumberFormat('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(number);
}

function normalizeItems(items = []) {
  return items.map((item) => {
    const quant = Number(item.quant || 0);
    const valorUnit = Number(item.valor_unit || 0);
    const total = item.valor_total == null || item.valor_total === ''
      ? quant * valorUnit
      : Number(item.valor_total || 0);

    return {
      ...item,
      codigo: item.codigo || '',
      quant: String(item.quant ?? ''),
      valor_unit: formatCurrency(valorUnit),
      valor_total: formatCurrency(total)
    };
  });
}

function sumItems(items = []) {
  return items.reduce((sum, item) => {
    const total = item.valor_total == null || item.valor_total === ''
      ? Number(item.quant || 0) * Number(item.valor_unit || 0)
      : Number(item.valor_total || 0);
    return sum + total;
  }, 0);
}

function legacyTopicsFromData(data) {
  const topics = [];
  if ((data.itens_servico || []).length) {
    topics.push({ titulo: 'SERVI\u00c7OS', tipo: 'servico', itens: data.itens_servico });
  }
  if ((data.itens_consumiveis || []).length) {
    topics.push({ titulo: 'MATERIAIS CONTRATO', tipo: 'consumivel', itens: data.itens_consumiveis });
  }
  return topics;
}

function normalizeTopics(data) {
  const topics = Array.isArray(data.topicos_preco) && data.topicos_preco.length
    ? data.topicos_preco
    : legacyTopicsFromData(data);

  return topics.map((topic) => {
    const itens = normalizeItems(topic.itens || []);
    const totalNumber = sumItems(topic.itens || []);
    return {
      ...topic,
      titulo: String(topic.titulo || '').toUpperCase(),
      itens,
      total: formatCurrency(totalNumber),
      totalNumber
    };
  }).filter((topic) => topic.itens.length);
}

function prepareData(data) {
  const topicosPreco = normalizeTopics(data);
  const itensServico = topicosPreco
    .filter((topic) => topic.tipo === 'servico')
    .flatMap((topic) => topic.itens);
  const itensConsumiveis = topicosPreco
    .filter((topic) => topic.tipo === 'consumivel')
    .flatMap((topic) => topic.itens);
  const totalServicosNumber = topicosPreco
    .filter((topic) => topic.tipo === 'servico')
    .reduce((sum, topic) => sum + topic.totalNumber, 0);
  const totalConsumiveisNumber = topicosPreco
    .filter((topic) => topic.tipo === 'consumivel')
    .reduce((sum, topic) => sum + topic.totalNumber, 0);
  const totalTopicosNumber = topicosPreco.reduce((sum, topic) => sum + topic.totalNumber, 0);
  const totalNumber = data.preco_total_numero || totalTopicosNumber;
  const technicalTeamSource = Array.isArray(data.equipe_tecnica_itens) && data.equipe_tecnica_itens.length
    ? data.equipe_tecnica_itens
    : String(data.equipe_tecnica || '').split('|');
  const technicalTeam = technicalTeamSource
    .map((item) => (typeof item === 'string' ? item : item?.item))
    .map((item) => String(item || '').trim())
    .filter(Boolean);
  return {
    ...data,
    servicos_descricao: (data.servicos_descricao || []).map((item) => (
      typeof item === 'string' ? { item } : item
    )),
    equipe_tecnica: technicalTeam.join(' | '),
    equipe_tecnica_itens: technicalTeam.map((item) => ({ item })),
    topicos_preco: topicosPreco,
    itens_servico: itensServico,
    itens_consumiveis: itensConsumiveis,
    total_servicos: formatCurrency(totalServicosNumber),
    total_consumiveis: formatCurrency(totalConsumiveisNumber),
    preco_total_formatado: formatCurrency(totalNumber),
    preco_total_numero: formatCurrency(totalNumber),
    preco_total_sem_moeda: formatNumber(totalNumber)
  };
}

function gerarDocx(data, outputPath) {
  if (!fs.existsSync(templatePath)) {
    throw new Error(`Template não encontrado: ${templatePath}`);
  }

  const content = fs.readFileSync(templatePath);
  const zip = new PizZip(content);
  normalizeTemplateCurrencyCells(zip);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: () => ''
  });

  doc.render(prepareData(data));

  injectAdditionalBlocks(doc.getZip(), data);

  const buffer = doc.getZip().generate({
    type: 'nodebuffer',
    compression: 'DEFLATE'
  });

  const outputDir = getPropostasDir();
  fs.mkdirSync(outputDir, { recursive: true });
  const finalPath = outputPath || path.join(outputDir, `${data.numero_documento || 'proposta'}-teste.docx`);
  fs.writeFileSync(finalPath, buffer);
  return finalPath;
}

function injectAdditionalBlocks(zip, proposalData = {}) {
  const validBlocks = Array.isArray(proposalData.blocos_adicionais)
    ? proposalData.blocos_adicionais.filter(Boolean)
    : [];
  const documentFile = zip.file('word/document.xml');
  if (!documentFile) return;

  const normalizedPriceHeaderXml = normalizeOfficialPriceHeader(documentFile.asText());
  const normalizedFieldAlignmentXml = normalizeFixedFieldAlignment(normalizedPriceHeaderXml);
  const originalDocumentXml = normalizeOptionalServiceLocation(normalizedFieldAlignmentXml, proposalData);
  const sectionLineXml = extractSectionLineParagraph(originalDocumentXml);
  const documentXml = removeFixedAdditionalInfoSection(originalDocumentXml);
  const rebuiltDocument = rebuildOrderedDocumentSections(
    documentXml,
    validBlocks,
    proposalData.ordem_secoes,
    sectionLineXml,
    proposalData.secoes_excluidas
  );
  zip.file('word/document.xml', rebuiltDocument);
}

function normalizeTemplateCurrencyCells(zip) {
  const documentFile = zip.file('word/document.xml');
  if (!documentFile) return;

  let documentXml = documentFile.asText();
  ['{valor_unit}', '{valor_total}', '{total}'].forEach((tag) => {
    const tagIndex = documentXml.indexOf(tag);
    if (tagIndex < 0) return;

    const currencyTag = '<w:t>R$</w:t>';
    const currencyIndex = documentXml.lastIndexOf(currencyTag, tagIndex);
    if (currencyIndex < 0 || tagIndex - currencyIndex > 700) return;

    documentXml = `${documentXml.slice(0, currencyIndex)}<w:t></w:t>${documentXml.slice(currencyIndex + currencyTag.length)}`;
  });

  zip.file('word/document.xml', documentXml);
}

function normalizeFixedFieldAlignment(documentXml) {
  const objectHeadingIndex = documentXml.indexOf('OBJETO');
  const scopeHeadingIndex = documentXml.indexOf('ESCOPO DE FORNECIMENTO', objectHeadingIndex);
  if (objectHeadingIndex < 0 || scopeHeadingIndex < 0) return documentXml;

  const objectSectionStart = findParagraphStart(documentXml, objectHeadingIndex);
  const scopeSectionStart = findParagraphStart(documentXml, scopeHeadingIndex);
  if (objectSectionStart < 0 || scopeSectionStart <= objectSectionStart) return documentXml;

  const objectSection = documentXml.slice(objectSectionStart, scopeSectionStart).replace(
    /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,
    (paragraphXml) => normalizeFirstLineIndent(paragraphXml)
  );

  return `${documentXml.slice(0, objectSectionStart)}${objectSection}${documentXml.slice(scopeSectionStart)}`;
}

function normalizeFirstLineIndent(paragraphXml) {
  return paragraphXml.replace(/<w:ind\b([^>]*)\/>/, (indentXml, attributes) => {
    const firstLineMatch = attributes.match(/\sw:firstLine="(-?\d+)"/);
    if (!firstLineMatch) return indentXml;

    const currentLeftMatch = attributes.match(/\sw:left="(-?\d+)"/);
    const currentLeft = Number(currentLeftMatch?.[1] || 0);
    const firstLine = Number(firstLineMatch[1] || 0);
    const alignedLeft = currentLeft + firstLine;
    let normalizedAttributes = attributes.replace(/\sw:firstLine="-?\d+"/, '');

    if (currentLeftMatch) {
      normalizedAttributes = normalizedAttributes.replace(/\sw:left="-?\d+"/, ` w:left="${alignedLeft}"`);
    } else {
      normalizedAttributes += ` w:left="${alignedLeft}"`;
    }

    return `<w:ind${normalizedAttributes}/>`;
  });
}

function normalizeOptionalServiceLocation(documentXml, proposalData = {}) {
  const location = String(proposalData.local_servico || '').trim();
  const serviceDate = String(proposalData.data_servico || '').trim();
  const headingIndex = documentXml.indexOf('Local e Data dos Servi\u00e7os');
  if (headingIndex < 0) return documentXml;

  const headingStart = findParagraphStart(documentXml, headingIndex);
  const headingEndMarker = '</w:p>';
  const headingEnd = documentXml.indexOf(headingEndMarker, headingIndex);
  const valueStart = findNextParagraphStart(documentXml, headingEnd + headingEndMarker.length);
  const valueEnd = valueStart >= 0 ? documentXml.indexOf(headingEndMarker, valueStart) : -1;
  if (headingStart < 0 || headingEnd < 0 || valueStart < 0 || valueEnd < 0) return documentXml;

  if (!location && !serviceDate) {
    let removalStart = headingStart;
    let removalEnd = valueEnd + headingEndMarker.length;
    const previousEnd = documentXml.lastIndexOf(headingEndMarker, headingStart);
    const previousStart = previousEnd >= 0 ? findParagraphStart(documentXml, previousEnd) : -1;
    if (previousStart >= 0 && isEmptyWordParagraph(documentXml.slice(previousStart, previousEnd + headingEndMarker.length))) {
      removalStart = previousStart;
    }
    const nextStart = findNextParagraphStart(documentXml, removalEnd);
    const nextEnd = nextStart >= 0 ? documentXml.indexOf(headingEndMarker, nextStart) : -1;
    if (nextStart >= 0 && nextEnd >= 0 && isEmptyWordParagraph(documentXml.slice(nextStart, nextEnd + headingEndMarker.length))) {
      removalEnd = nextEnd + headingEndMarker.length;
    }
    return `${documentXml.slice(0, removalStart)}${documentXml.slice(removalEnd)}`;
  }

  const displayValue = `${[location, serviceDate].filter(Boolean).join(' | ')}.`;
  const valueParagraph = documentXml.slice(valueStart, valueEnd + headingEndMarker.length);
  const normalizedParagraph = valueParagraph.replace(
    /(<w:t(?:\s[^>]*)?>)[\s\S]*?(<\/w:t>)/,
    `$1${escapeXml(displayValue)}$2`
  );
  return `${documentXml.slice(0, valueStart)}${normalizedParagraph}${documentXml.slice(valueEnd + headingEndMarker.length)}`;
}

function findNextParagraphStart(documentXml, fromIndex) {
  const match = /<w:p(?:\s[^>]*)?>/.exec(documentXml.slice(Math.max(0, fromIndex)));
  return match ? Math.max(0, fromIndex) + match.index : -1;
}

function isEmptyWordParagraph(paragraphXml) {
  const text = Array.from(paragraphXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g))
    .map((match) => match[1])
    .join('')
    .trim();
  return !text && !/<w:(?:drawing|pict|br|object)\b/.test(paragraphXml);
}

function normalizeOfficialPriceHeader(documentXml) {
  const markers = ['Cabe\u00e7alho de pre\u00e7os', 'cabecalho-preco.svg', 'rIdCabecalhoPreco'];
  const markerIndex = markers.reduce((found, marker) => (
    found >= 0 ? found : documentXml.indexOf(marker)
  ), -1);
  if (markerIndex < 0) return documentXml;

  const rowStart = Math.max(
    documentXml.lastIndexOf('<w:tr>', markerIndex),
    documentXml.lastIndexOf('<w:tr ', markerIndex)
  );
  const rowEndMarker = '</w:tr>';
  const rowEnd = documentXml.indexOf(rowEndMarker, markerIndex);
  if (rowStart < 0 || rowEnd < 0) return documentXml;

  const header = renderRoundedHeaderRow(
    ['ITEM', 'DESCRI\u00c7\u00c3O DOS ITENS', 'NCM', 'QTD', 'UNID.', 'VALOR UNIT\u00c1RIO', 'VALOR TOTAL'],
    [710, 4540, 945, 710, 590, 1315, 1435],
    {
      shapeName: 'CabecalhoPrecosArredondado',
      shapeNumber: 9290,
      anchorId: 'B0000FFE'
    }
  );
  return `${documentXml.slice(0, rowStart)}${header}${documentXml.slice(rowEnd + rowEndMarker.length)}`;
}

const FIXED_DOCUMENT_SECTIONS = [
  { id: 'dados_comerciais', marker: 'DADOS COMERCIAIS' },
  { id: 'objeto', marker: 'OBJETO' },
  { id: 'escopo', marker: 'ESCOPO DE FORNECIMENTO' },
  { id: 'preco', marker: 'PRE\u00c7O', removed: true }
];

const ACTIVE_FIXED_DOCUMENT_SECTIONS = FIXED_DOCUMENT_SECTIONS.filter((section) => !section.removed);
const SIGNATURE_SPACER = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>';

function rebuildOrderedDocumentSections(documentXml, blocks, requestedOrder, sectionLineXml, excludedSections = []) {
  const extracted = extractFixedDocumentSections(documentXml);
  if (!extracted) {
    return insertAdditionalBlocksBeforeSignature(documentXml, renderAdditionalBlocks(blocks, sectionLineXml));
  }

  const blockEntries = blocks.map((block, index) => ({
    id: `flex:${block.id || `legacy-${index}`}`,
    block
  }));
  const excludedFixedIds = new Set(Array.isArray(excludedSections) ? excludedSections : []);
  const activeFixedSections = ACTIVE_FIXED_DOCUMENT_SECTIONS
    .filter((section) => !excludedFixedIds.has(section.id));
  const availableIds = new Set([
    ...activeFixedSections.map((section) => section.id),
    ...blockEntries.map((entry) => entry.id)
  ]);
  const order = [];
  (Array.isArray(requestedOrder) ? requestedOrder : []).forEach((sectionId) => {
    if (!availableIds.has(sectionId) || order.includes(sectionId)) return;
    order.push(sectionId);
  });
  activeFixedSections.forEach((section) => {
    if (!order.includes(section.id)) order.push(section.id);
  });
  blockEntries.forEach((entry) => {
    if (!order.includes(entry.id)) order.push(entry.id);
  });

  const blocksById = new Map(blockEntries.map((entry) => [entry.id, entry.block]));
  let sectionNumber = 1;
  const orderedXml = order.map((sectionId) => {
    if (extracted.sections[sectionId]) {
      const rendered = numberFixedDocumentSection(extracted.sections[sectionId], sectionId, sectionNumber);
      sectionNumber += 1;
      return rendered;
    }
    const block = blocksById.get(sectionId);
    if (!block) return '';
    if (block.tipo === 'quebra_pagina') {
      return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    }
    const rendered = renderAdditionalBlock(
      block,
      sectionNumber,
      cloneSectionLineParagraph(sectionLineXml, sectionNumber)
    );
    sectionNumber += 1;
    return rendered;
  }).join('');

  return `${extracted.prefix}${orderedXml}${SIGNATURE_SPACER}${extracted.suffix}`;
}

function extractFixedDocumentSections(documentXml) {
  const starts = {};
  for (const section of FIXED_DOCUMENT_SECTIONS) {
    const markerIndex = documentXml.indexOf(section.marker);
    const start = markerIndex >= 0 ? findParagraphStart(documentXml, markerIndex) : -1;
    if (start < 0) return null;
    starts[section.id] = start;
  }
  const signatureIndex = documentXml.indexOf('Atenciosamente');
  const signatureStart = signatureIndex >= 0 ? findParagraphStart(documentXml, signatureIndex) : -1;
  if (signatureStart < 0) return null;

  const sections = {};
  FIXED_DOCUMENT_SECTIONS.forEach((section, index) => {
    const nextSection = FIXED_DOCUMENT_SECTIONS[index + 1];
    const end = nextSection ? starts[nextSection.id] : signatureStart;
    sections[section.id] = documentXml.slice(starts[section.id], end);
  });
  return {
    prefix: documentXml.slice(0, starts.dados_comerciais),
    sections,
    suffix: documentXml.slice(signatureStart)
  };
}

function numberFixedDocumentSection(sectionXml, sectionId, number) {
  const definition = FIXED_DOCUMENT_SECTIONS.find((section) => section.id === sectionId);
  let numberedXml = prefixParagraphMarker(sectionXml, definition.marker, `${number}.  `);
  if (sectionId === 'escopo') {
    ['Descri\u00e7\u00e3o dos Servi\u00e7os', 'Equipe T\u00e9cnica', 'Local e Data dos Servi\u00e7os'].forEach((marker, index) => {
      numberedXml = prefixParagraphMarker(numberedXml, marker, `${number}.${index + 1}  `);
    });
  }
  return numberedXml;
}

function prefixParagraphMarker(sectionXml, marker, prefix) {
  const markerIndex = sectionXml.indexOf(marker);
  if (markerIndex < 0) return sectionXml;
  const paragraphStart = findParagraphStart(sectionXml, markerIndex);
  const paragraphEnd = sectionXml.indexOf('</w:p>', markerIndex);
  if (paragraphStart < 0 || paragraphEnd < 0) return sectionXml;
  const paragraph = sectionXml.slice(paragraphStart, paragraphEnd + '</w:p>'.length)
    .replace(/<w:numPr>[\s\S]*?<\/w:numPr>/, '')
    .replace(marker, `${prefix}${marker}`);
  return `${sectionXml.slice(0, paragraphStart)}${paragraph}${sectionXml.slice(paragraphEnd + '</w:p>'.length)}`;
}

function insertAdditionalBlocksBeforeSignature(documentXml, contentXml) {
  if (!contentXml) return documentXml;
  const signatureIndex = documentXml.indexOf('Atenciosamente');
  const insertionIndex = signatureIndex >= 0 ? findParagraphStart(documentXml, signatureIndex) : -1;
  if (insertionIndex < 0) return documentXml;
  return `${documentXml.slice(0, insertionIndex)}${contentXml}${SIGNATURE_SPACER}${documentXml.slice(insertionIndex)}`;
}

function removeFixedAdditionalInfoSection(documentXml) {
  const sectionIndex = documentXml.indexOf('INFORMAÇÕES ADICIONAIS');
  const signatureIndex = documentXml.indexOf('Atenciosamente', Math.max(0, sectionIndex));
  if (sectionIndex < 0 || signatureIndex < 0) return documentXml;

  const sectionStart = findParagraphStart(documentXml, sectionIndex);
  const signatureStart = findParagraphStart(documentXml, signatureIndex);
  if (sectionStart < 0 || signatureStart < 0 || signatureStart <= sectionStart) return documentXml;
  return `${documentXml.slice(0, sectionStart)}${documentXml.slice(signatureStart)}`;
}

function findParagraphStart(documentXml, beforeIndex) {
  const pattern = /<w:p(?:\s[^>]*)?>/g;
  let paragraphStart = -1;
  let match;
  while ((match = pattern.exec(documentXml)) && match.index < beforeIndex) {
    paragraphStart = match.index;
  }
  return paragraphStart;
}

function extractSectionLineParagraph(documentXml) {
  const markerIndex = documentXml.indexOf('ESCOPO DE FORNECIMENTO');
  if (markerIndex < 0) return '';
  const headingEnd = documentXml.indexOf('</w:p>', markerIndex);
  if (headingEnd < 0) return '';
  const remainder = documentXml.slice(headingEnd + '</w:p>'.length);
  const lineParagraph = remainder.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/);
  return lineParagraph?.[0] || '';
}

function cloneSectionLineParagraph(lineXml, sequence) {
  if (!lineXml) return '';
  let idCounter = 0;
  const nextHexId = () => (0xA0000000 + (sequence * 16) + (idCounter += 1)).toString(16).toUpperCase();
  return lineXml
    .replace(/w14:paraId="[^"]+"/g, () => `w14:paraId="${nextHexId()}"`)
    .replace(/wp14:anchorId="[^"]+"/g, () => `wp14:anchorId="${nextHexId()}"`)
    .replace(/wp14:editId="[^"]+"/g, () => `wp14:editId="${nextHexId()}"`)
    .replace(/wp:docPr id="\d+" name="[^"]*"/g, `wp:docPr id="${9000 + sequence}" name="Linha de t\u00f3pico ${sequence}"`)
    .replace(/w14:anchorId="[^"]+"/g, () => `w14:anchorId="${nextHexId()}"`)
    .replace(/id="Conector reto \d+"/g, `id="Linha de t\u00f3pico ${sequence}"`)
    .replace(/o:spid="_x0000_s\d+"/g, `o:spid="_x0000_s${9000 + sequence}"`);
}

function renderAdditionalBlocks(blocks, sectionLineXml = '') {
  let sectionNumber = 4;
  return blocks.map((block) => {
    if (block.tipo === 'quebra_pagina') {
      return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    }

    const rendered = renderAdditionalBlock(
      block,
      sectionNumber,
      cloneSectionLineParagraph(sectionLineXml, sectionNumber)
    );
    sectionNumber += 1;
    return rendered;
  }).join('');
}

function renderAdditionalBlock(block, sectionNumber, sectionLine) {
  if (block.tipo === 'texto') {
    return renderStructuredTopic(block, sectionNumber, sectionLine);
  }
  if (block.tipo === 'lista') {
    return `${renderNumberedTopicHeading(sectionNumber, 'DOCUMENTA\u00c7\u00c3O', sectionLine)}${renderDocumentList(block, sectionNumber)}`;
  }
  if (block.tipo === 'preco') {
    return `${renderNumberedTopicHeading(sectionNumber, 'PRE\u00c7O', sectionLine)}${renderAdditionalPrice(block, sectionNumber)}`;
  }
  if (block.tipo === 'tabela') {
    const title = stripAutomaticNumber(block.titulo || '') || 'TABELA SEM NOME';
    return `${renderNumberedTopicHeading(sectionNumber, title, sectionLine)}${renderCustomTable(block, sectionNumber)}`;
  }
  return '';
}

function renderStructuredTopic(block, number, sectionLineXml = '') {
  const title = stripAutomaticNumber(block.titulo || '') || 'T\u00d3PICO SEM NOME';
  const observations = Array.isArray(block.observacoes) && block.observacoes.length
    ? block.observacoes
    : (block.conteudo ? String(block.conteudo).split(/\r?\n/) : []);
  const subtopics = Array.isArray(block.subtopicos) ? block.subtopicos : [];
  const heading = renderNumberedTopicHeading(number, title, sectionLineXml);
  const topicObservations = observations.map((observation) => renderTopicObservation(observation)).join('');
  const subtopicsXml = renderStructuredSubtopics(subtopics, String(number));
  return `${heading}${topicObservations}${subtopicsXml}`;
}

function renderStructuredSubtopics(subtopics, parentNumber, depth = 1) {
  return (Array.isArray(subtopics) ? subtopics : []).map((subtopic, index) => {
    const number = `${parentNumber}.${index + 1}`;
    const title = stripAutomaticNumber(subtopic?.titulo || '') || 'Subt\u00f3pico sem nome';
    const observations = Array.isArray(subtopic?.observacoes) ? subtopic.observacoes : [];
    const nested = Array.isArray(subtopic?.subtopicos) ? subtopic.subtopicos : [];
    return `${renderSubtopicHeading(number, title, depth)}${observations.map((observation) => renderTopicObservation(observation, depth - 1)).join('')}${renderStructuredSubtopics(nested, number, depth + 1)}`;
  }).join('');
}

function renderNumberedTopicHeading(number, title, sectionLineXml = '') {
  const line = sectionLineXml || `<w:p>
    <w:pPr>
      <w:ind w:left="390" w:right="817"/><w:spacing w:before="0" w:after="70" w:line="20" w:lineRule="exact"/>
      <w:pBdr><w:bottom w:val="single" w:sz="4" w:space="0" w:color="002060"/></w:pBdr>
    </w:pPr>
    <w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>
  </w:p>`;
  return `<w:p>
    <w:pPr>
      <w:pStyle w:val="PargrafodaLista"/><w:keepNext/>
      <w:ind w:left="425" w:right="799"/><w:spacing w:before="240" w:after="0"/>
      <w:contextualSpacing w:val="0"/>
    </w:pPr>
    <w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:b/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">${number}.  ${escapeXml(String(title).toUpperCase())}</w:t></w:r>
  </w:p>
  ${line}`;
}

function renderSubtopicHeading(number, title, depth = 1) {
  const numberIndent = 680 + (Math.max(1, depth) - 1) * 280;
  const hangingIndent = 284;
  const textIndent = numberIndent + hangingIndent;
  return `<w:p>
    <w:pPr><w:pStyle w:val="PargrafodaLista"/><w:keepNext/><w:spacing w:before="80" w:after="30"/><w:ind w:left="${textIndent}" w:right="799" w:hanging="${hangingIndent}"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Times New Roman" w:cs="Calibri"/><w:color w:val="000000"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">${escapeXml(number)}  ${escapeXml(title)}</w:t></w:r>
  </w:p>`;
}

function renderTopicObservation(value, nestedDepth = 0) {
  const text = typeof value === 'string' ? value : value?.texto || '';
  if (!String(text).trim()) return '';
  const leftIndent = 1191 + Math.max(0, nestedDepth) * 280;
  return `<w:p>
    <w:pPr>
      <w:pStyle w:val="PargrafodaLista"/>
      <w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>
      <w:ind w:left="${leftIndent}" w:right="799" w:hanging="227"/><w:spacing w:after="20"/>
    </w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Times New Roman" w:cs="Calibri"/><w:color w:val="000000"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">${escapeXml(String(text).trim())}</w:t></w:r>
  </w:p>`;
}

function stripAutomaticNumber(value) {
  return String(value || '').replace(/^\s*(?:\d+\.\s+|\d+(?:\.\d+)+\s+)/, '').trim();
}

function renderDocumentList(block, sectionNumber = 5) {
  const sourceRows = Array.isArray(block.linhas) && block.linhas.length
    ? block.linhas
    : (Array.isArray(block.itens) ? block.itens : []);
  const rows = sourceRows.map((row) => (
    typeof row === 'string'
      ? { descricao: row, numero_documento: '', data: '' }
      : {
        descricao: row?.descricao || row?.item || '',
        numero_documento: row?.numero_documento || '',
        data: row?.data || ''
      }
  )).filter((row) => row.descricao || row.numero_documento || row.data);
  if (!rows.length) return '';

  const widths = [4303, 3688, 2254];
  const header = renderRoundedDocumentListHeader(widths, sectionNumber);
  const content = rows.map((row) => renderDocumentListRow([
    row.descricao,
    row.numero_documento,
    row.data
  ], widths)).join('');

  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="10245" w:type="dxa"/><w:tblInd w:w="390" w:type="dxa"/><w:tblLayout w:type="fixed"/>
      <w:tblBorders>
        <w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>
        <w:insideH w:val="nil"/><w:insideV w:val="nil"/>
      </w:tblBorders>
      <w:tblCellMar>
        <w:top w:w="30" w:type="dxa"/><w:left w:w="60" w:type="dxa"/>
        <w:bottom w:w="30" w:type="dxa"/><w:right w:w="60" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
    <w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
    ${header}${content}
  </w:tbl><w:p><w:pPr><w:spacing w:after="80"/></w:pPr></w:p>`;
}

function renderRoundedDocumentListHeader(widths, sectionNumber) {
  const safeNumber = Number.isFinite(Number(sectionNumber)) ? Number(sectionNumber) : 5;
  const integerNumber = Math.max(0, Math.trunc(safeNumber));
  return renderRoundedHeaderRow(
    ['DESCRI\u00c7\u00c3O', 'N\u00ba DO DOCUMENTO', 'DATA'],
    widths,
    {
      shapeName: `CabecalhoDocumentacao${safeNumber}`,
      shapeNumber: 9300 + integerNumber,
      anchorId: (0xB0000000 + integerNumber).toString(16).toUpperCase().slice(-8)
    }
  );
}

function renderRoundedHeaderRow(headings, widths, options = {}) {
  const tableWidth = widths.reduce((sum, width) => sum + width, 0);
  const shapeName = options.shapeName || 'CabecalhoArredondado';
  const shapeNumber = options.shapeNumber || 9299;
  const anchorId = options.anchorId || 'B0000FFF';
  const outerGridSpan = options.outerGridSpan || headings.length;
  const nestedCells = headings.map((heading, index) => `<w:tc>
    <w:tcPr>
      <w:tcW w:w="${widths[index]}" w:type="dxa"/>
      <w:tcBorders>
        <w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>
        <w:right w:val="${index < headings.length - 1 ? 'single' : 'nil'}"${index < headings.length - 1 ? ' w:sz="6" w:space="0" w:color="404040"' : ''}/>
      </w:tcBorders>
      <w:tcMar><w:top w:w="35" w:type="dxa"/><w:left w:w="70" w:type="dxa"/><w:bottom w:w="35" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tcMar>
      <w:vAlign w:val="center"/>
    </w:tcPr>
    <w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:b/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t>${escapeXml(heading)}</w:t></w:r>
    </w:p>
  </w:tc>`).join('');

  return `<w:tr>
    <w:trPr><w:trHeight w:hRule="exact" w:val="340"/></w:trPr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="${tableWidth}" w:type="dxa"/><w:gridSpan w:val="${outerGridSpan}"/>
        <w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>
        <w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>
        <w:vAlign w:val="center"/>
      </w:tcPr>
      <w:p><w:pPr><w:spacing w:line="340" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr>
        <w:r><w:pict w14:anchorId="${anchorId}">
          <v:roundrect id="${shapeName}" o:spid="_x0000_s${shapeNumber}" style="width:${tableWidth / 20}pt;height:15.1pt;mso-left-percent:-10001;mso-top-percent:-10001;mso-position-horizontal:absolute;mso-position-horizontal-relative:char;mso-position-vertical:absolute;mso-position-vertical-relative:line;mso-left-percent:-10001;mso-top-percent:-10001" arcsize="7864f" fillcolor="#bfbfbf" strokecolor="#404040">
            <v:textbox inset="0,0,0,0"><w:txbxContent>
              <w:tbl>
                <w:tblPr><w:tblW w:w="${tableWidth}" w:type="dxa"/><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>
                <w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
                <w:tr>${nestedCells}</w:tr>
              </w:tbl>
              <w:p/>
            </w:txbxContent></v:textbox>
          </v:roundrect>
        </w:pict></w:r>
      </w:p>
    </w:tc>
  </w:tr>`;
}

function renderDocumentListRow(values, widths) {
  return `<w:tr><w:trPr><w:trHeight w:val="250" w:hRule="atLeast"/></w:trPr>${values.map((value, index) => `<w:tc>
    <w:tcPr>
      <w:tcW w:w="${widths[index]}" w:type="dxa"/>
      <w:tcBorders><w:bottom w:val="single" w:sz="4" w:color="B7B7B7"/></w:tcBorders>
      <w:vAlign w:val="center"/>
    </w:tcPr>
    <w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">${escapeXml(value)}</w:t></w:r>
    </w:p>
  </w:tc>`).join('')}</w:tr>`;
}

const ADDITIONAL_PRICE_COLUMNS = [710, 4540, 945, 710, 590, 1315, 1435];
const ADDITIONAL_PRICE_SUMMARY_COLUMNS = [2760, 2363, 2362, 2760];
const PRICE_NO_BORDERS = '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>';
const PRICE_ROW_BOTTOM = '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="nil"/></w:tcBorders>';
const PRICE_TOTAL_TOP = '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>';

function renderAdditionalPrice(block, sectionNumber) {
  const topics = (Array.isArray(block.topicos_preco) ? block.topicos_preco : []).map((topic) => {
    const sourceItems = Array.isArray(topic?.itens) ? topic.itens : [];
    return {
      titulo: String(topic?.titulo || 'T\u00d3PICO SEM NOME').toUpperCase(),
      itens: normalizeItems(sourceItems),
      totalNumber: sumItems(sourceItems)
    };
  }).filter((topic) => topic.itens.length);
  const calculatedTotal = topics.reduce((sum, topic) => sum + topic.totalNumber, 0);
  const requestedTotal = Number(block.preco_total_numero);
  const totalNumber = Number.isFinite(requestedTotal) && requestedTotal !== 0 ? requestedTotal : calculatedTotal;
  const totalText = formatCurrency(totalNumber);
  const totalInWords = String(block.preco_total_extenso || '').trim();
  const tableWidth = ADDITIONAL_PRICE_COLUMNS.reduce((sum, width) => sum + width, 0);
  const numberSeed = Math.max(0, Math.trunc(Number(sectionNumber) || 0));

  const header = renderRoundedHeaderRow(
    ['ITEM', 'DESCRI\u00c7\u00c3O DOS ITENS', 'NCM', 'QTD', 'UNID.', 'VALOR UNIT\u00c1RIO', 'VALOR TOTAL'],
    ADDITIONAL_PRICE_COLUMNS,
    {
      shapeName: `CabecalhoPrecoAdicional${numberSeed}`,
      shapeNumber: 9500 + numberSeed,
      anchorId: (0xC0000000 + numberSeed).toString(16).toUpperCase().slice(-8)
    }
  );
  const topicRows = topics.map((topic) => {
    const titleRow = renderAdditionalPriceRow([
      renderAdditionalPriceCell(
        renderAdditionalPriceParagraph(topic.titulo, { bold: true }),
        tableWidth,
        { gridSpan: 7, borders: PRICE_NO_BORDERS }
      )
    ]);
    const itemRows = topic.itens.map((item) => renderAdditionalPriceRow([
      renderAdditionalPriceCell(renderAdditionalPriceParagraph(item.item, { align: 'center' }), ADDITIONAL_PRICE_COLUMNS[0], { borders: PRICE_ROW_BOTTOM }),
      renderAdditionalPriceCell(renderAdditionalPriceParagraph(`${item.codigo ? `${item.codigo} ` : ''}${item.descricao || ''}`), ADDITIONAL_PRICE_COLUMNS[1], { borders: PRICE_ROW_BOTTOM }),
      renderAdditionalPriceCell(renderAdditionalPriceParagraph(item.ncm, { align: 'center' }), ADDITIONAL_PRICE_COLUMNS[2], { borders: PRICE_ROW_BOTTOM }),
      renderAdditionalPriceCell(renderAdditionalPriceParagraph(item.quant, { align: 'center' }), ADDITIONAL_PRICE_COLUMNS[3], { borders: PRICE_ROW_BOTTOM }),
      renderAdditionalPriceCell(renderAdditionalPriceParagraph(item.un, { align: 'center' }), ADDITIONAL_PRICE_COLUMNS[4], { borders: PRICE_ROW_BOTTOM }),
      renderAdditionalPriceCell(renderAccountingPriceParagraph(item.valor_unit, ADDITIONAL_PRICE_COLUMNS[5]), ADDITIONAL_PRICE_COLUMNS[5], { borders: PRICE_ROW_BOTTOM }),
      renderAdditionalPriceCell(renderAccountingPriceParagraph(item.valor_total, ADDITIONAL_PRICE_COLUMNS[6]), ADDITIONAL_PRICE_COLUMNS[6], { borders: PRICE_ROW_BOTTOM })
    ])).join('');
    const totalRow = renderAdditionalPriceRow([
      renderAdditionalPriceCell(renderAdditionalPriceParagraph(`TOTAL ${topic.titulo}:`, { bold: true, align: 'right' }), ADDITIONAL_PRICE_COLUMNS.slice(0, 6).reduce((sum, width) => sum + width, 0), { gridSpan: 6, borders: PRICE_TOTAL_TOP }),
      renderAdditionalPriceCell(renderAccountingPriceParagraph(formatCurrency(topic.totalNumber), ADDITIONAL_PRICE_COLUMNS[6], { bold: true }), ADDITIONAL_PRICE_COLUMNS[6], { borders: PRICE_TOTAL_TOP })
    ]);
    return `${titleRow}${itemRows}${totalRow}`;
  }).join('');

  const itemsTable = `${renderAdditionalPriceTableOpen(ADDITIONAL_PRICE_COLUMNS)}${header}${topicRows}</w:tbl>`;
  const grandTotal = renderRoundedHeaderRow(
    ['PRE\u00c7O TOTAL', `${totalText}${totalInWords ? ` (${totalInWords})` : ''}`],
    [2760, 7485],
    {
      shapeName: `ResumoPrecoAdicional${numberSeed}`,
      shapeNumber: 9600 + numberSeed,
      anchorId: (0xD0000000 + numberSeed).toString(16).toUpperCase().slice(-8),
      outerGridSpan: 4
    }
  );
  const termRows = [
    ['Moeda:', block.moeda || 'Real R$', 'Validade da Proposta:', block.validade_proposta || ''],
    ['Forma de Pagamento:', block.pagamento || '', 'Prazo de Entrega:', block.prazo_entrega || ''],
    ['Frete:', block.frete || '', 'Impostos:', block.impostos || '']
  ].map(([leftLabel, leftValue, rightLabel, rightValue]) => renderAdditionalPriceRow([
    renderAdditionalPriceCell(renderAdditionalPriceLabelValue(leftLabel, leftValue), 5123, {
      gridSpan: 2,
      borders: '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/></w:tcBorders>'
    }),
    renderAdditionalPriceCell(renderAdditionalPriceLabelValue(rightLabel, rightValue), 5122, { gridSpan: 2, borders: PRICE_NO_BORDERS })
  ])).join('');
  const summaryTable = `${renderAdditionalPriceTableOpen(ADDITIONAL_PRICE_SUMMARY_COLUMNS)}${grandTotal}${termRows}</w:tbl>`;
  const middleGap = '<w:p><w:pPr><w:spacing w:before="150" w:after="150" w:line="120" w:lineRule="exact"/><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr></w:pPr></w:p>';
  const bottomGap = '<w:p><w:pPr><w:spacing w:before="20" w:after="20" w:line="40" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
  return `${itemsTable}${middleGap}${summaryTable}${bottomGap}`;
}

function renderAdditionalPriceTableOpen(columns) {
  const width = columns.reduce((sum, column) => sum + column, 0);
  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="${width}" w:type="dxa"/><w:jc w:val="left"/><w:tblInd w:w="390" w:type="dxa"/><w:tblLayout w:type="fixed"/>
      <w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>
    </w:tblPr>
    <w:tblGrid>${columns.map((column) => `<w:gridCol w:w="${column}"/>`).join('')}</w:tblGrid>`;
}

function renderAdditionalPriceRow(cells) {
  return `<w:tr>${cells.join('')}</w:tr>`;
}

function renderAdditionalPriceCell(content, width, options = {}) {
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${options.gridSpan ? `<w:gridSpan w:val="${options.gridSpan}"/>` : ''}${options.borders || PRICE_NO_BORDERS}<w:tcMar><w:top w:w="35" w:type="dxa"/><w:left w:w="70" w:type="dxa"/><w:bottom w:w="35" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>${content}</w:tc>`;
}

function renderAdditionalPriceParagraph(value, options = {}) {
  const alignment = options.align ? `<w:jc w:val="${options.align}"/>` : '';
  const bold = options.bold ? '<w:b/>' : '';
  return `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>${alignment}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>${bold}<w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">${escapeXml(value ?? '')}</w:t></w:r></w:p>`;
}

function renderAccountingPriceParagraph(value, cellWidth, options = {}) {
  const formatted = String(value || formatCurrency(0)).trim();
  const match = formatted.match(/^(-?)R\$\s*(.*)$/);
  const amount = match ? `${match[1]}${match[2]}` : formatted;
  const symbol = match ? 'R$' : '';
  const bold = options.bold ? '<w:b/>' : '';
  const tabPosition = Math.max(0, Number(cellWidth || 0) - 140);
  const runProperties = `<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>${bold}<w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>`;

  return `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="${tabPosition}"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r>${runProperties}<w:t>${escapeXml(symbol)}</w:t></w:r>
    <w:r>${runProperties}<w:tab/></w:r>
    <w:r>${runProperties}<w:t>${escapeXml(amount)}</w:t></w:r>
  </w:p>`;
}

function renderAdditionalPriceLabelValue(label, value) {
  return `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:b/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t>${escapeXml(label)}</w:t></w:r>
    <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve"> ${escapeXml(value)}</w:t></w:r>
  </w:p>`;
}

function renderCustomTable(block, sectionNumber = 0) {
  const columns = Array.isArray(block.colunas)
    ? block.colunas.filter(Boolean).map((column, index) => ({
      id: String(column.id || `coluna-${index + 1}`),
      nome: String(column.nome || `Coluna ${index + 1}`)
    }))
    : [];
  if (!columns.length) return '';

  const rows = Array.isArray(block.linhas) ? block.linhas.filter(Boolean) : [];
  const tableWidth = 10245;
  const baseWidth = Math.floor(tableWidth / columns.length);
  const widths = columns.map((_, index) => (
    index === columns.length - 1 ? tableWidth - (baseWidth * index) : baseWidth
  ));
  const numberSeed = Math.max(0, Math.trunc(Number(sectionNumber) || 0));
  const headerXml = renderRoundedHeaderRow(
    columns.map((column) => column.nome.toUpperCase()),
    widths,
    {
      shapeName: `CabecalhoTabelaAdicional${numberSeed}`,
      shapeNumber: 9700 + numberSeed,
      anchorId: (0xE0000000 + numberSeed).toString(16).toUpperCase().slice(-8)
    }
  );
  const rowsXml = rows.map((row) => {
    const values = row.valores && typeof row.valores === 'object' ? row.valores : row;
    return renderDocumentListRow(columns.map((column) => values?.[column.id] ?? ''), widths);
  }).join('');

  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="${tableWidth}" w:type="dxa"/><w:tblInd w:w="390" w:type="dxa"/><w:tblLayout w:type="fixed"/>
      <w:tblBorders>
        <w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>
        <w:insideH w:val="nil"/><w:insideV w:val="nil"/>
      </w:tblBorders>
      <w:tblCellMar>
        <w:top w:w="30" w:type="dxa"/><w:left w:w="60" w:type="dxa"/>
        <w:bottom w:w="30" w:type="dxa"/><w:right w:w="60" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
    <w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
    ${headerXml}${rowsXml}
  </w:tbl><w:p><w:pPr><w:spacing w:after="80"/></w:pPr></w:p>`;
}

function escapeXml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

module.exports = {
  gerarDocx,
  prepareData,
  renderAdditionalBlocks
};

if (require.main === module) {
  const dataPath = process.argv[2] || path.resolve(__dirname, '..', '..', 'scripts', 'dados-teste.json');
  const outputPath = process.argv[3];
  const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
  const generated = gerarDocx(data, outputPath);
  console.log(`Documento gerado em ${generated}`);
}
