const fs = require('node:fs');
const path = require('node:path');
const PizZip = require('pizzip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

const projectRoot = path.resolve(__dirname, '..');
const sourceDocx = path.resolve(projectRoot, '..', 'SM_20260200_R01_PR_Oceanica_SubVIII.docx');
const outputDocx = path.join(projectRoot, 'src', 'templates', 'modelo_proposta.docx');

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PRICE_COLUMNS = [710, 4540, 945, 710, 590, 1315, 1435];
const SUMMARY_COLUMNS = [2760, 2363, 2362, 2760];
const PRICE_TABLE_INDENT = 390;
const PRICE_HEADER_HEIGHT = 300;
const SUMMARY_HEIGHT = 340;
const PRICE_HEADER_REL_ID = 'rIdCabecalhoPreco';
const FIXED_ADDITIONAL_INFO = [
  ['51', 'Prazo de Execução'],
  ['511', 'Serviços: Estimados em {prazo_execucao_dias}, incluindo translado ida & volta Equipe;'],
  ['512', 'Consumíveis/Materiais/Equipamentos: Imediato;'],
  ['513', 'O prazo de Execução dos serviços pode ser alterado de acordo com as condições de execução dos mesmos;'],
  ['514', 'Caso o prazo de execução seja diferente do mencionado, o valor total da presente proposta será alterado conforme as informações abaixo:'],
  ['514a', 'De Segunda-feira a Sexta-feira das 17h30min até as 7h30min 50% adicional;'],
  ['514b', 'Durante o Sábado 50% adicional o dia todo;'],
  ['514c', 'Domingos e Feriados 100% adicional o dia todo;'],
  ['514d', 'Diária Offshore quando o barco em operação ou fundeado e a equipe permanecer a bordo.'],
  ['52', 'Caso haja a necessidade de substituição de algum componente não previsto nesta proposta o mesmo será objeto de orçamento aditivo.'],
  ['53', 'Após o término do serviço será enviado uma medição, incluindo as horas de viagem, espera a bordo e a disposição, caso necessário.'],
  ['54', 'Todas as despesas com deslocamento, alimentação e estadia da equipe, caso necessário, serão por conta do cliente;'],
  ['55', 'Os Equipamentos e Ferramentas de propriedade da SUPPLY MARINE deverão ser devolvidos no prazo máximo de 03 (três) dias após a conclusão dos serviços. Caso contrário, a SUPPLY MARINE cobrará pelos custos de cessão dos mesmos conforme tabela abaixo:']
];

function attr(name) {
  return `w:${name}`;
}

function getText(node) {
  const texts = Array.from(node.getElementsByTagName('w:t'));
  return texts.map((item) => item.textContent).join('');
}

function createTextRun(doc, text, runProperties) {
  const run = doc.createElementNS(W_NS, 'w:r');

  if (runProperties) {
    run.appendChild(runProperties.cloneNode(true));
  }

  const textNode = doc.createElementNS(W_NS, 'w:t');
  if (/^\s|\s$/.test(text)) {
    textNode.setAttribute('xml:space', 'preserve');
  }
  textNode.appendChild(doc.createTextNode(text));
  run.appendChild(textNode);
  return run;
}

function setParagraphText(doc, paragraph, text) {
  const paragraphProperties = Array.from(paragraph.childNodes).find((node) => node.nodeName === 'w:pPr');
  const runProperties = paragraphProperties
    ? Array.from(paragraphProperties.getElementsByTagName('w:rPr'))[0]
    : Array.from(paragraph.getElementsByTagName('w:rPr'))[0];

  Array.from(paragraph.childNodes).forEach((child) => {
    if (child.nodeName !== 'w:pPr') {
      paragraph.removeChild(child);
    }
  });

  paragraph.appendChild(createTextRun(doc, text, runProperties));
}

function findParagraphs(doc, text) {
  return Array.from(doc.getElementsByTagName('w:p')).filter((paragraph) => getText(paragraph) === text);
}

function replaceFirstParagraph(doc, from, to) {
  const paragraph = findParagraphs(doc, from)[0];
  if (!paragraph) {
    throw new Error(`Texto não encontrado no modelo: ${from}`);
  }
  setParagraphText(doc, paragraph, to);
}

function p(text, options = {}) {
  const bold = options.bold ? '<w:b/>' : '';
  const color = options.color ? `<w:color w:val="${options.color}"/>` : '';
  const align = options.align ? `<w:jc w:val="${options.align}"/>` : '';
  const line = options.line || 240;
  const spacing = options.spacing === false ? '' : `<w:spacing w:after="0" w:line="${line}" w:lineRule="auto"/>`;
  const escaped = escapeXml(text);
  const xmlSpace = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';

  return `<w:p><w:pPr>${spacing}${align}<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>${bold}${color}<w:sz w:val="${options.size || 16}"/><w:szCs w:val="${options.size || 16}"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>${bold}${color}<w:sz w:val="${options.size || 16}"/><w:szCs w:val="${options.size || 16}"/></w:rPr><w:t${xmlSpace}>${escaped}</w:t></w:r></w:p>`;
}

function pLabelValue(label, value, options = {}) {
  const align = options.align ? `<w:jc w:val="${options.align}"/>` : '';
  const spacing = options.spacing === false ? '' : '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>';
  const size = options.size || 16;
  return `<w:p><w:pPr>${spacing}${align}<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:b/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr><w:t>${escapeXml(label)}</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr><w:t xml:space="preserve"> ${escapeXml(value)}</w:t></w:r></w:p>`;
}

function tc(content, width, options = {}) {
  const fill = options.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${options.fill}"/>` : '';
  const span = options.gridSpan ? `<w:gridSpan w:val="${options.gridSpan}"/>` : '';
  const borders = options.borders || '';
  const valign = '<w:vAlign w:val="center"/>';
  const margins = options.margins || '<w:tcMar><w:top w:w="35" w:type="dxa"/><w:left w:w="70" w:type="dxa"/><w:bottom w:w="35" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tcMar>';
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${span}${fill}${borders}${valign}${margins}</w:tcPr>${content}</w:tc>`;
}

function tr(cells) {
  return `<w:tr>${cells.join('')}</w:tr>`;
}

function priceHeaderDrawingXml(tableWidth) {
  const widthEmu = tableWidth * 635;
  const heightEmu = PRICE_HEADER_HEIGHT * 635;

  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${PRICE_HEADER_HEIGHT}" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${widthEmu}" cy="${heightEmu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="9001" name="Cabe\u00e7alho de pre\u00e7os"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="9001" name="cabecalho-preco.svg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${PRICE_HEADER_REL_ID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

function priceHeaderRowXml(tableWidth) {
  const noBorder = '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>';
  const noMargins = '<w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>';
  return `<w:tr><w:trPr><w:trHeight w:val="${PRICE_HEADER_HEIGHT}" w:hRule="exact"/></w:trPr>${tc(priceHeaderDrawingXml(tableWidth), tableWidth, { gridSpan: PRICE_COLUMNS.length, borders: noBorder, margins: noMargins })}</w:tr>`;
}

function priceHeaderSvg() {
  const width = 2256;
  const height = 60;
  const scale = width / PRICE_COLUMNS.reduce((sum, item) => sum + item, 0);
  const labels = ['ITEM', 'DESCRI\u00c7\u00c3O DOS ITENS', 'NCM', 'QTD', 'UNID.', 'VALOR UNIT\u00c1RIO', 'VALOR TOTAL'];
  let offset = 0;
  const dividers = [];
  const text = labels.map((label, index) => {
    const columnWidth = PRICE_COLUMNS[index] * scale;
    const center = offset + (columnWidth / 2);
    offset += columnWidth;
    if (index < labels.length - 1) {
      dividers.push(`<line x1="${offset.toFixed(2)}" y1="1.5" x2="${offset.toFixed(2)}" y2="58.5"/>`);
    }
    return `<text x="${center.toFixed(2)}" y="37" text-anchor="middle">${label}</text>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect x="1.5" y="1.5" width="${width - 3}" height="${height - 3}" rx="10" fill="#BFBFBF" stroke="#404040" stroke-width="3"/>
  <g fill="none" stroke="#7F7F7F" stroke-width="2">${dividers.join('')}</g>
  <g fill="#000000" font-family="Calibri, Arial, sans-serif" font-size="26" font-weight="400">${text}</g>
</svg>`;
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function priceTableXml() {
  const columns = PRICE_COLUMNS;
  const tableWidth = columns.reduce((sum, item) => sum + item, 0);
  const itemOptions = { size: 16, align: 'center' };
  const sectionOptions = { bold: true, size: 16 };
  const totalLabelOptions = { size: 16, align: 'right' };
  const totalValueOptions = { bold: true, size: 16, align: 'right' };
  const lineColor = 'BFBFBF';
  const noBorder = '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>';
  const rowBottom = `<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="${lineColor}"/><w:right w:val="nil"/></w:tcBorders>`;
  const totalBorders = `<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="${lineColor}"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>`;

  const tableOpen = (gridColumns, borders = false, indented = true) => `
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="${gridColumns.reduce((sum, width) => sum + width, 0)}" w:type="dxa"/>
    <w:jc w:val="left"/>
    ${indented ? `<w:tblInd w:w="${PRICE_TABLE_INDENT}" w:type="dxa"/>` : '<w:tblInd w:w="0" w:type="dxa"/>'}
    <w:tblLayout w:type="fixed"/>
    ${borders ? `<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>` : ''}
    <w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>
  </w:tblPr>
  <w:tblGrid>${gridColumns.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
`;
  const tableClose = '</w:tbl>';
  const gap = '<w:p><w:pPr><w:spacing w:before="150" w:after="150" w:line="120" w:lineRule="exact"/><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr></w:pPr></w:p>';
  const priceTopGap = '<w:p><w:pPr><w:spacing w:before="120" w:after="120" w:line="120" w:lineRule="exact"/><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr></w:pPr></w:p>';
  const priceBottomGap = '<w:p><w:pPr><w:spacing w:before="130" w:after="130" w:line="120" w:lineRule="exact"/><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr></w:pPr></w:p>';

  const header = priceHeaderRowXml(tableWidth);

  const section = () => tr([
    tc(p('{#topicos_preco}{titulo}', sectionOptions), tableWidth, { gridSpan: 7, borders: noBorder })
  ]);

  const itemRow = () => tr([
    tc(p('{#itens}{item}', { ...itemOptions, bold: true }), columns[0], { borders: rowBottom }),
    tc(p('{codigo} {descricao}', { ...itemOptions, bold: true, align: 'left' }), columns[1], { borders: rowBottom }),
    tc(p('{ncm}', { ...itemOptions, bold: true }), columns[2], { borders: rowBottom }),
    tc(p('{quant}', { ...itemOptions, bold: true }), columns[3], { borders: rowBottom }),
    tc(p('{un}', { ...itemOptions, bold: true }), columns[4], { borders: rowBottom }),
    tc(p('{valor_unit}', { ...itemOptions, bold: true }), columns[5], { borders: rowBottom }),
    tc(p('{valor_total}{/itens}', { ...itemOptions, bold: true }), columns[6], { borders: rowBottom })
  ]);

  const totalRow = () => tr([
    tc(p('TOTAL {titulo} :', totalLabelOptions), columns.slice(0, 6).reduce((sum, item) => sum + item, 0), { gridSpan: 6, borders: totalBorders }),
    tc(p('{total}{/topicos_preco}', totalValueOptions), columns[6], { borders: totalBorders })
  ]);

  const grandDivider = '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="single" w:sz="6" w:space="0" w:color="404040"/></w:tcBorders>';
  const grandContent = `${tableOpen(SUMMARY_COLUMNS, true, false)}${tr([
    tc(p('PRE\u00c7O TOTAL', { bold: true, size: 16, align: 'center' }), SUMMARY_COLUMNS[0], { borders: grandDivider }),
    tc(p('{preco_total_numero} ({preco_total_extenso})', { bold: true, size: 16, align: 'center' }), SUMMARY_COLUMNS.slice(1).reduce((sum, item) => sum + item, 0), { gridSpan: 3, borders: noBorder })
  ])}${tableClose}`;
  const grandShape = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${SUMMARY_HEIGHT}" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr><w:r><w:pict><v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="ResumoPrecoArredondado" arcsize="12%" fillcolor="#BFBFBF" strokecolor="#404040" strokeweight="0.75pt" o:allowincell="t" style="width:${tableWidth / 20}pt;height:${SUMMARY_HEIGHT / 20}pt"><v:textbox inset="0,0,0,0"><w:txbxContent>${grandContent}</w:txbxContent></v:textbox></v:roundrect></w:pict></w:r></w:p>`;
  const grandTotal = `<w:tr><w:trPr><w:trHeight w:val="${SUMMARY_HEIGHT}" w:hRule="exact"/></w:trPr>${tc(grandShape, tableWidth, { gridSpan: 4, borders: noBorder, margins: '<w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>' })}</w:tr>`;

  const termsWidth = SUMMARY_COLUMNS[0] + SUMMARY_COLUMNS[1];
  const termsLeftBorder = '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/></w:tcBorders>';
  const terms = [
    ['Moeda:', '{moeda}', 'Validade da Proposta:', '{validade_proposta}'],
    ['Forma de Pagamento:', '{pagamento}', 'Prazo de Entrega:', '{prazo_entrega}'],
    ['Frete:', '{frete}', 'Impostos:', '{impostos}']
  ].map(([a, b, c, d]) => tr([
    tc(pLabelValue(a, b, { size: 16 }), termsWidth, { gridSpan: 2, borders: termsLeftBorder }),
    tc(pLabelValue(c, d, { size: 16 }), termsWidth, { gridSpan: 2, borders: noBorder })
  ])).join('');

  return `
  ${priceTopGap}
  ${tableOpen(columns, true)}
  ${header}
  ${section()}
  ${itemRow()}
  ${totalRow()}
  ${tableClose}
  ${gap}
  ${tableOpen(SUMMARY_COLUMNS, true)}
  ${grandTotal}
  ${terms}
  ${tableClose}
  ${priceBottomGap}`;
}

function addPriceHeaderImage(zip) {
  zip.file('word/media/cabecalho-preco.svg', priceHeaderSvg());

  const contentTypesPath = '[Content_Types].xml';
  const contentTypesDoc = new DOMParser().parseFromString(zip.file(contentTypesPath).asText(), 'text/xml');
  const contentTypesRoot = contentTypesDoc.documentElement;
  const hasSvgContentType = Array.from(contentTypesDoc.getElementsByTagName('Default'))
    .some((node) => node.getAttribute('Extension') === 'svg');
  if (!hasSvgContentType) {
    const defaultNode = contentTypesDoc.createElementNS(
      'http://schemas.openxmlformats.org/package/2006/content-types',
      'Default'
    );
    defaultNode.setAttribute('Extension', 'svg');
    defaultNode.setAttribute('ContentType', 'image/svg+xml');
    contentTypesRoot.appendChild(defaultNode);
  }
  zip.file(contentTypesPath, new XMLSerializer().serializeToString(contentTypesDoc));

  const relsPath = 'word/_rels/document.xml.rels';
  const relsDoc = new DOMParser().parseFromString(zip.file(relsPath).asText(), 'text/xml');
  const relsRoot = relsDoc.documentElement;
  Array.from(relsDoc.getElementsByTagName('Relationship'))
    .filter((node) => node.getAttribute('Id') === PRICE_HEADER_REL_ID || node.getAttribute('Target') === 'media/cabecalho-preco.svg')
    .forEach((node) => relsRoot.removeChild(node));
  const relationship = relsDoc.createElementNS(
    'http://schemas.openxmlformats.org/package/2006/relationships',
    'Relationship'
  );
  relationship.setAttribute('Id', PRICE_HEADER_REL_ID);
  relationship.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
  relationship.setAttribute('Target', 'media/cabecalho-preco.svg');
  relsRoot.appendChild(relationship);
  zip.file(relsPath, new XMLSerializer().serializeToString(relsDoc));
}

function adjustExistingPriceTable(doc) {
  const tables = Array.from(doc.getElementsByTagName('w:tbl'));
  const priceTable = tables.find((candidate) => {
    const text = getText(candidate);
    const hasOriginalHeader = text.includes('DESCRI\u00c7\u00c3O DOS ITENS') && text.includes('VALOR TOTAL');
    const hasPriceLoops = text.includes('{#topicos_') && text.includes('valor_');
    return hasOriginalHeader || hasPriceLoops;
  });

  const summaryTable = tables.find((candidate) => getText(candidate).includes('PRE\u00c7O TOTAL'));
  if (!priceTable || !summaryTable) {
    throw new Error('Tabela de pre\u00e7os n\u00e3o encontrada no template existente.');
  }

  const wrapper = new DOMParser().parseFromString(`<root xmlns:w="${W_NS}">${priceTableXml()}</root>`, 'text/xml');
  const generatedTables = Array.from(wrapper.getElementsByTagName('w:tbl'));
  priceTable.parentNode.replaceChild(doc.importNode(generatedTables[0], true), priceTable);
  summaryTable.parentNode.replaceChild(doc.importNode(generatedTables[1], true), summaryTable);

  tables
    .filter((candidate) => candidate !== summaryTable && getText(candidate).includes('Validade da Proposta:'))
    .forEach((candidate) => candidate.parentNode?.removeChild(candidate));
}
function replacePriceBlock(doc) {
  const body = doc.getElementsByTagName('w:body')[0];
  const children = Array.from(body.childNodes);
  const priceIndex = children.findIndex((node) => getText(node) === 'PREÇO');
  const nextIndex = children.findIndex((node, index) => index > priceIndex && getText(node) === 'INFORMAÇÕES ADICIONAIS');

  if (priceIndex < 0 || nextIndex < 0) {
    throw new Error('Não foi possível localizar o bloco PREÇO no documento original.');
  }

  const priceParagraph = children[priceIndex];
  const infoParagraph = children[nextIndex];
  [priceParagraph, infoParagraph].forEach((paragraph) => {
    const numId = Array.from(paragraph.getElementsByTagName('w:numId'))[0];
    if (numId) {
      numId.setAttribute(attr('val'), '1');
    }
  });

  for (let index = nextIndex - 1; index > priceIndex + 1; index -= 1) {
    body.removeChild(children[index]);
  }

  const wrapper = new DOMParser().parseFromString(`<root xmlns:w="${W_NS}">${priceTableXml()}</root>`, 'text/xml');
  const insertionPoint = body.childNodes[priceIndex + 2];
  Array.from(wrapper.documentElement.childNodes)
    .filter((node) => node.nodeType === 1)
    .forEach((node) => {
      body.insertBefore(doc.importNode(node, true), insertionPoint);
    });
}

function setServiceLoop(doc) {
  const first = findParagraphs(doc, 'Instalação do compressor do ar-condicionado A2 de BB junto com a VC serviços marítimos.')[0];
  const second = findParagraphs(doc, 'Verificação da tubulação do compressor B1 de BE com vazamento junto com a VC serviços marítimos')[0];

  if (!first || !second) {
    throw new Error('Não foi possível localizar os parágrafos de descrição de serviços.');
  }

  const end = second.cloneNode(true);
  setParagraphText(doc, first, '{#servicos_descricao}');
  setParagraphText(doc, second, '{item}');
  setParagraphText(doc, end, '{/servicos_descricao}');
  second.parentNode.insertBefore(end, second.nextSibling);
}

function setTechnicalTeamLoop(doc) {
  if (findParagraphs(doc, '{#equipe_tecnica_itens}').length) {
    return;
  }

  const itemParagraph = findParagraphs(doc, '{equipe_tecnica}')[0];
  if (!itemParagraph) {
    throw new Error('Não foi possível localizar o campo de equipe técnica no template.');
  }

  const start = itemParagraph.cloneNode(true);
  const end = itemParagraph.cloneNode(true);
  setParagraphText(doc, start, '{#equipe_tecnica_itens}');
  setParagraphText(doc, itemParagraph, '{item}');
  setParagraphText(doc, end, '{/equipe_tecnica_itens}');
  itemParagraph.parentNode.insertBefore(start, itemParagraph);
  itemParagraph.parentNode.insertBefore(end, itemParagraph.nextSibling);
}

function setAdditionalInfoLoop(doc) {
  if (findParagraphs(doc, '{#informacoes_adicionais}').length) {
    return;
  }

  const anchorText = 'Os Equipamentos e Ferramentas de propriedade da SUPPLY MARINE deverão ser devolvidos no prazo máximo de 03 (três) dias após a conclusão dos serviços. Caso contrário, a SUPPLY MARINE cobrará pelos custos de cessão dos mesmos conforme tabela abaixo:';
  const anchor = findParagraphs(doc, anchorText)[0];
  if (!anchor) {
    throw new Error('Não foi possível localizar o último item de informações adicionais no template.');
  }

  const start = anchor.cloneNode(true);
  const item = anchor.cloneNode(true);
  const end = anchor.cloneNode(true);
  setParagraphText(doc, start, '{#informacoes_adicionais}');
  setParagraphText(doc, item, '{item}');
  setParagraphText(doc, end, '{/informacoes_adicionais}');
  anchor.parentNode.insertBefore(end, anchor.nextSibling);
  anchor.parentNode.insertBefore(item, end);
  anchor.parentNode.insertBefore(start, item);
}

function setEditableFixedAdditionalInfo(doc) {
  if (findParagraphs(doc, '{#mostrar_info_51}').length) {
    return;
  }

  FIXED_ADDITIONAL_INFO.forEach(([id, originalText]) => {
    const paragraph = findParagraphs(doc, originalText)[0];
    if (!paragraph) {
      throw new Error(`Não foi possível localizar o item fixo de informações adicionais: ${id}`);
    }

    const start = paragraph.cloneNode(true);
    const end = paragraph.cloneNode(true);
    setParagraphText(doc, start, `{#mostrar_info_${id}}`);
    setParagraphText(doc, paragraph, `{texto_info_${id}}`);
    setParagraphText(doc, end, `{/mostrar_info_${id}}`);
    paragraph.parentNode.insertBefore(start, paragraph);
    paragraph.parentNode.insertBefore(end, paragraph.nextSibling);
  });
}

function main() {
  if (process.argv.includes('--backup')) {
    const backupPath = backupExistingTemplate();
    if (!backupPath) {
      throw new Error(`Template não encontrado: ${outputDocx}`);
    }
    console.log(`Backup do template criado em ${backupPath}`);
    return;
  }

  if (process.argv.includes('--ajustar-existente')) {
    adjustExistingTemplate();
    return;
  }

  if (process.argv.includes('--ajustar-informacoes')) {
    adjustAdditionalInfoTemplate();
    return;
  }

  if (fs.existsSync(outputDocx) && !process.argv.includes('--recriar')) {
    console.log(`Template existente mantido sem alterações: ${outputDocx}`);
    console.log('Para recriar a partir do documento original, use --recriar.');
    return;
  }

  if (!fs.existsSync(sourceDocx)) {
    throw new Error(`Arquivo original não encontrado: ${sourceDocx}`);
  }

  const backupPath = backupExistingTemplate();
  if (backupPath) {
    console.log(`Backup do template atual criado em ${backupPath}`);
  }

  const zip = new PizZip(fs.readFileSync(sourceDocx));
  const xml = zip.file('word/document.xml').asText();
  const doc = new DOMParser().parseFromString(xml, 'text/xml');

  const replacements = new Map([
    ['Oceânica Engenharia E Consultoria Ltda', '{empresa_cliente}'],
    ['Oceânica Sub VIII', '{unidade}'],
    ['Fluig 626498', '{processo_fluig}'],
    ['Daniel Alves Fernandes | Supervisor de Manutenção Naval', '{solicitante_nome_cargo}'],
    ['daniel.fernandes@oceanica.com.br', '{contato_email}'],
    ['(+55 21) 2139-4250', '{contato_telefone}'],
    ['20260200', '{numero_documento}'],
    ['20/03/2026', '{data_documento}'],
    ['André Luis Souza', '{responsavel_nome}'],
    ['alsouza@supplymarine.com.br', '{responsavel_email}'],
    ['(+55 21) 2596-6262', '{responsavel_telefone}'],
    ['Manutenção do ar condicionado de BB & BE', '{objeto}'],
    ['01 (um) Técnico HVAC -R | 01 (um) Assistente Técnico.', '{equipe_tecnica}'],
    ['Maceió | AL | 13/03/2026.', '{local_servico} | {data_servico}.'],
    ['Serviços: Estimados em 03 (três) dias, incluindo translado ida & volta Equipe;', 'Serviços: Estimados em {prazo_execucao_dias}, incluindo translado ida & volta Equipe;'],
    ['André Luis Santos De Souza', '{responsavel_nome}']
  ]);

  replacements.forEach((to, from) => replaceFirstParagraph(doc, from, to));
  setServiceLoop(doc);
  setTechnicalTeamLoop(doc);
  setAdditionalInfoLoop(doc);
  setEditableFixedAdditionalInfo(doc);
  replacePriceBlock(doc);

  zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
  addPriceHeaderImage(zip);
  removeLegacyPriceObject(zip);
  fs.mkdirSync(path.dirname(outputDocx), { recursive: true });
  fs.writeFileSync(outputDocx, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(`Template criado em ${outputDocx}`);
}

function adjustExistingTemplate() {
  if (!fs.existsSync(outputDocx)) {
    throw new Error(`Template n\u00e3o encontrado: ${outputDocx}`);
  }

  const backupPath = backupExistingTemplate();

  const zip = new PizZip(fs.readFileSync(outputDocx));
  const documentFile = zip.file('word/document.xml');
  if (!documentFile) {
    throw new Error('Template existente sem word/document.xml.');
  }

  const doc = new DOMParser().parseFromString(documentFile.asText(), 'text/xml');
  adjustExistingPriceTable(doc);
  setTechnicalTeamLoop(doc);
  setAdditionalInfoLoop(doc);
  setEditableFixedAdditionalInfo(doc);
  zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
  addPriceHeaderImage(zip);
  fs.writeFileSync(outputDocx, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(`Cabe\u00e7alho do template ajustado em ${outputDocx}`);
  console.log(`Versão anterior preservada em ${backupPath}`);
}

function adjustAdditionalInfoTemplate() {
  if (!fs.existsSync(outputDocx)) {
    throw new Error(`Template não encontrado: ${outputDocx}`);
  }

  const backupPath = backupExistingTemplate();
  const zip = new PizZip(fs.readFileSync(outputDocx));
  const documentFile = zip.file('word/document.xml');
  if (!documentFile) {
    throw new Error('Template existente sem word/document.xml.');
  }

  const doc = new DOMParser().parseFromString(documentFile.asText(), 'text/xml');
  setAdditionalInfoLoop(doc);
  setEditableFixedAdditionalInfo(doc);
  zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
  fs.writeFileSync(outputDocx, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(`Informações adicionais habilitadas em ${outputDocx}`);
  console.log(`Versão anterior preservada em ${backupPath}`);
}

function backupExistingTemplate() {
  if (!fs.existsSync(outputDocx)) {
    return '';
  }

  const backupDir = path.join(projectRoot, '.work', 'template-backups');
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const backupPath = path.join(backupDir, `modelo_proposta-${timestamp}.docx`);
  fs.mkdirSync(backupDir, { recursive: true });
  fs.copyFileSync(outputDocx, backupPath);
  return backupPath;
}

function removeLegacyPriceObject(zip) {
  zip.remove('word/embeddings/Microsoft_Excel_Worksheet.xlsx');
  zip.remove('word/media/image1.emf');

  const relsPath = 'word/_rels/document.xml.rels';
  const relsDoc = new DOMParser().parseFromString(zip.file(relsPath).asText(), 'text/xml');
  Array.from(relsDoc.getElementsByTagName('Relationship')).forEach((relationship) => {
    const target = relationship.getAttribute('Target');
    if (target === 'embeddings/Microsoft_Excel_Worksheet.xlsx' || target === 'media/image1.emf') {
      relationship.parentNode.removeChild(relationship);
    }
  });
  zip.file(relsPath, new XMLSerializer().serializeToString(relsDoc));
}

main();
