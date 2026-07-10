const fs = require('node:fs');
const path = require('node:path');
const PizZip = require('pizzip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

const projectRoot = path.resolve(__dirname, '..');
const sourceDocx = path.resolve(projectRoot, '..', 'SM_20260200_R01_PR_Oceanica_SubVIII.docx');
const outputDocx = path.join(projectRoot, 'src', 'templates', 'modelo_proposta.docx');

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

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

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function priceTableXml() {
  const columns = [560, 900, 3000, 850, 650, 580, 1200, 1220];
  const tableWidth = columns.reduce((sum, item) => sum + item, 0);
  const headerOptions = { size: 14, align: 'center', line: 180 };
  const itemOptions = { size: 16 };
  const sectionOptions = { bold: true, size: 16 };
  const totalLabelOptions = { size: 16, align: 'right' };
  const totalValueOptions = { bold: true, size: 16, align: 'right' };
  const grayFill = 'BFBFBF';
  const lineColor = 'BFBFBF';
  const darkLine = '404040';
  const noBorder = '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>';
  const rowBottom = `<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="${lineColor}"/><w:right w:val="nil"/></w:tcBorders>`;
  const headerBorders = `<w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="${darkLine}"/><w:left w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="${darkLine}"/><w:right w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/></w:tcBorders>`;
  const totalBorders = `<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="${lineColor}"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>`;
  const boxBorders = '<w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="404040"/><w:left w:val="single" w:sz="6" w:space="0" w:color="404040"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="404040"/><w:right w:val="single" w:sz="6" w:space="0" w:color="404040"/></w:tcBorders>';
  const noMargins = '<w:tcMar><w:top w:w="8" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:bottom w:w="8" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tcMar>';

  const tableOpen = (borders = false) => `
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="${tableWidth}" w:type="dxa"/>
    <w:jc w:val="center"/>
    <w:tblLayout w:type="fixed"/>
    ${borders ? `<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>` : ''}
    <w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>
  </w:tblPr>
  <w:tblGrid>${columns.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
`;
  const tableClose = '</w:tbl>';
  const gap = '<w:p><w:pPr><w:spacing w:before="150" w:after="150" w:line="120" w:lineRule="exact"/><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr></w:pPr></w:p>';
  const priceTopGap = '<w:p><w:pPr><w:spacing w:before="120" w:after="120" w:line="120" w:lineRule="exact"/><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr></w:pPr></w:p>';
  const priceBottomGap = '<w:p><w:pPr><w:spacing w:before="130" w:after="130" w:line="120" w:lineRule="exact"/><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr></w:pPr></w:p>';

  const header = tr([
    tc(p('ITEM', headerOptions), columns[0], { fill: grayFill, borders: headerBorders, margins: noMargins }),
    tc(p('C\u00d3DIGO', headerOptions), columns[1], { fill: grayFill, borders: headerBorders, margins: noMargins }),
    tc(p('DESCRI\u00c7\u00c3O DOS ITENS', headerOptions), columns[2], { fill: grayFill, borders: headerBorders, margins: noMargins }),
    tc(p('NCM', headerOptions), columns[3], { fill: grayFill, borders: headerBorders, margins: noMargins }),
    tc(p('QTD', headerOptions), columns[4], { fill: grayFill, borders: headerBorders, margins: noMargins }),
    tc(p('UNID.', headerOptions), columns[5], { fill: grayFill, borders: headerBorders, margins: noMargins }),
    tc(p('VALOR UNT.', headerOptions), columns[6], { fill: grayFill, borders: headerBorders, margins: noMargins }),
    tc(p('VALOR TOTAL', headerOptions), columns[7], { fill: grayFill, borders: headerBorders, margins: noMargins })
  ]);

  const section = () => tr([
    tc(p('{#topicos_preco}{titulo}', sectionOptions), tableWidth, { gridSpan: 8, borders: noBorder })
  ]);

  const itemRow = () => tr([
    tc(p('{#itens}{item}', { ...itemOptions, bold: true }), columns[0], { borders: rowBottom }),
    tc(p('{codigo}', { ...itemOptions, bold: true }), columns[1], { borders: rowBottom }),
    tc(p('{descricao}', { ...itemOptions, bold: true }), columns[2], { borders: rowBottom }),
    tc(p('{ncm}', { ...itemOptions, bold: true, align: 'center' }), columns[3], { borders: rowBottom }),
    tc(p('{quant}', { ...itemOptions, bold: true, align: 'right' }), columns[4], { borders: rowBottom }),
    tc(p('{un}', { ...itemOptions, bold: true, align: 'center' }), columns[5], { borders: rowBottom }),
    tc(p('{valor_unit}', { ...itemOptions, bold: true, align: 'right' }), columns[6], { borders: rowBottom }),
    tc(p('{valor_total}{/itens}', { ...itemOptions, bold: true, align: 'right' }), columns[7], { borders: rowBottom })
  ]);

  const totalRow = () => tr([
    tc(p('TOTAL {titulo} :', totalLabelOptions), columns.slice(0, 7).reduce((sum, item) => sum + item, 0), { gridSpan: 7, borders: totalBorders }),
    tc(p('{total}{/topicos_preco}', totalValueOptions), columns[7], { borders: totalBorders })
  ]);

  const grandTotal = tr([
    tc(p('PRE\u00c7O TOTAL', { bold: true, size: 16, align: 'center' }), columns[0] + columns[1] + columns[2], { fill: grayFill, gridSpan: 3, borders: boxBorders }),
    tc(p('{preco_total_numero} ({preco_total_extenso})', { bold: true, size: 16, align: 'center' }), columns.slice(3).reduce((sum, item) => sum + item, 0), { fill: grayFill, gridSpan: 5, borders: boxBorders })
  ]);

  const termsColumns = [4580, 4580];
  const termsTableOpen = `
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="${tableWidth}" w:type="dxa"/>
    <w:jc w:val="center"/>
    <w:tblLayout w:type="fixed"/>
    <w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>
  </w:tblPr>
  <w:tblGrid>${termsColumns.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
`;
  const terms = [
    ['Moeda:', '{moeda}', 'Validade da Proposta:', '{validade_proposta}'],
    ['Forma de Pagamento:', '{pagamento}', 'Prazo de Entrega:', '{prazo_entrega}'],
    ['Frete:', '{frete}', 'Impostos:', '{impostos}']
  ].map(([a, b, c, d]) => tr([
    tc(pLabelValue(a, b, { size: 16 }), termsColumns[0], { borders: noBorder }),
    tc(pLabelValue(c, d, { size: 16 }), termsColumns[1], { borders: noBorder })
  ])).join('');

  return `
  ${priceTopGap}
  ${tableOpen(true)}
  ${header}
  ${section()}
  ${itemRow()}
  ${totalRow()}
  ${tableClose}
  ${gap}
  ${tableOpen(true)}
  ${grandTotal}
  ${tableClose}
  ${termsTableOpen}
  ${terms}
  ${tableClose}
  ${priceBottomGap}`;
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

function main() {
  if (!fs.existsSync(sourceDocx)) {
    throw new Error(`Arquivo original não encontrado: ${sourceDocx}`);
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
  replacePriceBlock(doc);

  zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
  removeLegacyPriceObject(zip);
  fs.mkdirSync(path.dirname(outputDocx), { recursive: true });
  fs.writeFileSync(outputDocx, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(`Template criado em ${outputDocx}`);
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
