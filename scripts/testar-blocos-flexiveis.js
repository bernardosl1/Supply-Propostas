const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const PizZip = require('pizzip');
const { DOMParser } = require('@xmldom/xmldom');
const { gerarDocx } = require('../src/lib/gerarDocx');

const base = JSON.parse(fs.readFileSync(path.join(__dirname, 'dados-teste.json'), 'utf8'));
const tests = [
  { name: 'legacy', data: base },
  {
    name: 'without-location',
    data: { ...base, local_servico: '', data_servico: '' }
  },
  {
    name: 'without-service-date',
    data: { ...base, data_servico: '' }
  },
  {
    name: 'excluded-object',
    data: { ...base, secoes_excluidas: ['objeto'] }
  },
  {
    name: 'excluded-scope',
    data: { ...base, secoes_excluidas: ['escopo'] }
  },
  {
    name: 'excluded-object-and-scope',
    data: { ...base, secoes_excluidas: ['objeto', 'escopo'] }
  },
  {
    name: 'flexible',
    data: {
      ...base,
      blocos_adicionais: [
        {
          id: 'texto-1',
          tipo: 'texto',
          titulo: 'CRONOGRAMA',
          observacoes: ['Execu\u00e7\u00e3o em duas etapas.', 'Sujeito a janela operacional.'],
          subtopicos: [
            {
              id: 'subtopico-1',
              titulo: 'Etapas do servi\u00e7o',
              observacoes: ['Mobiliza\u00e7\u00e3o', 'Execu\u00e7\u00e3o'],
              subtopicos: [
                {
                  id: 'subtopico-1-1',
                  titulo: 'Detalhamento da etapa',
                  observacoes: ['Inspe\u00e7\u00e3o inicial']
                }
              ]
            }
          ]
        },
        {
          id: 'lista-1',
          tipo: 'lista',
          linhas: [
            {
              descricao: 'Solicita\u00e7\u00e3o de cota\u00e7\u00e3o',
              numero_documento: 'SC35085 | RC73960 | Cota\u00e7\u00e3o 167383',
              data: '14/07/2026'
            }
          ]
        },
        {
          id: 'preco-1',
          tipo: 'preco',
          topicos_preco: [
            {
              titulo: 'LOCA\u00c7\u00c3O',
              tipo: 'personalizado',
              itens: [
                { item: '0001', descricao: 'Equipamento especial', ncm: '8424', quant: 2, un: 'UN', valor_unit: 1250, valor_total: 2500 }
              ]
            }
          ],
          preco_total_numero: 2500,
          preco_total_extenso: 'Dois mil e quinhentos reais',
          moeda: 'Real R$',
          validade_proposta: '30/08/2026',
          pagamento: '15 dias',
          prazo_entrega: 'Imediato',
          frete: 'CIF',
          impostos: 'inclusos no pre\u00e7o'
        },
        {
          id: 'tabela-1',
          tipo: 'tabela',
          titulo: 'EQUIPAMENTOS',
          colunas: [
            { id: 'equipamento', nome: 'Equipamento' },
            { id: 'modelo', nome: 'Modelo' },
            { id: 'quantidade', nome: 'Quantidade' }
          ],
          linhas: [
            { valores: { equipamento: 'Bomba A', modelo: 'XP-10', quantidade: '2' } },
            { valores: { equipamento: 'Mangueira', modelo: 'M-20', quantidade: '4' } }
          ]
        },
        { id: 'pagina-1', tipo: 'quebra_pagina' },
        {
          id: 'texto-2',
          tipo: 'texto',
          titulo: 'OBSERVA\u00c7\u00d5ES',
          conteudo: 'Conte\u00fado ap\u00f3s quebra.'
        }
      ]
    }
  }
];

tests.push({
  name: 'reordered',
  data: {
    ...tests.find((test) => test.name === 'flexible').data,
    ordem_secoes: [
      'dados_comerciais',
      'objeto',
      'flex:texto-1',
      'escopo',
      'flex:lista-1',
      'preco',
      'flex:preco-1',
      'flex:tabela-1',
      'flex:pagina-1',
      'flex:texto-2'
    ]
  }
});

for (const test of tests) {
  const outputPath = path.join(os.tmpdir(), `supply-${test.name}-${process.pid}.docx`);
  try {
    gerarDocx(test.data, outputPath);
    const zip = new PizZip(fs.readFileSync(outputPath));
    const xml = zip.file('word/document.xml').asText();
    const xmlErrors = [];
    new DOMParser({
      onError: (level, message) => {
        if (level !== 'warning') xmlErrors.push(message);
      }
    }).parseFromString(xml, 'application/xml');
    assert(xmlErrors.length === 0, `${test.name}: XML inv\u00e1lido (${xmlErrors.join('; ')})`);
    assert(xml.includes(base.empresa_cliente), `${test.name}: dados comerciais ausentes`);
    assert(!xml.includes('INFORMA\u00c7\u00d5ES ADICIONAIS'), `${test.name}: se\u00e7\u00e3o fixa de informa\u00e7\u00f5es adicionais ainda presente`);
    assert(!xml.includes('Prazo de Execu\u00e7\u00e3o'), `${test.name}: conte\u00fado fixo de informa\u00e7\u00f5es adicionais ainda presente`);
    if (test.name === 'legacy') {
      assert(!xml.includes('DESCRI\u00c7\u00c3O DOS ITENS'), 'Tabela fixa de pre\u00e7o ainda presente no documento');
      assert(!xml.includes('PRE\u00c7O TOTAL'), 'Resumo fixo de pre\u00e7o ainda presente no documento');
    }
    if (test.name === 'without-location') {
      assert(!xml.includes('Local e Data dos Servi\u00e7os'), 'Campo vazio de local e data ainda presente no documento');
      assert(!xml.includes(' | .'), 'Separador vazio de local e data ainda presente no documento');
    }
    if (test.name === 'without-service-date') {
      assert(xml.includes('Local e Data dos Servi\u00e7os'), 'T\u00edtulo de local removido mesmo com local preenchido');
      assert(xml.includes('Macei\u00f3 | AL.'), 'Local isolado n\u00e3o foi formatado corretamente');
      assert(!xml.includes('Macei\u00f3 | AL | .'), 'Separador vazio de data ainda presente no documento');
    }

    if (test.name === 'excluded-object' || test.name === 'excluded-object-and-scope') {
      assert(!xml.includes('OBJETO'), 'Bloco Objeto exclu\u00eddo ainda aparece no documento');
      assert(!xml.includes(String(base.objeto)), 'Conte\u00fado antigo do Objeto reapareceu no documento');
    }
    if (test.name === 'excluded-scope' || test.name === 'excluded-object-and-scope') {
      assert(!xml.includes('ESCOPO DE FORNECIMENTO'), 'Bloco Escopo exclu\u00eddo ainda aparece no documento');
      assert(!xml.includes('Descri\u00e7\u00e3o dos Servi\u00e7os'), 'Subse\u00e7\u00e3o do Escopo exclu\u00eddo ainda aparece');
    }

    if (test.name === 'flexible') {
      ['4.  CRONOGRAMA', '4.1  Etapas do servi\u00e7o', '4.1.1  Detalhamento da etapa', 'Inspe\u00e7\u00e3o inicial', 'Mobiliza\u00e7\u00e3o', '5.  DOCUMENTA\u00c7\u00c3O', 'DESCRI\u00c7\u00c3O', 'N\u00ba DO DOCUMENTO', 'Solicita\u00e7\u00e3o de cota\u00e7\u00e3o', 'SC35085 | RC73960 | Cota\u00e7\u00e3o 167383', '14/07/2026', '6.  PRE\u00c7O', 'LOCA\u00c7\u00c3O', 'Equipamento especial', 'R$1.250,00', 'R$ 2.500,00 (Dois mil e quinhentos reais)', 'Validade da Proposta:', '30/08/2026', '7.  EQUIPAMENTOS', 'Bomba A', '8.  OBSERVA\u00c7\u00d5ES', 'Conte\u00fado ap\u00f3s quebra.'].forEach((value) => {
        assert(xml.includes(value), `Conte\u00fado flex\u00edvel ausente: ${value}`);
      });
      assert(xml.indexOf('Bomba A') < xml.indexOf('Atenciosamente'), 'Blocos inseridos depois da assinatura');
      const signatureGap = xml.slice(xml.indexOf('Conte\u00fado ap\u00f3s quebra.'), xml.indexOf('Atenciosamente'));
      assert(signatureGap.includes('<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'), 'Linha vazia antes da assinatura ausente');
      assert(xml.includes('w:type="page"'), 'Quebra de p\u00e1gina ausente');
      const documentationIndex = xml.indexOf('Solicita\u00e7\u00e3o de cota\u00e7\u00e3o');
      const documentationIndentIndex = xml.lastIndexOf('<w:tblInd w:w="390" w:type="dxa"/>', documentationIndex);
      const documentationTableStart = xml.lastIndexOf('<w:tbl>', documentationIndentIndex);
      const documentationTableEnd = xml.indexOf('6.  PRE\u00c7O', documentationIndex);
      const documentationTable = xml.slice(documentationTableStart, documentationTableEnd);
      assert(documentationTable.includes('<w:tblW w:w="10245" w:type="dxa"/>'), 'Documenta\u00e7\u00e3o n\u00e3o ocupa a largura padr\u00e3o');
      assert(documentationTable.includes('<w:tblInd w:w="390" w:type="dxa"/>'), 'Documenta\u00e7\u00e3o desalinhada da tabela de pre\u00e7os');
      assert(documentationTable.includes('<v:roundrect'), 'Cabe\u00e7alho arredondado da documenta\u00e7\u00e3o ausente');
      assert(documentationTable.includes('arcsize="7864f"'), 'Arredondamento do cabe\u00e7alho da documenta\u00e7\u00e3o incorreto');
      assert(documentationTable.includes('fillcolor="#bfbfbf"'), 'Cabe\u00e7alho da documenta\u00e7\u00e3o com cor incorreta');
      assert(documentationTable.includes('strokecolor="#404040"'), 'Contorno do cabe\u00e7alho da documenta\u00e7\u00e3o incorreto');
      assert(documentationTable.includes('width:512.25pt;height:15.1pt'), 'Tamanho do cabe\u00e7alho da documenta\u00e7\u00e3o incorreto');
      assert(documentationTable.includes('w:color="B7B7B7"'), 'Linhas da documenta\u00e7\u00e3o com cor incorreta');
      assert(documentationTable.includes('w:ascii="Calibri"'), 'Fonte padr\u00e3o da documenta\u00e7\u00e3o ausente');
      assert(documentationTable.includes('<w:sz w:val="16"/>'), 'Tamanho padr\u00e3o da fonte da documenta\u00e7\u00e3o ausente');
      const additionalPriceIndex = xml.indexOf('Equipamento especial');
      const additionalPriceHeaderIndex = xml.lastIndexOf('DESCRI\u00c7\u00c3O DOS ITENS', additionalPriceIndex);
      const additionalPriceHeaderStart = xml.lastIndexOf('<v:roundrect', additionalPriceHeaderIndex);
      const additionalPriceHeader = xml.slice(additionalPriceHeaderStart, xml.indexOf('</v:roundrect>', additionalPriceHeaderIndex));
      assert(additionalPriceHeader.includes('width:512.25pt;height:15.1pt'), 'Cabe\u00e7alho da se\u00e7\u00e3o adicional de pre\u00e7o com tamanho incorreto');
      assert(additionalPriceHeader.includes('fillcolor="#bfbfbf"'), 'Cabe\u00e7alho da se\u00e7\u00e3o adicional de pre\u00e7o com cor incorreta');
      const additionalItemParagraphStart = Math.max(
        xml.lastIndexOf('<w:p>', additionalPriceIndex),
        xml.lastIndexOf('<w:p ', additionalPriceIndex)
      );
      const additionalItemParagraph = xml.slice(additionalItemParagraphStart, xml.indexOf('</w:p>', additionalPriceIndex));
      assert(!additionalItemParagraph.includes('<w:b/>'), 'Item da se\u00e7\u00e3o adicional de pre\u00e7o ficou em negrito');
      const additionalTotalIndex = xml.indexOf('TOTAL LOCA\u00c7\u00c3O:', additionalPriceIndex);
      const additionalTotalParagraphStart = Math.max(
        xml.lastIndexOf('<w:p>', additionalTotalIndex),
        xml.lastIndexOf('<w:p ', additionalTotalIndex)
      );
      const additionalTotalParagraph = xml.slice(additionalTotalParagraphStart, xml.indexOf('</w:p>', additionalTotalIndex));
      assert(additionalTotalParagraph.includes('<w:b/>'), 'Total da se\u00e7\u00e3o adicional de pre\u00e7o perdeu o negrito');
      const customTableIndex = xml.indexOf('Bomba A');
      const customTableIndentIndex = xml.lastIndexOf('<w:tblInd w:w="390" w:type="dxa"/>', customTableIndex);
      const customTableStart = xml.lastIndexOf('<w:tbl>', customTableIndentIndex);
      const customTableEnd = xml.indexOf('8.  OBSERVA\u00c7\u00d5ES', customTableIndex);
      const customTable = xml.slice(customTableStart, customTableEnd);
      assert(customTable.includes('<w:tblW w:w="10245" w:type="dxa"/>'), 'Tabela adicional n\u00e3o ocupa a largura padr\u00e3o');
      assert(customTable.includes('<w:tblInd w:w="390" w:type="dxa"/>'), 'Tabela adicional desalinhada das demais tabelas');
      assert(customTable.includes('<v:roundrect'), 'Cabe\u00e7alho arredondado da tabela adicional ausente');
      assert(customTable.includes('arcsize="7864f"'), 'Arredondamento da tabela adicional incorreto');
      assert(customTable.includes('fillcolor="#bfbfbf"'), 'Cor do cabe\u00e7alho da tabela adicional incorreta');
      assert(customTable.includes('strokecolor="#404040"'), 'Contorno do cabe\u00e7alho da tabela adicional incorreto');
      assert(customTable.includes('width:512.25pt;height:15.1pt'), 'Cabe\u00e7alho da tabela adicional fora do tamanho padr\u00e3o');
      assert((customTable.match(/<w:gridCol w:w="3415"\/>/g) || []).length >= 3, 'Colunas da tabela adicional n\u00e3o est\u00e3o sim\u00e9tricas');
      assert(customTable.includes('w:color="B7B7B7"'), 'Linhas da tabela adicional com cor incorreta');
      assert(customTable.includes('w:ascii="Calibri"'), 'Fonte Calibri ausente na tabela adicional');
      assert(customTable.includes('<w:sz w:val="16"/>'), 'Tamanho 8 pt ausente na tabela adicional');
      assert(!customTable.includes('w:color="808080"'), 'Tabela adicional ainda usa a borda externa antiga');
      const topicIndex = xml.indexOf('4.  CRONOGRAMA');
      const topicSection = xml.slice(xml.lastIndexOf('<w:p>', topicIndex), xml.indexOf('4.1  Etapas do servi\u00e7o'));
      assert(topicSection.includes('w:asciiTheme="minorHAnsi"'), 'Fonte padr\u00e3o do t\u00f3pico ausente');
      assert(topicSection.includes('<w:sz w:val="16"/>'), 'Tamanho padr\u00e3o do t\u00f3pico ausente');
      assert(topicSection.includes('strokecolor="#002060"'), 'Cor da linha do t\u00f3pico incorreta');
      assert(topicSection.includes('from="19.55pt,1.55pt" to="528.05pt,2.3pt"'), 'Comprimento ou posi\u00e7\u00e3o da linha do t\u00f3pico incorreto');
    }

    if (test.name === 'reordered') {
      ['1.  DADOS COMERCIAIS', '2.  OBJETO', '3.  CRONOGRAMA', '3.1  Etapas do servi\u00e7o', '3.1.1  Detalhamento da etapa', '4.  ESCOPO DE FORNECIMENTO', '4.1  Descri\u00e7\u00e3o dos Servi\u00e7os', '5.  DOCUMENTA\u00c7\u00c3O', '6.  PRE\u00c7O', '7.  EQUIPAMENTOS', '8.  OBSERVA\u00c7\u00d5ES'].forEach((value) => {
        assert(xml.includes(value), `Se\u00e7\u00e3o reordenada ou renumerada ausente: ${value}`);
      });
      const orderedMarkers = ['1.  DADOS COMERCIAIS', '2.  OBJETO', '3.  CRONOGRAMA', '4.  ESCOPO DE FORNECIMENTO', '5.  DOCUMENTA\u00c7\u00c3O', '6.  PRE\u00c7O', '7.  EQUIPAMENTOS', '8.  OBSERVA\u00c7\u00d5ES'];
      const markerPositions = orderedMarkers.map((marker) => xml.indexOf(marker));
      assert(markerPositions.every((position, index) => index === 0 || position > markerPositions[index - 1]), 'Ordem escolhida n\u00e3o foi aplicada ao DOCX');
    }

    console.log(`${test.name}: DOCX v\u00e1lido`);
  } finally {
    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
  }
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}
