const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const PizZip = require('pizzip');
const { DOMParser } = require('@xmldom/xmldom');
const { gerarDocx, prepareData } = require('../src/lib/gerarDocx');

const base = JSON.parse(fs.readFileSync(path.join(__dirname, 'dados-teste.json'), 'utf8'));
const currencyData = prepareData({
  topicos_preco: [{
    titulo: 'SERVI\u00c7OS',
    tipo: 'servico',
    itens: [
      { quant: 1, valor_unit: 372, valor_total: 372 },
      { quant: 1, valor_unit: -6708.5, valor_total: -6708.5 }
    ]
  }],
  preco_total_numero: -6336.5
});
assert(currencyData.topicos_preco[0].itens[0].valor_unit === 'R$ 372,00', 'Valor unit\u00e1rio positivo fora do padr\u00e3o monet\u00e1rio');
assert(currencyData.topicos_preco[0].itens[1].valor_unit === '-R$ 6.708,50', 'Valor unit\u00e1rio negativo fora do padr\u00e3o monet\u00e1rio');
assert(currencyData.topicos_preco[0].total === '-R$ 6.336,50', 'Total do t\u00f3pico fora do padr\u00e3o monet\u00e1rio');
assert(currencyData.preco_total_numero === '-R$ 6.336,50', 'Pre\u00e7o total fora do padr\u00e3o monet\u00e1rio');

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
    name: 'object-observations',
    data: {
      ...base,
      objeto_observacoes: [
        base.objeto,
        'Inspe\u00e7\u00e3o complementar do equipamento'
      ]
    }
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
          titulo: 'INVESTIMENTO',
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
  name: 'custom-table-widths',
  data: {
    ...base,
    blocos_adicionais: [
      {
        id: 'tabela-larguras',
        tipo: 'tabela',
        titulo: 'DIMENSIONAMENTO',
        colunas: [
          { id: 'descricao', nome: 'Descri\u00e7\u00e3o', largura: 60 },
          { id: 'codigo', nome: 'C\u00f3digo de identifica\u00e7\u00e3o completo do equipamento', largura: 30 },
          { id: 'quantidade', nome: 'Quantidade', largura: 20 }
        ],
        linhas: [
          { valores: { descricao: 'Bomba dimensionada', codigo: 'XP-20', quantidade: '3' } }
        ]
      }
    ]
  }
});

tests.push({
  name: 'custom-table-single-line-long',
  data: {
    ...base,
    blocos_adicionais: [
      {
        id: 'tabela-linha-unica',
        tipo: 'tabela',
        titulo: 'DADOS DO EQUIPAMENTO',
        colunas: [
          { id: 'item', nome: 'Item', largura: 10 },
          { id: 'descricao', nome: 'Descri\u00e7\u00e3o t\u00e9cnica detalhada do equipamento principal', largura: 60 },
          { id: 'fabricante', nome: 'Fabricante', largura: 15 },
          { id: 'modelo', nome: 'Modelo', largura: 15 }
        ],
        linhas: [
          { valores: { item: '1', descricao: 'Compressor principal', fabricante: 'Supply', modelo: 'SM-10' } }
        ]
      }
    ]
  }
});

tests.push({
  name: 'subtopic-hierarchy',
  data: {
    ...base,
    blocos_adicionais: [
      {
        id: 'topico-hierarquia',
        tipo: 'texto',
        titulo: 'HIERARQUIA',
        subtopicos: [
          {
            id: 'nivel-1-a',
            titulo: 'Primeiro n\u00edvel A',
            subtopicos: [
              {
                id: 'nivel-2-a',
                titulo: 'Segundo n\u00edvel A',
                subtopicos: [
                  { id: 'nivel-3-a', titulo: 'Terceiro n\u00edvel A' }
                ]
              },
              { id: 'nivel-2-b', titulo: 'Segundo n\u00edvel B' }
            ]
          },
          { id: 'nivel-1-b', titulo: 'Primeiro n\u00edvel B' }
        ]
      }
    ]
  }
});

tests.push({
  name: 'additional-topics-standard',
  data: {
    ...base,
    blocos_adicionais: [
      {
        id: 'topico-prazos',
        tipo: 'texto',
        titulo: 'PRAZOS',
        subtopicos: [
          {
            id: 'prazo-material',
            titulo: 'Prazo de material',
            observacoes: ['At\u00e9 10 dias ap\u00f3s o recebimento da PO.']
          }
        ]
      },
      {
        id: 'topico-documentacao-tecnica',
        tipo: 'texto',
        titulo: 'DOCUMENTA\u00c7\u00c3O T\u00c9CNICA',
        subtopicos: [
          {
            id: 'folha-dados',
            titulo: 'Folha de dados',
            observacoes: ['Anexo ao final da proposta t\u00e9cnica.']
          }
        ]
      }
    ]
  }
});

tests.push({
  name: 'price-without-title',
  data: {
    ...base,
    blocos_adicionais: [
      {
        id: 'preco-sem-titulo',
        tipo: 'preco',
        titulo: '',
        topicos_preco: [
          {
            titulo: 'ITENS',
            tipo: 'personalizado',
            itens: [
              { item: '0001', descricao: 'Item sem t\u00edtulo de se\u00e7\u00e3o', quant: 1, un: 'UN', valor_unit: 100, valor_total: 100 }
            ]
          }
        ],
        preco_total_numero: 100
      }
    ]
  }
});

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
    const numberingXml = zip.file('word/numbering.xml').asText();
    const xmlErrors = [];
    new DOMParser({
      onError: (level, message) => {
        if (level !== 'warning') xmlErrors.push(message);
      }
    }).parseFromString(xml, 'application/xml');
    new DOMParser({
      onError: (level, message) => {
        if (level !== 'warning') xmlErrors.push(`numbering.xml: ${message}`);
      }
    }).parseFromString(numberingXml, 'application/xml');
    assert(xmlErrors.length === 0, `${test.name}: XML inv\u00e1lido (${xmlErrors.join('; ')})`);
    assert(xml.includes(base.empresa_cliente), `${test.name}: dados comerciais ausentes`);
    assert(!xml.includes('INFORMA\u00c7\u00d5ES ADICIONAIS'), `${test.name}: se\u00e7\u00e3o fixa de informa\u00e7\u00f5es adicionais ainda presente`);
    assert(!xml.includes('Prazo de Execu\u00e7\u00e3o'), `${test.name}: conte\u00fado fixo de informa\u00e7\u00f5es adicionais ainda presente`);
    if (!test.data.secoes_excluidas?.includes('objeto')) {
      const objectValueIndex = xml.indexOf(String(base.objeto).split('&')[0].trim());
      const objectParagraphStart = Math.max(
        xml.lastIndexOf('<w:p>', objectValueIndex),
        xml.lastIndexOf('<w:p ', objectValueIndex)
      );
      const objectParagraph = xml.slice(objectParagraphStart, xml.indexOf('</w:p>', objectValueIndex));
      assert(objectParagraph.includes('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>'), `${test.name}: marcador do Objeto ausente`);
      assert(objectParagraph.includes('<w:ind w:left="1191" w:right="799" w:hanging="227"/>'), `${test.name}: observa\u00e7\u00e3o do Objeto fora das margens`);
    }
    if (test.name === 'legacy') {
      const numberingInstance = Array.from(numberingXml.matchAll(/<w:num\b[\s\S]*?<\/w:num>/g))
        .map((match) => match[0])
        .find((value) => /\sw:numId="1"/.test(value));
      const abstractNumberId = numberingInstance?.match(/<w:abstractNumId w:val="(\d+)"\/>/)?.[1];
      const abstractDefinition = numberingXml.match(new RegExp(
        `<w:abstractNum\\b[^>]*\\sw:abstractNumId="${abstractNumberId}"[^>]*>[\\s\\S]*?<\\/w:abstractNum>`
      ))?.[0] || '';
      const bulletLevel = abstractDefinition.match(
        /<w:lvl\b[^>]*\sw:ilvl="1"[^>]*>[\s\S]*?<\/w:lvl>/
      )?.[0] || '';
      assert(bulletLevel.includes('<w:lvlText w:val="\u2022"/>'), 'Marcador das observa\u00e7\u00f5es n\u00e3o \u00e9 uma bolinha preta padronizada');
      assert(
        bulletLevel.includes('<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>')
          && bulletLevel.includes('<w:sz w:val="16"/><w:szCs w:val="16"/>'),
        'Fonte ou tamanho do marcador est\u00e1 desalinhado do texto'
      );
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
    if (test.name === 'object-observations') {
      const objectStart = xml.indexOf('OBJETO');
      const scopeStart = xml.indexOf('ESCOPO DE FORNECIMENTO', objectStart);
      const objectSection = xml.slice(objectStart, scopeStart);
      assert(objectSection.includes('Inspe\u00e7\u00e3o complementar do equipamento'), 'Segunda observa\u00e7\u00e3o do Objeto ausente');
      assert((objectSection.match(/<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="1"\/><\/w:numPr>/g) || []).length === 2, 'Quantidade de marcadores do Objeto incorreta');
    }
    if (test.name === 'price-without-title') {
      assert(xml.includes('Item sem t\u00edtulo de se\u00e7\u00e3o'), 'Itens do bloco de pre\u00e7o sem t\u00edtulo foram removidos');
      assert(!xml.includes('4.  PRE\u00c7O'), 'T\u00edtulo vazio do bloco de pre\u00e7o voltou ao DOCX');
      assert(!xml.includes('ITENS / PRE\u00c7O'), 'Nome padr\u00e3o apareceu mesmo ap\u00f3s ser apagado');
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
      ['4.  CRONOGRAMA', '4.1  ', 'Etapas do servi\u00e7o', '4.1.1  ', 'Detalhamento da etapa', 'Inspe\u00e7\u00e3o inicial', 'Mobiliza\u00e7\u00e3o', '5.  DOCUMENTA\u00c7\u00c3O', 'DESCRI\u00c7\u00c3O', 'N\u00ba DO DOCUMENTO', 'Solicita\u00e7\u00e3o de cota\u00e7\u00e3o', 'SC35085 | RC73960 | Cota\u00e7\u00e3o 167383', '14/07/2026', '6.  INVESTIMENTO', 'LOCA\u00c7\u00c3O', 'Equipamento especial', '1.250,00', 'R$ 2.500,00 (Dois mil e quinhentos reais)', 'Validade da Proposta:', '30/08/2026', '7.  EQUIPAMENTOS', 'Bomba A', '8.  OBSERVA\u00c7\u00d5ES', 'Conte\u00fado ap\u00f3s quebra.'].forEach((value) => {
        assert(xml.includes(value), `Conte\u00fado flex\u00edvel ausente: ${value}`);
      });
      assert(xml.indexOf('Bomba A') < xml.indexOf('Atenciosamente'), 'Blocos inseridos depois da assinatura');
      const signatureGap = xml.slice(xml.indexOf('Conte\u00fado ap\u00f3s quebra.'), xml.indexOf('Atenciosamente'));
      assert(signatureGap.includes('<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'), 'Linha vazia antes da assinatura ausente');
      assert(xml.includes('w:type="page"'), 'Quebra de p\u00e1gina ausente');
      const documentationIndex = xml.indexOf('Solicita\u00e7\u00e3o de cota\u00e7\u00e3o');
      const documentationIndentIndex = xml.lastIndexOf('<w:tblInd w:w="390" w:type="dxa"/>', documentationIndex);
      const documentationTableStart = xml.lastIndexOf('<w:tbl>', documentationIndentIndex);
      const documentationTableEnd = xml.indexOf('6.  INVESTIMENTO', documentationIndex);
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
      const accountingValueIndex = xml.indexOf('1.250,00', additionalPriceIndex);
      const accountingParagraphStart = Math.max(
        xml.lastIndexOf('<w:p>', accountingValueIndex),
        xml.lastIndexOf('<w:p ', accountingValueIndex)
      );
      const accountingParagraph = xml.slice(accountingParagraphStart, xml.indexOf('</w:p>', accountingValueIndex));
      assert(accountingParagraph.includes('<w:t>R$</w:t>'), 'S\u00edmbolo monet\u00e1rio n\u00e3o ficou em uma posi\u00e7\u00e3o cont\u00e1bil separada');
      assert(accountingParagraph.includes('<w:tab/>'), 'Tabula\u00e7\u00e3o cont\u00e1bil ausente no valor unit\u00e1rio');
      assert(accountingParagraph.includes('<w:tab w:val="right" w:pos="1175"/>'), 'Casas decimais do valor unit\u00e1rio sem alinhamento fixo');
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
      assert(customTable.includes('<w:vAlign w:val="center"/>'), 'Cabe\u00e7alho da tabela adicional com uma linha n\u00e3o voltou ao alinhamento central');
      assert(customTable.includes('<w:noWrap/>'), 'Cabe\u00e7alho da tabela adicional ainda permite quebra de linha');
      assert(customTable.includes('<w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/>'), 'Espa\u00e7amento vertical da tabela adicional difere dos outros cabe\u00e7alhos');
      assert(!customTable.includes('<w:trPr><w:trHeight w:hRule="exact" w:val="340"/><w:cantSplit/></w:trPr>'), 'Tabela adicional ainda possui uma altura interna que desloca o texto');
      assert((customTable.match(/<w:gridCol w:w="3415"\/>/g) || []).length >= 3, 'Colunas da tabela adicional n\u00e3o est\u00e3o sim\u00e9tricas');
      assert(customTable.includes('w:color="B7B7B7"'), 'Linhas da tabela adicional com cor incorreta');
      assert(customTable.includes('w:ascii="Calibri"'), 'Fonte Calibri ausente na tabela adicional');
      assert(customTable.includes('<w:sz w:val="16"/>'), 'Tamanho 8 pt ausente na tabela adicional');
      assert(!customTable.includes('w:color="808080"'), 'Tabela adicional ainda usa a borda externa antiga');
      const topicIndex = xml.indexOf('4.  CRONOGRAMA');
      const topicSection = xml.slice(xml.lastIndexOf('<w:p>', topicIndex), xml.indexOf('Etapas do servi\u00e7o'));
      assert(topicSection.includes('w:asciiTheme="minorHAnsi"'), 'Fonte padr\u00e3o do t\u00f3pico ausente');
      assert(topicSection.includes('<w:sz w:val="16"/>'), 'Tamanho padr\u00e3o do t\u00f3pico ausente');
      assert(topicSection.includes('strokecolor="#002060"'), 'Cor da linha do t\u00f3pico incorreta');
      assert(topicSection.includes('from="19.55pt,1.55pt" to="528.05pt,2.3pt"'), 'Comprimento ou posi\u00e7\u00e3o da linha do t\u00f3pico incorreto');
      const subtopicIndex = xml.indexOf('Etapas do servi\u00e7o');
      const subtopicParagraphStart = Math.max(
        xml.lastIndexOf('<w:p>', subtopicIndex),
        xml.lastIndexOf('<w:p ', subtopicIndex)
      );
      const subtopicParagraph = xml.slice(subtopicParagraphStart, xml.indexOf('</w:p>', subtopicIndex));
      assert(subtopicParagraph.includes('<w:ind w:left="964" w:right="799" w:hanging="284"/>'), 'Subt\u00f3pico adicional n\u00e3o segue o alinhamento do Escopo');
      assert(subtopicParagraph.includes('w:ascii="Calibri"'), 'Fonte do subt\u00f3pico adicional difere do padr\u00e3o do Escopo');
      assert(!subtopicParagraph.includes('<w:spacing'), 'Espa\u00e7amento do subt\u00f3pico adicional difere do padr\u00e3o do Escopo');
      const subtopicNumberRunEnd = subtopicParagraph.indexOf('</w:r>');
      const subtopicNumberRun = subtopicParagraph.slice(0, subtopicNumberRunEnd);
      const subtopicTitleRun = subtopicParagraph.slice(subtopicNumberRunEnd);
      assert(subtopicNumberRun.includes('<w:b/>'), 'Autonumera\u00e7\u00e3o do subt\u00f3pico adicional n\u00e3o est\u00e1 em negrito');
      assert(!subtopicTitleRun.includes('<w:b/>'), 'Texto do subt\u00f3pico adicional ficou em negrito');
      assert((subtopicParagraph.match(/Etapas do servi\u00e7o/g) || []).length === 1, 'Texto do subt\u00f3pico adicional foi duplicado');
      const nestedSubtopicIndex = xml.indexOf('Detalhamento da etapa');
      const nestedSubtopicStart = Math.max(
        xml.lastIndexOf('<w:p>', nestedSubtopicIndex),
        xml.lastIndexOf('<w:p ', nestedSubtopicIndex)
      );
      const nestedSubtopicParagraph = xml.slice(nestedSubtopicStart, xml.indexOf('</w:p>', nestedSubtopicIndex));
      assert(nestedSubtopicParagraph.includes('<w:ind w:left="1244" w:right="799" w:hanging="284"/>'), 'Subt\u00f3pico aninhado n\u00e3o seguiu o recuo hier\u00e1rquico');
      const nestedNumberRunEnd = nestedSubtopicParagraph.indexOf('</w:r>');
      assert(nestedSubtopicParagraph.slice(0, nestedNumberRunEnd).includes('<w:b/>'), 'Autonumera\u00e7\u00e3o do subt\u00f3pico aninhado n\u00e3o est\u00e1 em negrito');
      assert(!nestedSubtopicParagraph.slice(nestedNumberRunEnd).includes('<w:b/>'), 'Texto do subt\u00f3pico aninhado ficou em negrito');
      const observationIndex = xml.indexOf('Execu\u00e7\u00e3o em duas etapas.');
      const observationParagraphStart = Math.max(
        xml.lastIndexOf('<w:p>', observationIndex),
        xml.lastIndexOf('<w:p ', observationIndex)
      );
      const observationParagraph = xml.slice(observationParagraphStart, xml.indexOf('</w:p>', observationIndex));
      assert(observationParagraph.includes('<w:ind w:left="1191" w:right="799" w:hanging="227"/>'), 'Observa\u00e7\u00e3o adicional n\u00e3o segue o alinhamento do Escopo');
      assert(observationParagraph.includes('w:ascii="Calibri"'), 'Fonte da observa\u00e7\u00e3o adicional difere do padr\u00e3o do Escopo');
      assert(!observationParagraph.includes('<w:spacing'), 'Espa\u00e7amento da observa\u00e7\u00e3o adicional difere do padr\u00e3o do Escopo');
      assert(observationParagraph.includes('w:right="799"'), 'Texto do t\u00f3pico adicional ultrapassa a margem direita');
      assert(!observationParagraph.includes('<w:keepNext/>'), 'Pagina\u00e7\u00e3o normal do t\u00f3pico adicional foi alterada');
    }

    if (test.name === 'reordered') {
      ['1.  DADOS COMERCIAIS', '2.  OBJETO', '3.  CRONOGRAMA', '3.1  ', 'Etapas do servi\u00e7o', '3.1.1  ', 'Detalhamento da etapa', '4.  ESCOPO DE FORNECIMENTO', '4.1  Descri\u00e7\u00e3o dos Servi\u00e7os', '5.  DOCUMENTA\u00c7\u00c3O', '6.  INVESTIMENTO', '7.  EQUIPAMENTOS', '8.  OBSERVA\u00c7\u00d5ES'].forEach((value) => {
        assert(xml.includes(value), `Se\u00e7\u00e3o reordenada ou renumerada ausente: ${value}`);
      });
      const orderedMarkers = ['1.  DADOS COMERCIAIS', '2.  OBJETO', '3.  CRONOGRAMA', '4.  ESCOPO DE FORNECIMENTO', '5.  DOCUMENTA\u00c7\u00c3O', '6.  INVESTIMENTO', '7.  EQUIPAMENTOS', '8.  OBSERVA\u00c7\u00d5ES'];
      const markerPositions = orderedMarkers.map((marker) => xml.indexOf(marker));
      assert(markerPositions.every((position, index) => index === 0 || position > markerPositions[index - 1]), 'Ordem escolhida n\u00e3o foi aplicada ao DOCX');
    }

    if (test.name === 'custom-table-widths') {
      const customTableIndex = xml.indexOf('Bomba dimensionada');
      const customTableIndentIndex = xml.lastIndexOf('<w:tblInd w:w="390" w:type="dxa"/>', customTableIndex);
      const customTableStart = xml.lastIndexOf('<w:tbl>', customTableIndentIndex);
      const customTableEnd = xml.indexOf('</w:tbl>', customTableIndex);
      const customTable = xml.slice(customTableStart, customTableEnd);
      assert(customTable.includes('<w:gridCol w:w="5588"/>'), 'Largura expandida da primeira coluna n\u00e3o foi aplicada no DOCX');
      assert(customTable.includes('<w:gridCol w:w="2794"/>'), 'Propor\u00e7\u00e3o da segunda coluna n\u00e3o foi preservada no DOCX');
      assert(customTable.includes('<w:gridCol w:w="1863"/>'), 'Propor\u00e7\u00e3o da terceira coluna n\u00e3o foi preservada no DOCX');
      assert(customTable.includes('height:15.1pt'), 'Cabe\u00e7alho da tabela adicional aumentou com t\u00edtulo longo');
      assert(customTable.includes('<w:trHeight w:hRule="exact" w:val="340"/>'), 'Cabe\u00e7alho da tabela adicional perdeu a altura padr\u00e3o');
      assert(customTable.includes('<w:vAlign w:val="center"/>'), 'Cabe\u00e7alho com t\u00edtulo longo n\u00e3o ficou centralizado');
      assert(customTable.includes('<w:noWrap/>'), 'T\u00edtulo longo ainda pode aumentar a altura do cabe\u00e7alho');
    }

    if (test.name === 'custom-table-single-line-long') {
      const customTableIndex = xml.indexOf('Compressor principal');
      const customTableIndentIndex = xml.lastIndexOf('<w:tblInd w:w="390" w:type="dxa"/>', customTableIndex);
      const customTableStart = xml.lastIndexOf('<w:tbl>', customTableIndentIndex);
      const customTableEnd = xml.indexOf('</w:tbl>', customTableIndex);
      const customTable = xml.slice(customTableStart, customTableEnd);
      assert(customTable.includes('height:15.1pt'), 'Cabe\u00e7alho sem quebra n\u00e3o voltou \u00e0 altura fina');
      assert(customTable.includes('<w:trHeight w:hRule="exact" w:val="340"/>'), 'Linha \u00fanica do cabe\u00e7alho n\u00e3o manteve a altura padr\u00e3o');
      assert(customTable.includes('<w:vAlign w:val="center"/>'), 'Linha \u00fanica do cabe\u00e7alho n\u00e3o ficou centralizada');
    }

    if (test.name === 'subtopic-hierarchy') {
      [
        ['Primeiro n\u00edvel A', 964],
        ['Primeiro n\u00edvel B', 964],
        ['Segundo n\u00edvel A', 1244],
        ['Segundo n\u00edvel B', 1244],
        ['Terceiro n\u00edvel A', 1244]
      ].forEach(([title, leftIndent]) => {
        const titleIndex = xml.indexOf(title);
        const paragraphStart = Math.max(
          xml.lastIndexOf('<w:p>', titleIndex),
          xml.lastIndexOf('<w:p ', titleIndex)
        );
        const paragraph = xml.slice(paragraphStart, xml.indexOf('</w:p>', titleIndex));
        const numberRunEnd = paragraph.indexOf('</w:r>');
        assert(paragraph.includes(`<w:ind w:left="${leftIndent}" w:right="799" w:hanging="284"/>`), `Recuo hier\u00e1rquico incorreto em ${title}`);
        assert(paragraph.slice(0, numberRunEnd).includes('<w:b/>'), `Autonumera\u00e7\u00e3o sem negrito em ${title}`);
        assert(!paragraph.slice(numberRunEnd).includes('<w:b/>'), `Texto em negrito indevido em ${title}`);
      });
      ['4.1.1  ', '4.1.2  ', '4.1.3  '].forEach((number) => {
        assert(xml.includes(number), `Numera\u00e7\u00e3o horizontal ausente: ${number.trim()}`);
      });
      assert(!xml.includes('4.1.1.1  '), 'Hierarquia documental aprofundou al\u00e9m do padr\u00e3o de dois n\u00edveis');
    }

    if (test.name === 'additional-topics-standard') {
      ['Prazo de material', 'Folha de dados'].forEach((title) => {
        const titleIndex = xml.indexOf(title);
        const paragraphStart = Math.max(
          xml.lastIndexOf('<w:p>', titleIndex),
          xml.lastIndexOf('<w:p ', titleIndex)
        );
        const paragraph = xml.slice(paragraphStart, xml.indexOf('</w:p>', titleIndex));
        const numberRunEnd = paragraph.indexOf('</w:r>');
        assert(paragraph.includes('<w:ind w:left="964" w:right="799" w:hanging="284"/>'), `${title}: alinhamento diferente do Escopo`);
        assert(paragraph.includes('w:ascii="Calibri"'), `${title}: fonte diferente do Escopo`);
        assert(paragraph.slice(0, numberRunEnd).includes('<w:b/>'), `${title}: autonumera\u00e7\u00e3o sem negrito`);
        assert(!paragraph.slice(numberRunEnd).includes('<w:b/>'), `${title}: texto em negrito indevido`);
      });
      ['At\u00e9 10 dias ap\u00f3s o recebimento da PO.', 'Anexo ao final da proposta t\u00e9cnica.'].forEach((observation) => {
        const observationIndex = xml.indexOf(observation);
        const paragraphStart = Math.max(
          xml.lastIndexOf('<w:p>', observationIndex),
          xml.lastIndexOf('<w:p ', observationIndex)
        );
        const paragraph = xml.slice(paragraphStart, xml.indexOf('</w:p>', observationIndex));
        assert(paragraph.includes('<w:ind w:left="1191" w:right="799" w:hanging="227"/>'), `${observation}: marcador fora do padr\u00e3o do Escopo`);
        assert(paragraph.includes('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>'), `${observation}: marcador preto ausente`);
        assert(paragraph.includes('w:ascii="Calibri"'), `${observation}: fonte diferente do Escopo`);
      });
    }

    console.log(`${test.name}: DOCX v\u00e1lido`);
  } finally {
    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
  }
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}
