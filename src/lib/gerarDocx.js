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
  }).format(number);
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
      valor_unit: formatNumber(valorUnit),
      valor_total: formatNumber(total)
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
      total: formatNumber(totalNumber)
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
  const totalServicosNumber = sumItems(itensServico);
  const totalConsumiveisNumber = sumItems(itensConsumiveis);
  const totalTopicosNumber = topicosPreco.reduce((sum, topic) => sum + sumItems(topic.itens), 0);
  const totalNumber = data.preco_total_numero || totalTopicosNumber;

  return {
    ...data,
    servicos_descricao: (data.servicos_descricao || []).map((item) => (
      typeof item === 'string' ? { item } : item
    )),
    topicos_preco: topicosPreco,
    itens_servico: itensServico,
    itens_consumiveis: itensConsumiveis,
    total_servicos: formatNumber(totalServicosNumber),
    total_consumiveis: formatNumber(totalConsumiveisNumber),
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
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: () => ''
  });

  doc.render(prepareData(data));

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

module.exports = {
  gerarDocx,
  prepareData
};

if (require.main === module) {
  const dataPath = process.argv[2] || path.resolve(__dirname, '..', '..', 'scripts', 'dados-teste.json');
  const outputPath = process.argv[3];
  const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
  const generated = gerarDocx(data, outputPath);
  console.log(`Documento gerado em ${generated}`);
}
