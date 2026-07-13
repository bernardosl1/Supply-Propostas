const fs = require('node:fs');
const path = require('node:path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { getPropostasDir } = require('./paths');

const templatePath = path.resolve(__dirname, '..', 'templates', 'modelo_proposta.docx');
const FIXED_ADDITIONAL_INFO = [
  ['51', 'Prazo de Execução'],
  ['511', 'Serviços: Estimados em 03 (três) dias, incluindo translado ida & volta Equipe;'],
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
  const technicalTeamSource = Array.isArray(data.equipe_tecnica_itens) && data.equipe_tecnica_itens.length
    ? data.equipe_tecnica_itens
    : String(data.equipe_tecnica || '').split('|');
  const technicalTeam = technicalTeamSource
    .map((item) => (typeof item === 'string' ? item : item?.item))
    .map((item) => String(item || '').trim())
    .filter(Boolean);
  const savedFixedInfo = data.informacoes_adicionais_fixas;
  const hasSavedFixedInfo = savedFixedInfo && typeof savedFixedInfo === 'object' && !Array.isArray(savedFixedInfo);
  const fixedInfoData = {};
  FIXED_ADDITIONAL_INFO.forEach(([id, defaultText]) => {
    let text = hasSavedFixedInfo ? savedFixedInfo[id] : defaultText;
    if (!hasSavedFixedInfo && id === '511' && data.prazo_execucao_dias) {
      text = defaultText.replace('03 (três) dias', String(data.prazo_execucao_dias));
    }
    text = String(text || '').trim();
    fixedInfoData[`mostrar_info_${id}`] = Boolean(text);
    fixedInfoData[`texto_info_${id}`] = text;
  });

  return {
    ...data,
    ...fixedInfoData,
    servicos_descricao: (data.servicos_descricao || []).map((item) => (
      typeof item === 'string' ? { item } : item
    )),
    equipe_tecnica: technicalTeam.join(' | '),
    equipe_tecnica_itens: technicalTeam.map((item) => ({ item })),
    informacoes_adicionais: (data.informacoes_adicionais || []).map((item) => (
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
