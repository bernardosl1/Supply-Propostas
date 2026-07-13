const state = {
  proposals: [],
  filteredProposals: [],
  lastOutputPath: '',
  editingId: ''
};

const moneyFormatter = new Intl.NumberFormat('pt-BR', {
  style: 'currency',
  currency: 'BRL'
});

const numberFormatter = new Intl.NumberFormat('pt-BR', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

const versionTarget = document.querySelector('.sidebar-note span:last-child');
const form = document.querySelector('#proposal-form');
const servicesList = document.querySelector('#services-list');
const technicalTeamList = document.querySelector('#technical-team-list');
const fixedAdditionalInfoList = document.querySelector('#fixed-additional-info-list');
const additionalInfoList = document.querySelector('#additional-info-list');
const priceTopics = document.querySelector('#price-topics');
const resultPanel = document.querySelector('#result-panel');
const resultPath = document.querySelector('#result-path');
const pdfStatus = document.querySelector('#pdf-status');
const formAlert = document.querySelector('#form-alert');
const openDocxButton = document.querySelector('#open-docx');
const exportPdfButton = document.querySelector('#export-pdf');
const grandTotal = document.querySelector('#grand-total');
const proposalList = document.querySelector('#proposal-list');
const emptyState = document.querySelector('#empty-state');
const summaryCount = document.querySelector('#summary-count');
const summaryLast = document.querySelector('#summary-last');
const historySearch = document.querySelector('#history-search');
const selectDocxButton = document.querySelector('#select-docx');
const docxStatus = document.querySelector('#docx-status');

const DEFAULT_PRICE_TOPICS = [
  { tipo: 'servico', titulo: 'Itens de servi\u00e7o', ncm: '-----', un: 'HH' },
  { tipo: 'consumivel', titulo: 'Consum\u00edveis', ncm: '', un: 'PC' }
];

const DEFAULT_FIXED_ADDITIONAL_INFO = [
  { id: '51', number: '5.1', text: 'Prazo de Execução' },
  { id: '511', number: '5.1.1', text: 'Serviços: Estimados em 03 (três) dias, incluindo translado ida & volta Equipe;' },
  { id: '512', number: '5.1.2', text: 'Consumíveis/Materiais/Equipamentos: Imediato;' },
  { id: '513', number: '5.1.3', text: 'O prazo de Execução dos serviços pode ser alterado de acordo com as condições de execução dos mesmos;' },
  { id: '514', number: '5.1.4', text: 'Caso o prazo de execução seja diferente do mencionado, o valor total da presente proposta será alterado conforme as informações abaixo:' },
  { id: '514a', number: '•', text: 'De Segunda-feira a Sexta-feira das 17h30min até as 7h30min 50% adicional;' },
  { id: '514b', number: '•', text: 'Durante o Sábado 50% adicional o dia todo;' },
  { id: '514c', number: '•', text: 'Domingos e Feriados 100% adicional o dia todo;' },
  { id: '514d', number: '•', text: 'Diária Offshore quando o barco em operação ou fundeado e a equipe permanecer a bordo.' },
  { id: '52', number: '5.2', text: 'Caso haja a necessidade de substituição de algum componente não previsto nesta proposta o mesmo será objeto de orçamento aditivo.' },
  { id: '53', number: '5.3', text: 'Após o término do serviço será enviado uma medição, incluindo as horas de viagem, espera a bordo e a disposição, caso necessário.' },
  { id: '54', number: '5.4', text: 'Todas as despesas com deslocamento, alimentação e estadia da equipe, caso necessário, serão por conta do cliente;' },
  { id: '55', number: '5.5', text: 'Os Equipamentos e Ferramentas de propriedade da SUPPLY MARINE deverão ser devolvidos no prazo máximo de 03 (três) dias após a conclusão dos serviços. Caso contrário, a SUPPLY MARINE cobrará pelos custos de cessão dos mesmos conforme tabela abaixo:' }
];

init();

function init() {
  bindNavigation();
  bindProposalImport();
  bindForm();
  bindHistory();
  resetFormToDefaults();
  loadHistory();

  if (window.supplyMarine && versionTarget) {
    window.supplyMarine.getVersion().then((version) => {
      versionTarget.textContent = `App local pronto. v${version}`;
    });
  }
}

function bindNavigation() {
  document.querySelectorAll('[data-view-button]').forEach((button) => {
    button.addEventListener('click', () => {
      if (button.dataset.viewButton === 'form') {
        resetFormToDefaults();
      }
      showView(button.dataset.viewButton);
    });
  });
  document.querySelectorAll('[data-scroll-target]').forEach((button) => {
    button.addEventListener('click', () => {
      showView('form');
      document.querySelector(`#${button.dataset.scrollTarget}`).scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });
}

function bindHistory() {
  historySearch.addEventListener('input', renderHistory);
}

function showView(viewName) {
  document.querySelectorAll('.view').forEach((view) => view.classList.remove('active'));
  document.querySelector(`#${viewName}-view`).classList.add('active');
  document.querySelectorAll('.nav-link').forEach((button) => {
    button.classList.toggle('active', button.dataset.viewButton === viewName);
  });
}

function bindForm() {
  document.querySelector('[data-add-service]').addEventListener('click', () => addServiceDescription(''));
  document.querySelector('[data-add-technical-team]').addEventListener('click', () => addTechnicalTeamMember(''));
  document.querySelector('[data-add-additional-info]').addEventListener('click', () => addAdditionalInfo(''));
  document.querySelector('[data-add-price-topic]').addEventListener('click', () => addPriceTopic());
  form.addEventListener('input', recalculate);
  form.addEventListener('submit', handleSubmit);
  openDocxButton.addEventListener('click', () => {
    if (state.lastOutputPath) {
      window.supplyMarine.abrirArquivo(state.lastOutputPath);
    }
  });
  exportPdfButton.addEventListener('click', () => exportPdf(state.lastOutputPath));
}

function bindProposalImport() {
  selectDocxButton.addEventListener('click', handleSelectDocx);
}

async function handleSelectDocx() {
  if (!window.supplyMarine) {
    docxStatus.textContent = 'API do Electron indisponível.';
    return;
  }

  const filePath = await window.supplyMarine.selecionarDocx();
  if (!filePath) {
    return;
  }

  selectDocxButton.disabled = true;
  docxStatus.textContent = 'Importando proposta Word...';

  try {
    const result = await window.supplyMarine.importarDocx(filePath);
    resetFormToDefaults();
    applyImportedData(result);
    state.editingId = '';
    state.lastOutputPath = '';
    resultPanel.hidden = true;
    document.querySelector('#form-title').textContent = 'Proposta importada';
    docxStatus.textContent = `Importado: ${result.data?.numero_documento || filePath}`;
  } catch (error) {
    docxStatus.textContent = error.message || String(error);
  } finally {
    selectDocxButton.disabled = false;
  }
}

function fillDefaults() {
  setValue('data_documento', todayPtBr());
  setValue('validade_proposta', addDaysPtBr(30));
  setValue('moeda', 'Real R$');
  setValue('impostos', 'inclusos no preço');
}

function resetFormToDefaults() {
  state.editingId = '';
  form.reset();
  servicesList.innerHTML = '';
  technicalTeamList.innerHTML = '';
  fixedAdditionalInfoList.innerHTML = '';
  additionalInfoList.innerHTML = '';
  priceTopics.innerHTML = '';
  resultPanel.hidden = true;
  state.lastOutputPath = '';
  document.querySelector('#form-title').textContent = 'Formulário da proposta';

  addServiceDescription('');
  addTechnicalTeamMember('');
  renderFixedAdditionalInfo();
  addAdditionalInfo('');
  DEFAULT_PRICE_TOPICS.forEach((topic) => addPriceTopic({ ...topic }, true));
  fillDefaults();
  recalculate();
}

function setValue(name, value) {
  const field = form.elements[name];
  if (field) {
    field.value = value;
  }
}

function addServiceDescription(value) {
  const row = document.createElement('div');
  row.className = 'stack-row';
  row.innerHTML = `
    <input data-service-description value="${escapeHtml(value)}" placeholder="Descrição do serviço">
    <button class="danger-action" type="button" title="Remover">x</button>
  `;
  row.querySelector('button').addEventListener('click', () => {
    row.remove();
    recalculate();
  });
  servicesList.appendChild(row);
}

function addTechnicalTeamMember(value) {
  const row = document.createElement('div');
  row.className = 'stack-row';
  row.innerHTML = `
    <input data-technical-team value="${escapeHtml(value)}" placeholder="Integrante / função da equipe técnica">
    <button class="danger-action" type="button" title="Remover">x</button>
  `;
  row.querySelector('button').addEventListener('click', () => {
    row.remove();
    recalculate();
  });
  technicalTeamList.appendChild(row);
}

function addAdditionalInfo(value) {
  const row = document.createElement('div');
  row.className = 'stack-row';
  row.innerHTML = `
    <input data-additional-info value="${escapeHtml(value)}" placeholder="Informação adicional">
    <button class="danger-action" type="button" title="Remover">x</button>
  `;
  row.querySelector('button').addEventListener('click', () => row.remove());
  additionalInfoList.appendChild(row);
}

function renderFixedAdditionalInfo(data = {}) {
  const saved = data.informacoes_adicionais_fixas;
  const hasSavedValues = saved && typeof saved === 'object' && !Array.isArray(saved);
  fixedAdditionalInfoList.innerHTML = '';

  DEFAULT_FIXED_ADDITIONAL_INFO.forEach((definition) => {
    if (hasSavedValues && !Object.prototype.hasOwnProperty.call(saved, definition.id)) {
      return;
    }
    let value = hasSavedValues ? saved[definition.id] : definition.text;
    if (!hasSavedValues && definition.id === '511' && data.prazo_execucao_dias) {
      value = definition.text.replace('03 (três) dias', String(data.prazo_execucao_dias));
    }
    addFixedAdditionalInfo(definition, value);
  });
}

function addFixedAdditionalInfo(definition, value) {
  const row = document.createElement('div');
  row.className = 'additional-info-row';
  row.dataset.fixedAdditionalInfo = definition.id;
  row.innerHTML = `
    <span class="additional-info-number">${escapeHtml(definition.number)}</span>
    <textarea rows="1" data-fixed-additional-text placeholder="Informação adicional">${escapeHtml(value)}</textarea>
    <button class="danger-action" type="button" title="Remover">x</button>
  `;
  row.querySelector('button').addEventListener('click', () => row.remove());
  fixedAdditionalInfoList.appendChild(row);
}

function addPriceTopic(topic = {}, withDefaultItem = false) {
  const defaults = DEFAULT_PRICE_TOPICS.find((item) => item.tipo === topic.tipo) || {};
  const title = topic.titulo || defaults.titulo || 'Novo t\u00f3pico';
  const type = topic.tipo || 'personalizado';
  const section = document.createElement('div');
  section.className = 'price-topic';
  section.dataset.topicType = type;
  section.dataset.defaultNcm = topic.ncm ?? defaults.ncm ?? '';
  section.dataset.defaultUn = topic.un ?? defaults.un ?? 'PC';
  section.innerHTML = `
    <div class="topic-heading">
      <input class="topic-title-input" data-topic-title value="${escapeHtml(title)}" placeholder="Nome do t\u00f3pico">
      <div class="topic-actions">
        <button class="small-action" type="button" data-add-topic-item>Adicionar</button>
        <button class="danger-action" type="button" data-remove-topic title="Remover t\u00f3pico">x</button>
      </div>
    </div>
    <div class="item-table"></div>
  `;

  section.querySelector('[data-add-topic-item]').addEventListener('click', () => addItem(section));
  section.querySelector('[data-remove-topic]').addEventListener('click', () => {
    section.remove();
    renumberItemCodes();
    recalculate();
  });

  priceTopics.appendChild(section);
  const items = Array.isArray(topic.itens) ? topic.itens : [];
  items.forEach((item) => addItem(section, item));
  if (withDefaultItem || !items.length) {
    addItem(section);
  }
  recalculate();
  return section;
}

function addItem(topicElement, values = {}) {
  const container = topicElement.querySelector('.item-table');
  if (!container.querySelector('.item-head')) {
    container.appendChild(createItemHeader());
  }

  const row = document.createElement('div');
  row.className = 'item-row';
  row.dataset.itemType = topicElement.dataset.topicType || 'personalizado';
  row.innerHTML = `
    <input data-field="item" value="${escapeHtml(values.item || nextItemCode())}">
    <input data-field="descricao" value="${escapeHtml(values.descricao || '')}">
    <input data-field="ncm" value="${escapeHtml(values.ncm ?? topicElement.dataset.defaultNcm ?? '')}">
    <input data-field="quant" type="text" inputmode="decimal" value="${escapeHtml(values.quant ?? 1)}">
    <input data-field="un" value="${escapeHtml(values.un || topicElement.dataset.defaultUn || 'PC')}">
    <input data-field="valor_unit" type="text" inputmode="decimal" value="${escapeHtml(values.valor_unit ?? 0)}">
    <input data-field="valor_total" readonly value="0,00">
    <button class="danger-action" type="button" title="Remover">x</button>
  `;
  row.querySelector('button').addEventListener('click', () => {
    row.remove();
    renumberItemCodes();
    recalculate();
  });
  container.appendChild(row);
  renumberItemCodes();
  recalculate();
}

function applyImportedData(result) {
  Object.entries(result.data || {}).forEach(([field, value]) => {
    if (form.elements[field] && value != null && String(value).trim() !== '') {
      form.elements[field].value = value;
    }
  });

  const importedData = result.data || {};
  const importedTopics = importedData.topicos_preco || result.topicos_preco;
  const topics = Array.isArray(importedTopics) && importedTopics.length
    ? importedTopics
    : legacyTopicsFromData({
      itens_servico: result.itens_servico || importedData.itens_servico,
      itens_consumiveis: result.itens_consumiveis || importedData.itens_consumiveis
    });

  if (topics.length) {
    priceTopics.innerHTML = '';
    topics.forEach((topic) => addPriceTopic(topic));
  }

  const descriptions = result.servicos_descricao || result.data?.servicos_descricao || [];
  if (descriptions.length) {
    servicesList.innerHTML = '';
    descriptions.forEach((item) => addServiceDescription(typeof item === 'string' ? item : item.item));
  }

  const technicalTeam = technicalTeamEntries(result.data || {});
  if (technicalTeam.length) {
    technicalTeamList.innerHTML = '';
    technicalTeam.forEach((item) => addTechnicalTeamMember(item));
  }

  const additionalInfo = result.informacoes_adicionais || result.data?.informacoes_adicionais || [];
  if (additionalInfo.length) {
    additionalInfoList.innerHTML = '';
    additionalInfo.forEach((item) => addAdditionalInfo(typeof item === 'string' ? item : item.item));
  }
  renderFixedAdditionalInfo(result.data || {});

  renumberItemCodes();
  recalculate();
}

function legacyTopicsFromData(data = {}) {
  return DEFAULT_PRICE_TOPICS.map((topic) => {
    const itens = topic.tipo === 'servico'
      ? data.itens_servico || []
      : data.itens_consumiveis || [];
    return { ...topic, itens };
  }).filter((topic) => topic.itens.length);
}

function createItemHeader() {
  const header = document.createElement('div');
  header.className = 'item-head';
  header.innerHTML = `
    <span>Item</span>
    <span>Descrição</span>
    <span>NCM</span>
    <span>Quant</span>
    <span>UN</span>
    <span>Valor unit.</span>
    <span>Total</span>
    <span></span>
  `;
  return header;
}

function nextItemCode() {
  const numbers = Array.from(document.querySelectorAll('.item-row [data-field="item"]'))
    .map((input) => Number.parseInt(input.value, 10))
    .filter(Number.isFinite);
  const next = numbers.length ? Math.max(...numbers) + 1 : 1;
  return String(next).padStart(4, '0');
}

function renumberItemCodes() {
  const rows = Array.from(priceTopics.querySelectorAll('.item-row'));
  rows.forEach((row, index) => {
    row.querySelector('[data-field="item"]').value = String(index + 1).padStart(4, '0');
  });
}

function recalculate() {
  let total = 0;
  document.querySelectorAll('.item-row').forEach((row) => {
    const quant = readNumber(row.querySelector('[data-field="quant"]').value);
    const unit = readNumber(row.querySelector('[data-field="valor_unit"]').value);
    const rowTotal = quant * unit;
    row.querySelector('[data-field="valor_total"]').value = numberFormatter.format(rowTotal);
    total += rowTotal;
  });

  grandTotal.textContent = moneyFormatter.format(total);
  form.elements.preco_total_numero.value = moneyFormatter.format(total);
  form.elements.preco_total_extenso.value = moneyToWords(total);
}

async function handleSubmit(event) {
  event.preventDefault();
  if (!window.supplyMarine) {
    showFormError(['API do Electron indisponível.']);
    return;
  }

  const validation = validateProposal();
  if (!validation.valid) {
    showFormError(validation.messages);
    validation.focusTarget?.focus();
    return;
  }

  const submitButton = document.querySelector('[form="proposal-form"][type="submit"]');
  submitButton.disabled = true;
  submitButton.textContent = 'Gerando...';

  try {
    const data = collectFormData();
    if (state.editingId) {
      data._historico_id = state.editingId;
    }
    const result = await window.supplyMarine.gerarDocx(data);
    state.lastOutputPath = result.outputPath;
    state.editingId = result.proposta?.id || state.editingId;
    await loadHistory();
    showResult('Proposta pronta', result.outputPath);
  } catch (error) {
    showFormError([friendlyError(error)]);
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = 'Gerar Proposta';
  }
}

function validateProposal() {
  clearValidation();
  const messages = [];
  let focusTarget = null;
  const requiredFields = [
    ['empresa_cliente', 'Informe a empresa cliente.'],
    ['numero_documento', 'Informe o número do documento.'],
    ['data_documento', 'Informe a data do documento.'],
    ['responsavel_nome', 'Informe o responsável.'],
    ['objeto', 'Informe o objeto da proposta.']
  ];

  requiredFields.forEach(([name, message]) => {
    const field = form.elements[name];
    if (!field || String(field.value || '').trim()) return;
    field.classList.add('invalid');
    messages.push(message);
    if (!focusTarget) focusTarget = field;
  });

  const descriptions = Array.from(document.querySelectorAll('[data-service-description]'))
    .filter((input) => input.value.trim());
  if (!descriptions.length) {
    messages.push('Adicione pelo menos uma descrição de serviço.');
    const field = document.querySelector('[data-service-description]');
    field?.classList.add('invalid');
    if (!focusTarget) focusTarget = field;
  }

  const items = collectPriceTopics().flatMap((topic) => topic.itens);
  if (!items.length) {
    messages.push('Adicione pelo menos um item de preço.');
    const field = priceTopics.querySelector('[data-field="descricao"]');
    field?.classList.add('invalid');
    if (!focusTarget) focusTarget = field;
  }

  document.querySelectorAll('.item-row').forEach((row) => {
    const descricao = readRowField(row, 'descricao');
    const unit = readNumber(readRowField(row, 'valor_unit'));
    if (!descricao && unit > 0) {
      row.querySelector('[data-field="descricao"]').classList.add('invalid');
      messages.push('Itens com valor precisam de descrição.');
      if (!focusTarget) focusTarget = row.querySelector('[data-field="descricao"]');
    }
  });

  return {
    valid: messages.length === 0,
    messages: [...new Set(messages)],
    focusTarget
  };
}

function clearValidation() {
  formAlert.hidden = true;
  formAlert.textContent = '';
  form.querySelectorAll('.invalid').forEach((field) => field.classList.remove('invalid'));
}

function showFormError(messages) {
  formAlert.innerHTML = messages.map((message) => `<div>${escapeHtml(message)}</div>`).join('');
  formAlert.hidden = false;
  formAlert.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function friendlyError(error) {
  const message = error?.message || String(error);
  if (message.includes('Template não encontrado') || message.includes('Template nao encontrado')) {
    return 'Modelo da proposta não encontrado. Gere novamente o template ou verifique a pasta src/templates.';
  }
  if (message.includes('EBUSY')) {
    return 'O arquivo está aberto em outro programa. Feche o documento ou gere uma nova proposta com outro número.';
  }
  if (message.includes('EACCES') || message.includes('EPERM')) {
    return 'Sem permissão para gravar o arquivo. Verifique se a pasta propostas está acessível.';
  }
  return message;
}

function collectFormData() {
  const data = {};
  new FormData(form).forEach((value, key) => {
    data[key] = String(value).trim();
  });

  data.servicos_descricao = Array.from(document.querySelectorAll('[data-service-description]'))
    .map((input) => input.value.trim())
    .filter(Boolean);
  data.equipe_tecnica_itens = Array.from(document.querySelectorAll('[data-technical-team]'))
    .map((input) => input.value.trim())
    .filter(Boolean);
  data.equipe_tecnica = data.equipe_tecnica_itens.join(' | ');
  data.informacoes_adicionais = Array.from(document.querySelectorAll('[data-additional-info]'))
    .map((input) => input.value.trim())
    .filter(Boolean);
  data.informacoes_adicionais_fixas = Object.fromEntries(
    Array.from(fixedAdditionalInfoList.querySelectorAll('[data-fixed-additional-text]'))
      .map((input) => [input.closest('[data-fixed-additional-info]')?.dataset.fixedAdditionalInfo, input.value.trim()])
      .filter(([id, value]) => id && value)
  );
  data.prazo_execucao_dias = extractExecutionTime(data.informacoes_adicionais_fixas['511']);
  data.topicos_preco = collectPriceTopics();
  data.itens_servico = data.topicos_preco
    .filter((topic) => topic.tipo === 'servico')
    .flatMap((topic) => topic.itens);
  data.itens_consumiveis = data.topicos_preco
    .filter((topic) => topic.tipo === 'consumivel')
    .flatMap((topic) => topic.itens);
  data.preco_total_numero = readNumber(form.elements.preco_total_numero.value);
  return data;
}

function collectPriceTopics() {
  return Array.from(priceTopics.querySelectorAll('.price-topic')).map((topicElement) => {
    const itens = collectItems(topicElement);
    const titulo = topicElement.querySelector('[data-topic-title]').value.trim();
    return {
      titulo: titulo || 'T\u00f3pico sem nome',
      tipo: topicElement.dataset.topicType || 'personalizado',
      itens,
      total: itens.reduce((sum, item) => sum + Number(item.valor_total || 0), 0)
    };
  }).filter((topic) => topic.itens.length);
}

function collectItems(container) {
  return Array.from(container.querySelectorAll('.item-row')).map((row) => ({
    item: readRowField(row, 'item'),
    codigo: '',
    descricao: readRowField(row, 'descricao'),
    ncm: readRowField(row, 'ncm'),
    quant: readNumber(readRowField(row, 'quant')),
    un: readRowField(row, 'un'),
    valor_unit: readNumber(readRowField(row, 'valor_unit')),
    valor_total: readNumber(readRowField(row, 'valor_total'))
  })).filter((item) => item.descricao || item.valor_unit || item.valor_total);
}

function readRowField(row, field) {
  return row.querySelector(`[data-field="${field}"]`).value.trim();
}

function readNumber(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }
  const cleaned = String(value || '0')
    .trim()
    .replace(/[^\d,.-]/g, '');
  const normalized = normalizeDecimalNumber(cleaned);
  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function normalizeDecimalNumber(value) {
  if (!value) {
    return '0';
  }

  const lastComma = value.lastIndexOf(',');
  const lastDot = value.lastIndexOf('.');

  if (lastComma !== -1 && lastDot !== -1) {
    const decimalSeparator = lastComma > lastDot ? ',' : '.';
    return removeThousands(value, decimalSeparator).replace(decimalSeparator, '.');
  }

  if (lastComma !== -1) {
    return value.replace(/\./g, '').replace(',', '.');
  }

  const dotCount = (value.match(/\./g) || []).length;
  if (dotCount > 1) {
    return value.replace(/\.(?=.*\.)/g, '');
  }

  return value;
}

function removeThousands(value, decimalSeparator) {
  return decimalSeparator === ','
    ? value.replace(/\./g, '')
    : value.replace(/,/g, '');
}

function showResult(title, path) {
  clearValidation();
  document.querySelector('#result-title').textContent = title;
  resultPath.textContent = path;
  pdfStatus.textContent = '';
  resultPanel.hidden = false;
  resultPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

async function loadHistory() {
  if (!window.supplyMarine) {
    return;
  }

  state.proposals = await window.supplyMarine.listarPropostas();
  renderHistory();
}

function renderHistory() {
  const query = historySearch.value.trim().toLowerCase();
  state.filteredProposals = state.proposals.filter((item) => {
    if (!query) return true;
    return [
      item.numero_documento,
      item.empresa_cliente,
      item.unidade || item.data?.unidade,
      item.data_documento
    ].some((value) => String(value || '').toLowerCase().includes(query));
  });

  emptyState.hidden = state.filteredProposals.length > 0;
  proposalList.innerHTML = state.filteredProposals.map((item) => `
    <div class="proposal-row" role="row">
      <span>${escapeHtml(item.numero_documento || '-')}</span>
      <span>${escapeHtml(item.empresa_cliente || '-')}</span>
      <span>${escapeHtml(item.unidade || item.data?.unidade || '-')}</span>
      <span>${escapeHtml(item.data_documento || '-')}</span>
      <span>${moneyFormatter.format(Number(item.preco_total_numero || 0))}</span>
      <div class="row-actions">
        <button class="small-action" type="button" data-history-action="pdf" data-id="${item.id}">PDF</button>
        <button class="small-action" type="button" data-history-action="open" data-id="${item.id}">Abrir</button>
        <button class="small-action" type="button" data-history-action="edit" data-id="${item.id}">Editar</button>
        <button class="small-action" type="button" data-history-action="duplicate" data-id="${item.id}">Duplicar</button>
        <button class="danger-action compact" type="button" data-history-action="delete" data-id="${item.id}" title="Excluir">x</button>
      </div>
    </div>
  `).join('');

  proposalList.querySelectorAll('[data-history-action]').forEach((button) => {
    button.addEventListener('click', () => handleHistoryAction(button.dataset.historyAction, button.dataset.id));
  });

  summaryCount.textContent = String(state.proposals.length);
  summaryLast.textContent = state.proposals[0]?.numero_documento || 'Nenhuma';
}

async function handleHistoryAction(action, id) {
  const proposal = state.proposals.find((item) => item.id === id);
  if (!proposal) return;

  if (action === 'open') {
    window.supplyMarine.abrirArquivo(proposal.docxPath);
    return;
  }

  if (action === 'pdf') {
    state.lastOutputPath = proposal.docxPath;
    showResult('Proposta selecionada', proposal.docxPath);
    exportPdf(proposal.docxPath);
    return;
  }

  if (action === 'edit') {
    loadProposalIntoForm(proposal, { duplicate: false });
    return;
  }

  if (action === 'duplicate') {
    loadProposalIntoForm(proposal, { duplicate: true });
    return;
  }

  if (action === 'delete') {
    const confirmed = window.confirm(`Excluir a proposta ${proposal.numero_documento || ''}?`);
    if (!confirmed) return;
    await window.supplyMarine.excluirProposta(id);
    await loadHistory();
  }
}

function loadProposalIntoForm(proposal, options = {}) {
  const data = proposal.data || {};
  state.editingId = options.duplicate ? '' : proposal.id;
  form.reset();
  servicesList.innerHTML = '';
  technicalTeamList.innerHTML = '';
  fixedAdditionalInfoList.innerHTML = '';
  additionalInfoList.innerHTML = '';
  priceTopics.innerHTML = '';
  resultPanel.hidden = true;
  state.lastOutputPath = options.duplicate ? '' : proposal.docxPath;

  fillFormFields({
    ...data,
    numero_documento: options.duplicate ? `${data.numero_documento || proposal.numero_documento}-COPIA` : data.numero_documento
  });

  (data.servicos_descricao || []).forEach((item) => addServiceDescription(typeof item === 'string' ? item : item.item));
  if (!servicesList.children.length) addServiceDescription('');

  technicalTeamEntries(data).forEach((item) => addTechnicalTeamMember(item));
  if (!technicalTeamList.children.length) addTechnicalTeamMember('');

  (data.informacoes_adicionais || []).forEach((item) => addAdditionalInfo(typeof item === 'string' ? item : item.item));
  if (!additionalInfoList.children.length) addAdditionalInfo('');
  renderFixedAdditionalInfo(data);

  const topics = Array.isArray(data.topicos_preco) && data.topicos_preco.length
    ? data.topicos_preco
    : legacyTopicsFromData(data);
  topics.forEach((topic) => addPriceTopic(topic));
  if (!priceTopics.querySelector('.price-topic')) {
    DEFAULT_PRICE_TOPICS.forEach((topic) => addPriceTopic({ ...topic }, true));
  }

  document.querySelector('#form-title').textContent = options.duplicate ? 'Duplicar proposta' : 'Editar proposta';
  recalculate();
  showView('form');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function fillFormFields(data) {
  Array.from(form.elements).forEach((field) => {
    if (!field.name || data[field.name] == null || Array.isArray(data[field.name])) return;
    field.value = data[field.name];
  });
}

function technicalTeamEntries(data = {}) {
  if (Array.isArray(data.equipe_tecnica_itens)) {
    return data.equipe_tecnica_itens.map((item) => String(item).trim()).filter(Boolean);
  }
  return String(data.equipe_tecnica || '')
    .split('|')
    .map((item) => item.trim())
    .filter(Boolean);
}

function extractExecutionTime(value) {
  const match = String(value || '').match(/estimados em\s+(.+?),\s+incluindo/i);
  return match ? match[1].trim() : '';
}

async function exportPdf(docxPath) {
  if (!docxPath || !window.supplyMarine) {
    return;
  }

  exportPdfButton.disabled = true;
  exportPdfButton.textContent = 'Exportando...';
  pdfStatus.textContent = 'Exportando PDF...';

  try {
    const result = await window.supplyMarine.exportarPdf(docxPath);
    if (result.converted) {
      pdfStatus.textContent = `PDF gerado: ${result.pdfPath}`;
      state.lastPdfPath = result.pdfPath;
      if (!result.skipOpen) {
        await window.supplyMarine.abrirArquivo(result.pdfPath);
      }
    } else {
      pdfStatus.textContent = result.openedDocx
        ? `${result.reason}. O DOCX foi aberto para imprimir/exportar em PDF pelo app padrão.`
        : `${result.reason}. Instale o LibreOffice ou deixe o Microsoft Word disponível para converter automaticamente.`;
    }
  } catch (error) {
    pdfStatus.textContent = error.message || String(error);
  } finally {
    exportPdfButton.disabled = false;
    exportPdfButton.textContent = 'Exportar PDF';
  }
}

function todayPtBr() {
  return new Intl.DateTimeFormat('pt-BR').format(new Date());
}

function addDaysPtBr(days) {
  const date = new Date();
  date.setDate(date.getDate() + days);
  return new Intl.DateTimeFormat('pt-BR').format(date);
}

function moneyToWords(value) {
  const totalCents = Math.round((Number(value) || 0) * 100);
  const reais = Math.floor(totalCents / 100);
  const cents = totalCents % 100;
  const parts = [];

  if (reais > 0) {
    parts.push(`${numberToWords(reais)} ${reais === 1 ? 'real' : 'reais'}`);
  }

  if (cents > 0) {
    parts.push(`${numberToWords(cents)} ${cents === 1 ? 'centavo' : 'centavos'}`);
  }

  if (!parts.length) {
    return 'Zero real';
  }

  return capitalize(parts.join(' e '));
}

function numberToWords(number) {
  const units = ['zero', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove'];
  const teens = ['dez', 'onze', 'doze', 'treze', 'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove'];
  const tens = ['', '', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa'];
  const hundreds = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos'];

  if (number < 10) return units[number];
  if (number < 20) return teens[number - 10];
  if (number < 100) {
    const ten = Math.floor(number / 10);
    const unit = number % 10;
    return unit ? `${tens[ten]} e ${units[unit]}` : tens[ten];
  }
  if (number === 100) return 'cem';
  if (number < 1000) {
    const hundred = Math.floor(number / 100);
    const rest = number % 100;
    return rest ? `${hundreds[hundred]} e ${numberToWords(rest)}` : hundreds[hundred];
  }
  if (number < 1000000) {
    const thousand = Math.floor(number / 1000);
    const rest = number % 1000;
    const prefix = thousand === 1 ? 'mil' : `${numberToWords(thousand)} mil`;
    return rest ? `${prefix} ${rest < 100 ? 'e ' : ''}${numberToWords(rest)}` : prefix;
  }

  const million = Math.floor(number / 1000000);
  const rest = number % 1000000;
  const prefix = `${numberToWords(million)} ${million === 1 ? 'milhão' : 'milhões'}`;
  return rest ? `${prefix} ${numberToWords(rest)}` : prefix;
}

function capitalize(value) {
  return value.charAt(0).toUpperCase() + value.slice(1);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
