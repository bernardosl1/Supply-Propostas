const state = {
  proposals: [],
  filteredProposals: [],
  models: [],
  lastOutputPath: '',
  editingId: '',
  draggingSection: null,
  draggingSubtopic: null
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
const resultPanel = document.querySelector('#result-panel');
const resultPath = document.querySelector('#result-path');
const pdfStatus = document.querySelector('#pdf-status');
const formAlert = document.querySelector('#form-alert');
const openDocxButton = document.querySelector('#open-docx');
const exportPdfButton = document.querySelector('#export-pdf');
const proposalList = document.querySelector('#proposal-list');
const emptyState = document.querySelector('#empty-state');
const summaryCount = document.querySelector('#summary-count');
const summaryLast = document.querySelector('#summary-last');
const historySearch = document.querySelector('#history-search');
const selectDocxButton = document.querySelector('#select-docx');
const docxStatus = document.querySelector('#docx-status');
const openAiKeyInput = document.querySelector('#openai-api-key');
const saveOpenAiKeyButton = document.querySelector('#save-openai-key');
const openAiStatus = document.querySelector('#openai-status');
const flexibleBlocks = document.querySelector('#proposal-sections');
const flexibleEmptyState = document.querySelector('#flexible-empty-state');
const saveProposalModelButton = document.querySelector('#save-proposal-model');
const modelList = document.querySelector('#model-list');
const modelEmptyState = document.querySelector('#model-empty-state');
const modelDialog = document.querySelector('#model-dialog');
const modelDialogForm = document.querySelector('#model-dialog-form');
const modelNameInput = document.querySelector('#model-name');
const modelCompanyInput = document.querySelector('#model-company');
const modelDialogError = document.querySelector('#model-dialog-error');
const previewProposalButton = document.querySelector('#preview-proposal');
const previewDialog = document.querySelector('#preview-dialog');
const previewContent = document.querySelector('#proposal-preview-content');

const FLEXIBLE_BLOCK_LABELS = {
  texto: 'T\u00f3pico',
  lista: 'Documenta\u00e7\u00e3o',
  preco: 'Pre\u00e7o',
  tabela: 'Tabela',
  quebra_pagina: 'Quebra de p\u00e1gina'
};

const DEFAULT_FIXED_SECTION_ORDER = ['dados_comerciais'];

const DEFAULT_PRICE_TOPICS = [
  { tipo: 'servico', titulo: 'Itens de servi\u00e7o', ncm: '-----', un: 'HH' },
  { tipo: 'consumivel', titulo: 'Consum\u00edveis', ncm: '-----', un: 'PC' }
];

const CUSTOM_TABLE_MIN_COLUMN_WIDTH = 96;
const CUSTOM_TABLE_WIDTH_PRECISION = 4;

init();

function init() {
  bindNavigation();
  bindProposalImport();
  bindForm();
  bindModels();
  bindPreview();
  bindSectionDragAndDrop();
  bindHistory();
  resetFormToDefaults();
  loadHistory();
  loadModels();
  loadOpenAiStatus();

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
      if (button.dataset.viewButton === 'models') loadModels();
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
  document.querySelectorAll('[data-add-flex-block]').forEach((button) => {
    button.addEventListener('click', () => addFlexibleBlock({ tipo: button.dataset.addFlexBlock }));
  });
  form.addEventListener('input', recalculate);
  form.addEventListener('submit', handleSubmit);
  openDocxButton.addEventListener('click', () => {
    if (state.lastOutputPath) {
      window.supplyMarine.abrirArquivo(state.lastOutputPath);
    }
  });
  exportPdfButton.addEventListener('click', () => exportPdf(state.lastOutputPath));
}

function bindModels() {
  saveProposalModelButton.addEventListener('click', openModelDialog);
  document.querySelector('#cancel-model-dialog').addEventListener('click', () => modelDialog.close());
  modelDialogForm.addEventListener('submit', handleSaveModel);
}

function bindPreview() {
  previewProposalButton.addEventListener('click', openProposalPreview);
  document.querySelector('#close-proposal-preview').addEventListener('click', () => previewDialog.close());
  document.querySelector('#generate-from-preview').addEventListener('click', () => {
    previewDialog.close();
    form.requestSubmit();
  });
}

function openProposalPreview() {
  const data = collectFormData();
  previewContent.innerHTML = renderProposalPreview(data);
  previewDialog.showModal();
  previewDialog.querySelector('.proposal-preview-scroll').scrollTop = 0;
}

function renderProposalPreview(data) {
  const blocks = Array.isArray(data.blocos_adicionais) ? data.blocos_adicionais : [];
  const sections = new Map();
  sections.set('dados_comerciais', { type: 'fixed', id: 'dados_comerciais' });
  blocks.forEach((block, index) => sections.set(`flex:${block.id || index}`, { type: 'flex', block }));

  const ordered = [];
  (data.ordem_secoes || []).forEach((id) => {
    if (!sections.has(id)) return;
    ordered.push(sections.get(id));
    sections.delete(id);
  });
  sections.forEach((section) => ordered.push(section));

  let sectionNumber = 1;
  const body = ordered.map((section) => {
    if (section.type === 'flex' && section.block.tipo === 'quebra_pagina') {
      return '<div class="preview-page-break"><span>Quebra de página</span></div>';
    }
    const html = section.type === 'fixed'
      ? renderFixedPreviewSection(section.id, sectionNumber, data)
      : renderFlexiblePreviewSection(section.block, sectionNumber);
    sectionNumber += 1;
    return html;
  }).join('');

  return `
    <header class="preview-document-header">
      <img src="../../assets/logo.ico" alt="Supply Marine">
      <div>
        <strong>${previewValue(data.empresa_cliente, 'Empresa não informada')}</strong>
        <span>${previewValue(data.numero_documento, 'Número não informado')} · ${previewValue(data.data_documento, 'Data não informada')}</span>
      </div>
    </header>
    ${body}
    <section class="preview-section">
      <p>Atenciosamente,</p>
      <p><strong>${previewValue(data.responsavel_nome, 'Responsável não informado')}</strong></p>
      ${data.responsavel_email ? `<p>${escapeHtml(data.responsavel_email)}</p>` : ''}
      ${data.responsavel_telefone ? `<p>${escapeHtml(data.responsavel_telefone)}</p>` : ''}
    </section>
  `;
}

function renderFixedPreviewSection(sectionId, number, data) {
  if (sectionId === 'dados_comerciais') {
    const fields = [
      ['Empresa', data.empresa_cliente],
      ['Nº do documento', data.numero_documento],
      ['Unidade', data.unidade],
      ['Data do documento', data.data_documento],
      ['Solicitação nº', data.processo_fluig],
      ['Responsável', data.responsavel_nome],
      ['Contato', data.solicitante_nome_cargo],
      ['E-mail', data.responsavel_email],
      ['E-mail do contato', data.contato_email],
      ['Telefone', data.responsavel_telefone || data.contato_telefone]
    ];
    return `
      <section class="preview-section">
        <h3 class="preview-section-title">${number}. Dados comerciais</h3>
        <div class="preview-commercial-grid">
          ${fields.map(([label, value]) => `<p><b>${escapeHtml(label)}:</b> ${previewValue(value)}</p>`).join('')}
        </div>
      </section>
    `;
  }
  return '';
}

function renderFlexiblePreviewSection(block, number) {
  const title = block.tipo === 'preco'
    ? String(block.titulo ?? 'Preço').trim()
    : block.tipo === 'lista'
      ? 'Documentação'
      : block.titulo || 'Tópico sem nome';
  let content = '';
  if (block.tipo === 'texto') {
    content = `${renderPreviewList(block.observacoes)}${renderPreviewSubtopics(block.subtopicos, String(number))}`;
  } else if (block.tipo === 'lista') {
    content = renderDocumentationPreview(block.linhas || []);
  } else if (block.tipo === 'tabela') {
    content = renderCustomTablePreview(block);
  } else if (block.tipo === 'preco') {
    content = renderPricePreview(block);
  }
  return `
    <section class="preview-section">
      ${title ? `<h3 class="preview-section-title">${number}. ${escapeHtml(title)}</h3>` : ''}
      ${content}
    </section>
  `;
}

function renderPreviewSubtopics(subtopics = [], parentNumber) {
  return subtopics.map((subtopic, index) => {
    const number = `${parentNumber}.${index + 1}`;
    const descendants = flattenSubtopicDescendants(subtopic.subtopicos || []);
    return `
      <h4 class="preview-subsection-title">${number} ${escapeHtml(subtopic.titulo || 'Subtópico sem nome')}</h4>
      ${renderPreviewList(subtopic.observacoes)}
      ${descendants.map((descendant, descendantIndex) => `
        <h4 class="preview-subsection-title preview-nested-subsection-title">${number}.${descendantIndex + 1} ${escapeHtml(descendant.titulo || 'Subtópico sem nome')}</h4>
        ${renderPreviewList(descendant.observacoes)}
      `).join('')}
    `;
  }).join('');
}

function flattenSubtopicDescendants(subtopics = []) {
  return (Array.isArray(subtopics) ? subtopics : []).flatMap((subtopic) => [
    { ...subtopic, subtopicos: [] },
    ...flattenSubtopicDescendants(subtopic?.subtopicos)
  ]);
}

function normalizeSubtopicHierarchy(subtopics = []) {
  return (Array.isArray(subtopics) ? subtopics : []).map((subtopic) => ({
    ...subtopic,
    subtopicos: flattenSubtopicDescendants(subtopic?.subtopicos)
  }));
}

function renderDocumentationPreview(rows) {
  if (!rows.length) return previewEmpty();
  return `
    <table class="preview-table">
      <thead><tr><th>Descrição</th><th>Nº do documento</th><th>Data</th></tr></thead>
      <tbody>${rows.map((row) => `<tr><td>${previewValue(row.descricao)}</td><td>${previewValue(row.numero_documento)}</td><td>${previewValue(row.data)}</td></tr>`).join('')}</tbody>
    </table>
  `;
}

function renderCustomTablePreview(block) {
  const columns = normalizeFlexibleTableColumns(
    Array.isArray(block.colunas) ? block.colunas : [],
    { createDefaults: false }
  );
  const rows = Array.isArray(block.linhas) ? block.linhas : [];
  if (!columns.length) return previewEmpty();
  const totalWidth = columns.reduce((sum, column) => sum + column.largura, 0);
  return `
    <table class="preview-table">
      <colgroup>${columns.map((column) => `<col style="width: ${(column.largura / totalWidth) * 100}%">`).join('')}</colgroup>
      <thead><tr>${columns.map((column) => `<th>${previewValue(column.nome, 'Coluna')}</th>`).join('')}</tr></thead>
      <tbody>${rows.map((row) => `<tr>${columns.map((column) => `<td>${previewValue(row.valores?.[column.id])}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>
    ${rows.length ? '' : previewEmpty()}
  `;
}

function renderPricePreview(block) {
  const topics = Array.isArray(block.topicos_preco) ? block.topicos_preco : [];
  const topicHtml = topics.map((topic) => `
    <h4 class="preview-subsection-title">${escapeHtml(topic.titulo || 'Itens')}</h4>
    <table class="preview-table">
      <thead><tr><th>Item</th><th>Descrição</th><th>NCM</th><th>Quant.</th><th>UN</th><th>Valor unit.</th><th>Total</th></tr></thead>
      <tbody>${(topic.itens || []).map((item) => `<tr><td>${previewValue(item.item)}</td><td>${previewValue(item.descricao)}</td><td>${previewValue(item.ncm)}</td><td>${previewValue(item.quant)}</td><td>${previewValue(item.un)}</td><td>${moneyFormatter.format(Number(item.valor_unit || 0))}</td><td>${moneyFormatter.format(Number(item.valor_total || 0))}</td></tr>`).join('')}</tbody>
    </table>
  `).join('');
  const terms = [
    ['Validade da proposta', block.validade_proposta],
    ['Pagamento', block.pagamento],
    ['Prazo de entrega', block.prazo_entrega],
    ['Frete', block.frete],
    ['Impostos', block.impostos]
  ].filter(([, value]) => value);
  return `
    ${topicHtml || previewEmpty()}
    <p class="preview-price-total">Preço total: ${moneyFormatter.format(Number(block.preco_total_numero || 0))}${block.preco_total_extenso ? ` (${escapeHtml(block.preco_total_extenso)})` : ''}</p>
    ${terms.map(([label, value]) => `<p><strong>${escapeHtml(label)}:</strong> ${escapeHtml(value)}</p>`).join('')}
  `;
}

function renderPreviewList(items = []) {
  const normalized = items.map((item) => typeof item === 'string' ? item : item?.item || '').filter(Boolean);
  return normalized.length ? `<ul>${normalized.map((item) => `<li>${previewMultiline(item)}</li>`).join('')}</ul>` : previewEmpty();
}

function previewMultiline(value) {
  const text = String(value || '').trim();
  return text ? escapeHtml(text).replace(/\r?\n/g, '<br>') : previewEmpty();
}

function previewValue(value, fallback = 'Não informado') {
  const text = String(value ?? '').trim();
  return text ? escapeHtml(text) : `<span class="preview-empty-value">${escapeHtml(fallback)}</span>`;
}

function previewEmpty() {
  return '<span class="preview-empty-value">Nenhum conteúdo informado</span>';
}

function openModelDialog() {
  modelDialogForm.reset();
  modelDialogError.hidden = true;
  modelDialogError.textContent = '';
  modelCompanyInput.value = form.elements.empresa_cliente?.value.trim() || '';
  modelDialog.showModal();
  modelNameInput.focus();
}

async function handleSaveModel(event) {
  event.preventDefault();
  const nome = modelNameInput.value.trim();
  const empresa = modelCompanyInput.value.trim();
  if (!nome || !empresa) {
    modelDialogError.textContent = 'Informe o nome do modelo e a empresa.';
    modelDialogError.hidden = false;
    (!nome ? modelNameInput : modelCompanyInput).focus();
    return;
  }

  const submitButton = modelDialogForm.querySelector('[type="submit"]');
  submitButton.disabled = true;
  submitButton.textContent = 'Salvando...';
  try {
    await window.supplyMarine.salvarModelo({ nome, empresa, estrutura: collectFormData() });
    modelDialog.close();
    await loadModels();
    showView('models');
  } catch (error) {
    modelDialogError.textContent = error.message || String(error);
    modelDialogError.hidden = false;
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = 'Salvar modelo';
  }
}

function bindSectionDragAndDrop() {
  flexibleBlocks.addEventListener('dragstart', (event) => {
    const subtopicHandle = event.target.closest('.subtopic-drag-handle');
    const subtopic = subtopicHandle?.closest('.flex-subtopic');
    if (subtopic) {
      state.draggingSubtopic = subtopic;
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', subtopic.dataset.subtopicId || 'subtopic');
      requestAnimationFrame(() => subtopic.classList.add('dragging'));
      return;
    }
    const handle = event.target.closest('.section-drag-handle');
    const section = handle?.closest('.proposal-section');
    if (!section || section.parentElement !== flexibleBlocks) {
      event.preventDefault();
      return;
    }
    state.draggingSection = section;
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', section.dataset.proposalSection || 'proposal-section');
    requestAnimationFrame(() => section.classList.add('dragging'));
  });

  flexibleBlocks.addEventListener('dragover', (event) => {
    if (state.draggingSubtopic) {
      const targetSubtopic = event.target.closest('.flex-subtopic');
      const targetContainer = targetSubtopic?.parentElement || event.target.closest('[data-flex-subtopics]');
      if (!targetContainer
        || !targetContainer.hasAttribute('data-flex-subtopics')
        || subtopicContainerDepth(targetContainer) > 1
        || targetSubtopic === state.draggingSubtopic
        || state.draggingSubtopic.contains(targetContainer)) return;
      event.preventDefault();
      event.dataTransfer.dropEffect = 'move';
      if (!targetSubtopic) {
        targetContainer.appendChild(state.draggingSubtopic);
      } else {
        const targetRect = targetSubtopic.getBoundingClientRect();
        const insertBefore = event.clientY < targetRect.top + (targetRect.height / 2);
        targetContainer.insertBefore(state.draggingSubtopic, insertBefore ? targetSubtopic : targetSubtopic.nextElementSibling);
      }
      renumberFlexibleTopics();
      return;
    }
    const dragging = state.draggingSection;
    const target = event.target.closest('.proposal-section');
    if (!dragging || !target || target === dragging || target.parentElement !== flexibleBlocks) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
    const targetRect = target.getBoundingClientRect();
    const insertBefore = event.clientY < targetRect.top + (targetRect.height / 2);
    flexibleBlocks.insertBefore(dragging, insertBefore ? target : target.nextElementSibling);
    renumberFlexibleTopics();
  });

  flexibleBlocks.addEventListener('drop', (event) => {
    if (!state.draggingSection && !state.draggingSubtopic) return;
    event.preventDefault();
    finishSectionDrag();
  });

  flexibleBlocks.addEventListener('dragend', finishSectionDrag);
}

function subtopicContainerDepth(container) {
  let depth = 0;
  let owner = container?.closest('.flex-subtopic');
  while (owner) {
    depth += 1;
    owner = owner.parentElement?.closest('.flex-subtopic');
  }
  return depth;
}

function finishSectionDrag() {
  if (state.draggingSection) state.draggingSection.classList.remove('dragging');
  if (state.draggingSubtopic) state.draggingSubtopic.classList.remove('dragging');
  state.draggingSection = null;
  state.draggingSubtopic = null;
  renumberFlexibleTopics();
}

function bindProposalImport() {
  selectDocxButton.addEventListener('click', handleSelectDocx);
  saveOpenAiKeyButton.addEventListener('click', handleSaveOpenAiKey);
}

async function loadOpenAiStatus() {
  if (!window.supplyMarine?.getOpenAiStatus) {
    openAiStatus.textContent = 'Configuração da OpenAI indisponível.';
    return;
  }
  try {
    const status = await window.supplyMarine.getOpenAiStatus();
    selectDocxButton.disabled = !status.configured;
    openAiStatus.textContent = status.configured
      ? `GPT-5.6 Luna configurado${status.source === 'environment' ? ' pelo ambiente' : ''}.`
      : 'Informe sua chave da API OpenAI para habilitar a importação.';
  } catch (error) {
    selectDocxButton.disabled = true;
    openAiStatus.textContent = error.message || String(error);
  }
}

async function handleSaveOpenAiKey() {
  const apiKey = openAiKeyInput.value.trim();
  saveOpenAiKeyButton.disabled = true;
  openAiStatus.textContent = apiKey ? 'Protegendo a chave...' : 'Removendo a chave...';
  try {
    const status = await window.supplyMarine.salvarOpenAiKey(apiKey);
    openAiKeyInput.value = '';
    selectDocxButton.disabled = !status.configured;
    openAiStatus.textContent = status.configured
      ? 'Chave salva com criptografia. GPT-5.6 Luna pronto.'
      : 'Chave removida. Informe uma chave para importar.';
  } catch (error) {
    openAiStatus.textContent = error.message || String(error);
  } finally {
    saveOpenAiKeyButton.disabled = false;
  }
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
  docxStatus.textContent = 'GPT-5.6 Luna está lendo a proposta...';

  try {
    const result = await window.supplyMarine.importarDocx(filePath);
    resetFormToDefaults();
    applyImportedData(result);
    state.editingId = '';
    state.lastOutputPath = '';
    resultPanel.hidden = true;
    document.querySelector('#form-title').textContent = 'Proposta importada';
    const doubtfulFields = result.campos_duvidosos || [];
    const doubtful = doubtfulFields.length;
    const tokens = result.ai?.total_tokens || 0;
    const details = [
      tokens ? `${tokens.toLocaleString('pt-BR')} tokens` : '',
      doubtful ? `revisar: ${doubtfulFields.join(', ')}` : 'nenhuma dúvida sinalizada'
    ].filter(Boolean).join(' · ');
    docxStatus.textContent = `Importado com GPT-5.6 Luna: ${result.data?.numero_documento || filePath}${details ? ` · ${details}` : ''}`;
  } catch (error) {
    docxStatus.textContent = error.message || String(error);
  } finally {
    selectDocxButton.disabled = false;
  }
}

function fillDefaults() {
  setValue('data_documento', todayPtBr());
  setValue('moeda', 'Real R$');
}

function resetFormToDefaults() {
  state.editingId = '';
  form.reset();
  clearFlexibleBlocks();
  restoreDefaultFixedSectionOrder();
  updateFlexibleEmptyState();
  resultPanel.hidden = true;
  state.lastOutputPath = '';
  document.querySelector('#form-title').textContent = 'Formulário da proposta';

  fillDefaults();
  recalculate();
}

function setValue(name, value) {
  const field = form.elements[name];
  if (field) {
    field.value = value;
  }
}

function addFlexibleBlock(data = {}, afterElement = null) {
  const type = Object.prototype.hasOwnProperty.call(FLEXIBLE_BLOCK_LABELS, data.tipo)
    ? data.tipo
    : 'texto';
  const block = document.createElement('article');
  block.className = 'form-section proposal-section flexible-block';
  block.dataset.flexType = type;
  block.dataset.flexId = data.id || createFlexibleId('bloco');
  block.dataset.proposalSection = `flex:${block.dataset.flexId}`;
  const normalizedTitle = type === 'texto' ? stripAutomaticNumber(data.titulo || '') : data.titulo || '';
  const titleField = type === 'quebra_pagina'
    ? '<span></span>'
    : (type === 'lista' || type === 'preco'
      ? `<div class="flex-topic-title"><strong data-flex-section-number></strong><span class="flex-fixed-section-title">${type === 'preco' ? 'PRE\u00c7O' : 'DOCUMENTA\u00c7\u00c3O'}</span></div>`
      : `<div class="flex-topic-title"><strong data-flex-section-number></strong><input data-flex-title value="${escapeHtml(normalizedTitle)}" placeholder="${type === 'texto' ? 'Nome do t\u00f3pico' : 'T\u00edtulo da tabela'}"></div>`);

  const moveButton = '<button class="section-drag-handle" type="button" draggable="true" title="Arrastar seção" aria-label="Arrastar seção">\u22ee\u22ee</button>';
  const managementButtons = `
    <button class="small-action" type="button" data-flex-action="up" title="Mover para cima">\u2191</button>
    <button class="small-action" type="button" data-flex-action="down" title="Mover para baixo">\u2193</button>
    <button class="small-action" type="button" data-flex-action="duplicate">Duplicar</button>
    <button class="danger-action" type="button" data-flex-action="delete" title="Remover">x</button>
  `;
  const topicCreationButtons = type === 'texto'
    ? `
      <button class="small-action" type="button" data-add-subtopic>+ Subt\u00f3pico</button>
      <button class="small-action" type="button" data-add-topic-observation>+ Observa\u00e7\u00e3o</button>
    `
    : '';
  const actionButtons = `${topicCreationButtons}${moveButton}${managementButtons}`;
  block.innerHTML = type === 'preco'
    ? `
      <div class="section-heading additional-price-heading">
        <div class="additional-price-title flex-topic-title">
          <strong data-flex-section-number></strong>
          <input data-flex-title value="${escapeHtml(data.titulo ?? 'ITENS / PREÇO')}" placeholder="Título opcional">
        </div>
        <div class="price-actions">
          <button class="small-action" type="button" data-add-flex-price-topic>Adicionar tópico</button>
          <div class="total-box">
            <span>Total</span>
            <strong data-flex-grand-total>R$ 0,00</strong>
          </div>
          <div class="flexible-block-actions">${managementButtons}${moveButton}</div>
        </div>
      </div>
      <div class="flexible-block-body"></div>
    `
    : `
      <div class="flexible-block-header${type === 'texto' || type === 'lista' || type === 'tabela' ? ' fixed-style-flexible-header' : ''}">
        <span class="flexible-block-kind">${FLEXIBLE_BLOCK_LABELS[type]}</span>
        ${titleField}
        <div class="flexible-block-actions">${actionButtons}</div>
      </div>
      <div class="flexible-block-body"></div>
    `;

  renderFlexibleBlockBody(block, data);
  block.querySelector('[data-flex-action="up"]').addEventListener('click', () => {
    if (block.previousElementSibling) {
      flexibleBlocks.insertBefore(block, block.previousElementSibling);
      renumberFlexibleTopics();
    }
  });
  block.querySelector('[data-flex-action="down"]').addEventListener('click', () => {
    if (block.nextElementSibling) {
      flexibleBlocks.insertBefore(block.nextElementSibling, block);
      renumberFlexibleTopics();
    }
  });
  block.querySelector('[data-flex-action="duplicate"]').addEventListener('click', () => {
    const copy = collectFlexibleBlock(block);
    copy.id = createFlexibleId('bloco');
    addFlexibleBlock(copy, block);
  });
  block.querySelector('[data-flex-action="delete"]').addEventListener('click', () => {
    block.remove();
    updateFlexibleEmptyState();
    renumberFlexibleTopics();
  });

  if (afterElement?.parentElement === flexibleBlocks) {
    flexibleBlocks.insertBefore(block, afterElement.nextElementSibling);
  } else {
    flexibleBlocks.appendChild(block);
  }
  updateFlexibleEmptyState();
  renumberFlexibleTopics();
  recalculate();
  return block;
}

function renderFlexibleBlockBody(block, data = {}) {
  const body = block.querySelector('.flexible-block-body');
  const type = block.dataset.flexType;

  if (type === 'texto') {
    body.innerHTML = `
      <div class="topic-observations" data-topic-observations></div>
      <div class="flex-subtopics" data-flex-subtopics></div>
    `;
    const observations = Array.isArray(data.observacoes) && data.observacoes.length
      ? data.observacoes
      : (data.conteudo ? String(data.conteudo).split(/\r?\n/).filter(Boolean) : []);
    observations.forEach((observation) => addTopicObservation(body.querySelector('[data-topic-observations]'), observation));
    const subtopicsContainer = body.querySelector('[data-flex-subtopics]');
    normalizeSubtopicHierarchy(data.subtopicos || []).forEach((subtopic) => addFlexibleSubtopic(block, subtopic, subtopicsContainer));
    block.querySelector('[data-add-topic-observation]').addEventListener('click', () => {
      addTopicObservation(body.querySelector('[data-topic-observations]'), '');
    });
    block.querySelector('[data-add-subtopic]').addEventListener('click', () => addFlexibleSubtopic(block, {}));
    return;
  }

  if (type === 'lista') {
    body.innerHTML = `
      <div class="document-list-editor">
        <table class="document-list-table">
          <thead>
            <tr>
              <th>DESCRI\u00c7\u00c3O</th>
              <th>N\u00ba DO DOCUMENTO</th>
              <th>DATA</th>
              <th></th>
            </tr>
          </thead>
          <tbody data-flex-list></tbody>
        </table>
        <button class="small-action" type="button" data-add-flex-list-item>Adicionar linha</button>
      </div>
    `;
    const rows = Array.isArray(data.linhas) && data.linhas.length
      ? data.linhas
      : (Array.isArray(data.itens) && data.itens.length
        ? data.itens.map((item) => ({ descricao: typeof item === 'string' ? item : item?.descricao || item?.item || '' }))
        : [{}]);
    rows.forEach((row) => addFlexibleListItem(body, row));
    body.querySelector('[data-add-flex-list-item]').addEventListener('click', () => addFlexibleListItem(body, {}));
    return;
  }

  if (type === 'tabela') {
    const tableData = normalizeFlexibleTable(data);
    renderFlexibleTableEditor(block, tableData);
    return;
  }

  if (type === 'preco') {
    renderFlexiblePriceEditor(block, data);
    return;
  }

  body.innerHTML = '<div class="page-break-preview">O conte\u00fado seguinte come\u00e7ar\u00e1 em uma nova p\u00e1gina.</div>';
}

function renderFlexiblePriceEditor(block, data = {}) {
  const body = block.querySelector('.flexible-block-body');
  const valueOrDefault = (field, fallback = '') => data[field] ?? form.elements[field]?.value ?? fallback;
  body.innerHTML = `
    <div class="flex-price-editor" data-flex-price-editor>
      <div class="price-topics" data-flex-price-topics></div>
      <div class="price-terms-grid">
        <label>Pre\u00e7o total<input data-flex-price-field="preco_total_numero" data-money></label>
        <label>Pre\u00e7o por extenso<input data-flex-price-field="preco_total_extenso"></label>
        <label>Moeda<input data-flex-price-field="moeda" value="${escapeHtml(valueOrDefault('moeda', 'Real R$'))}"></label>
        <label>Validade da proposta<input data-flex-price-field="validade_proposta" value="${escapeHtml(formatDateInput(valueOrDefault('validade_proposta')))}" placeholder="dd/mm/aaaa" inputmode="numeric" maxlength="10" autocomplete="off"></label>
        <label>Pagamento<input data-flex-price-field="pagamento" value="${escapeHtml(valueOrDefault('pagamento'))}"></label>
        <label>Prazo de entrega<input data-flex-price-field="prazo_entrega" value="${escapeHtml(valueOrDefault('prazo_entrega'))}"></label>
        <label>Frete<input data-flex-price-field="frete" value="${escapeHtml(valueOrDefault('frete'))}"></label>
        <label>Impostos<input data-flex-price-field="impostos" value="${escapeHtml(valueOrDefault('impostos'))}"></label>
      </div>
    </div>
  `;

  const topicsContainer = body.querySelector('[data-flex-price-topics]');
  const topics = Array.isArray(data.topicos_preco)
    ? data.topicos_preco
    : DEFAULT_PRICE_TOPICS.map((topic) => ({ ...topic }));
  topics.forEach((topic) => addPriceTopic(topic, !Array.isArray(topic.itens) || !topic.itens.length, topicsContainer));
  block.querySelector('[data-add-flex-price-topic]').addEventListener('click', () => addPriceTopic({}, false, topicsContainer));
  bindDateMask(body.querySelector('[data-flex-price-field="validade_proposta"]'));
  recalculate();
}

function addTopicObservation(container, value, options = {}) {
  const row = document.createElement('div');
  row.className = 'topic-observation-row';
  row.innerHTML = `
    <span class="topic-bullet">\u2022</span>
    <input data-topic-observation value="${escapeHtml(typeof value === 'string' ? value : value?.texto || '')}" placeholder="${escapeHtml(options.placeholder || 'Observa\u00e7\u00e3o')}">
    <button class="danger-action" type="button" title="Remover observa\u00e7\u00e3o">x</button>
  `;
  row.querySelector('button').addEventListener('click', () => {
    if (options.keepOne && container.querySelectorAll(':scope > .topic-observation-row').length <= 1) {
      const input = row.querySelector('[data-topic-observation]');
      input.value = '';
      input.focus();
      return;
    }
    row.remove();
  });
  container.appendChild(row);
  return row;
}

function addFlexibleSubtopic(block, data = {}, targetContainer = null) {
  const container = targetContainer || block.querySelector('[data-flex-subtopics]');
  if (!container) return null;
  const isNestedSubtopic = Boolean(container.closest('.flex-subtopic'));
  const subtopic = document.createElement('div');
  subtopic.className = 'flex-subtopic';
  subtopic.dataset.subtopicId = data.id || createFlexibleId('subtopico');
  subtopic.innerHTML = `
    <div class="flex-subtopic-heading">
      <strong data-flex-subtopic-number></strong>
      <input data-flex-subtopic-title value="${escapeHtml(stripAutomaticNumber(data.titulo || ''))}" placeholder="Nome do subt\u00f3pico">
      <div class="flex-subtopic-actions">
        <button class="small-action" type="button" data-add-nested-subtopic title="${isNestedSubtopic ? 'Adicionar no mesmo n\u00edvel' : 'Adicionar n\u00edvel abaixo'}">${isNestedSubtopic ? '+ Mesmo n\u00edvel' : '+ N\u00edvel abaixo'}</button>
        <button class="small-action" type="button" data-add-subtopic-observation title="Adicionar observa\u00e7\u00e3o">+ Observa\u00e7\u00e3o</button>
        <button class="subtopic-drag-handle" type="button" draggable="true" title="Arrastar subt\u00f3pico" aria-label="Arrastar subt\u00f3pico">\u22ee\u22ee</button>
        <button class="danger-action" type="button" data-remove-subtopic title="Remover subt\u00f3pico">x</button>
      </div>
    </div>
    <div class="topic-observations" data-subtopic-observations></div>
    <div class="flex-subtopics" data-flex-subtopics></div>
  `;
  const observations = Array.isArray(data.observacoes) ? data.observacoes : [];
  observations.forEach((observation) => addTopicObservation(subtopic.querySelector('[data-subtopic-observations]'), observation));
  const nestedContainer = subtopic.querySelector('[data-flex-subtopics]');
  (Array.isArray(data.subtopicos) ? data.subtopicos : []).forEach((nested) => {
    addFlexibleSubtopic(block, nested, nestedContainer);
  });
  subtopic.querySelector('[data-add-nested-subtopic]').addEventListener('click', () => {
    addFlexibleSubtopic(block, {}, isNestedSubtopic ? container : nestedContainer);
  });
  subtopic.querySelector('[data-add-subtopic-observation]').addEventListener('click', () => {
    addTopicObservation(subtopic.querySelector('[data-subtopic-observations]'), '');
  });
  subtopic.querySelector('[data-remove-subtopic]').addEventListener('click', () => {
    subtopic.remove();
    renumberFlexibleTopics();
  });
  container.appendChild(subtopic);
  renumberFlexibleTopics();
  return subtopic;
}

function addFlexibleListItem(body, value = {}) {
  const list = body.querySelector('[data-flex-list]');
  const rowData = typeof value === 'string' ? { descricao: value } : value;
  const row = document.createElement('tr');
  row.innerHTML = `
    <td><input data-flex-list-field="descricao" value="${escapeHtml(rowData.descricao || '')}" placeholder="Descri\u00e7\u00e3o"></td>
    <td><input data-flex-list-field="numero_documento" value="${escapeHtml(rowData.numero_documento || '')}" placeholder="N\u00ba do documento"></td>
    <td><input data-flex-list-field="data" value="${escapeHtml(formatDateInput(rowData.data || ''))}" placeholder="dd/mm/aaaa" inputmode="numeric" maxlength="10" autocomplete="off"></td>
    <td><button class="danger-action" type="button" title="Remover linha">x</button></td>
  `;
  bindDateMask(row.querySelector('[data-flex-list-field="data"]'));
  row.querySelector('button').addEventListener('click', () => row.remove());
  list.appendChild(row);
}

function bindDateMask(input) {
  if (!input) return;
  input.addEventListener('input', () => {
    input.value = formatDateInput(input.value);
  });
}

function formatDateInput(value) {
  const digits = String(value || '').replace(/\D/g, '').slice(0, 8);
  if (digits.length <= 2) return digits;
  if (digits.length <= 4) return `${digits.slice(0, 2)}/${digits.slice(2)}`;
  return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
}

function normalizeFlexibleTableColumns(rawColumns, options = {}) {
  const createDefaults = options.createDefaults !== false;
  const source = Array.isArray(rawColumns) && rawColumns.length
    ? rawColumns
    : (createDefaults
      ? [
        { id: createFlexibleId('coluna'), nome: 'Coluna 1' },
        { id: createFlexibleId('coluna'), nome: 'Coluna 2' }
      ]
      : []);
  if (!source.length) return [];

  const columns = source.map((column, index) => ({
    id: column?.id || createFlexibleId(`coluna_${index + 1}`),
    nome: column?.nome || '',
    largura: Number(column?.largura)
  }));
  const validWidths = columns
    .map((column) => column.largura)
    .filter((width) => Number.isFinite(width) && width > 0);
  const fallbackWeight = validWidths.length
    ? validWidths.reduce((sum, width) => sum + width, 0) / validWidths.length
    : 100 / columns.length;
  const weights = columns.map((column) => (
    Number.isFinite(column.largura) && column.largura > 0 ? column.largura : fallbackWeight
  ));
  const totalWeight = weights.reduce((sum, width) => sum + width, 0);
  let allocated = 0;

  return columns.map((column, index) => {
    const width = index === columns.length - 1
      ? 100 - allocated
      : Number(((weights[index] / totalWeight) * 100).toFixed(CUSTOM_TABLE_WIDTH_PRECISION));
    allocated += width;
    return { ...column, largura: Number(width.toFixed(CUSTOM_TABLE_WIDTH_PRECISION)) };
  });
}

function normalizeFlexibleTable(data = {}) {
  const columns = normalizeFlexibleTableColumns(data.colunas);
  const rows = Array.isArray(data.linhas) && data.linhas.length ? data.linhas : [{}];
  return { colunas: columns, linhas: rows };
}

function renderFlexibleTableEditor(block, data) {
  data = {
    ...data,
    colunas: normalizeFlexibleTableColumns(data.colunas)
  };
  const body = block.querySelector('.flexible-block-body');
  body.innerHTML = `
    <div class="custom-table-editor">
      <div class="custom-table-actions">
        <span class="custom-table-resize-hint">Arraste as divis\u00f3rias para ajustar as colunas</span>
        <button class="small-action" type="button" data-custom-table-action="column">Adicionar coluna</button>
        <button class="small-action" type="button" data-custom-table-action="row">Adicionar linha</button>
      </div>
      <table class="custom-table">
        <colgroup>
          ${data.colunas.map((column) => `<col data-custom-column-width="${escapeHtml(column.id)}">`).join('')}
        </colgroup>
        <thead><tr></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  `;
  const head = body.querySelector('thead tr');
  const tableBody = body.querySelector('tbody');

  const table = body.querySelector('.custom-table');
  table.style.minWidth = `${Math.max(
    560,
    (data.colunas.length * (CUSTOM_TABLE_MIN_COLUMN_WIDTH + 32)) + 16
  )}px`;

  data.colunas.forEach((column, columnIndex) => {
    const cell = document.createElement('th');
    cell.dataset.customColumnId = column.id;
    cell.innerHTML = `
      <div class="custom-column-heading">
        <div class="custom-column-field">
          <input data-custom-column-name data-column-id="${escapeHtml(column.id)}" data-column-width="${column.largura}" value="${escapeHtml(column.nome)}" placeholder="Nome da coluna">
          ${columnIndex < data.colunas.length - 1
            ? `<span class="custom-column-resizer" data-custom-column-resizer="${escapeHtml(column.id)}" role="separator" tabindex="0" aria-orientation="vertical" aria-label="Redimensionar coluna ${columnIndex + 1}" title="Arraste para ajustar a largura"></span>`
            : ''}
        </div>
        <button class="danger-action" type="button" data-remove-custom-column="${escapeHtml(column.id)}" title="Remover coluna">x</button>
      </div>
    `;
    head.appendChild(cell);
  });

  data.linhas.forEach((rowData) => {
    const row = document.createElement('tr');
    data.colunas.forEach((column, columnIndex) => {
      const cell = document.createElement('td');
      const value = rowData?.valores?.[column.id] ?? rowData?.[column.id] ?? '';
      cell.innerHTML = `
        <div class="custom-table-cell-layout">
          <input data-custom-cell data-column-id="${escapeHtml(column.id)}" value="${escapeHtml(value)}">
          ${columnIndex === data.colunas.length - 1
            ? '<button class="danger-action" type="button" data-remove-custom-row title="Remover linha">x</button>'
            : '<span aria-hidden="true"></span>'}
        </div>
      `;
      row.appendChild(cell);
    });
    row.querySelector('[data-remove-custom-row]').addEventListener('click', () => row.remove());
    tableBody.appendChild(row);
  });
  applyCustomTableColumnLayout(table);

  body.querySelector('[data-custom-table-action="column"]').addEventListener('click', () => {
    const current = readFlexibleTable(block, { keepEmptyRows: true });
    const averageWidth = current.colunas.reduce((sum, column) => sum + column.largura, 0) / current.colunas.length;
    current.colunas.push({
      id: createFlexibleId('coluna'),
      nome: `Coluna ${current.colunas.length + 1}`,
      largura: averageWidth
    });
    renderFlexibleTableEditor(block, current);
  });
  body.querySelector('[data-custom-table-action="row"]').addEventListener('click', () => {
    const current = readFlexibleTable(block, { keepEmptyRows: true });
    current.linhas.push({ valores: {} });
    renderFlexibleTableEditor(block, current);
  });
  body.querySelectorAll('[data-remove-custom-column]').forEach((button) => {
    button.addEventListener('click', () => {
      const current = readFlexibleTable(block, { keepEmptyRows: true });
      if (current.colunas.length <= 1) return;
      const columnId = button.dataset.removeCustomColumn;
      current.colunas = current.colunas.filter((column) => column.id !== columnId);
      current.linhas.forEach((row) => delete row.valores[columnId]);
      renderFlexibleTableEditor(block, current);
    });
  });
  bindCustomTableColumnResize(table);
}

function getCustomTableColumnState(table) {
  const inputs = Array.from(table.querySelectorAll('[data-custom-column-name]'));
  const columns = Array.from(table.querySelectorAll('col[data-custom-column-width]'));
  return inputs.map((input, index) => ({
    input,
    column: columns[index],
    width: Number(input.dataset.columnWidth) || 0
  }));
}

function applyCustomTableColumnLayout(table) {
  const states = getCustomTableColumnState(table);
  const totalWidth = states.reduce((sum, state) => sum + state.width, 0);
  if (!totalWidth) return;
  states.forEach((state) => {
    state.column.style.width = `${(state.width / totalWidth) * 100}%`;
  });
  table.style.width = '100%';
}

function applyCustomTableColumnWidths(table, widths) {
  const states = getCustomTableColumnState(table);
  if (states.length !== widths.length) return;
  const targetTotal = states.reduce((sum, state) => sum + state.width, 0);
  let allocated = 0;
  states.forEach((state, index) => {
    const width = index === states.length - 1
      ? targetTotal - allocated
      : Number(widths[index].toFixed(CUSTOM_TABLE_WIDTH_PRECISION));
    const normalizedWidth = Number(width.toFixed(CUSTOM_TABLE_WIDTH_PRECISION));
    state.input.dataset.columnWidth = String(normalizedWidth);
    allocated += normalizedWidth;
  });
  applyCustomTableColumnLayout(table);
}

function adjustCustomTableColumns(table, columnIndex, deltaWidth, sourceWidths = null) {
  const states = getCustomTableColumnState(table);
  if (!states[columnIndex] || columnIndex >= states.length - 1) return;
  const headerCells = Array.from(table.querySelectorAll('th[data-custom-column-id]'));
  const totalDataWidth = headerCells.reduce((sum, cell) => sum + cell.getBoundingClientRect().width, 0);
  if (!totalDataWidth) return;

  const widths = sourceWidths ? [...sourceWidths] : states.map((state) => state.width);
  const totalWidth = widths.reduce((sum, width) => sum + width, 0);
  const pixelsPerWidthUnit = totalDataWidth / totalWidth;
  const minimumWidth = Math.min(
    totalWidth / states.length,
    Math.max(
      1,
      (CUSTOM_TABLE_MIN_COLUMN_WIDTH + 32) / pixelsPerWidthUnit
    )
  );
  const followingIndexes = widths
    .map((_, index) => index)
    .slice(columnIndex + 1);

  if (deltaWidth > 0) {
    const availableByColumn = followingIndexes.map((index) => Math.max(0, widths[index] - minimumWidth));
    const totalAvailable = availableByColumn.reduce((sum, width) => sum + width, 0);
    const appliedDelta = Math.min(deltaWidth, totalAvailable);
    if (appliedDelta <= 0) return;
    widths[columnIndex] += appliedDelta;
    followingIndexes.forEach((index, offset) => {
      widths[index] -= appliedDelta * (availableByColumn[offset] / totalAvailable);
    });
  } else if (deltaWidth < 0) {
    const availableCurrentWidth = Math.max(0, widths[columnIndex] - minimumWidth);
    const releasedWidth = Math.min(-deltaWidth, availableCurrentWidth);
    if (releasedWidth <= 0) return;
    const followingTotal = followingIndexes.reduce((sum, index) => sum + widths[index], 0);
    widths[columnIndex] -= releasedWidth;
    followingIndexes.forEach((index) => {
      widths[index] += releasedWidth * (widths[index] / followingTotal);
    });
  } else {
    return;
  }

  const adjustedTotal = widths.reduce((sum, width) => sum + width, 0);
  widths[widths.length - 1] += totalWidth - adjustedTotal;
  applyCustomTableColumnWidths(table, widths);
}

function bindCustomTableColumnResize(table) {
  table.querySelectorAll('[data-custom-column-resizer]').forEach((handle, columnIndex) => {
    handle.addEventListener('pointerdown', (event) => {
      if (event.button !== 0) return;
      event.preventDefault();
      const states = getCustomTableColumnState(table);
      const state = states[columnIndex];
      const headerCells = Array.from(table.querySelectorAll('th[data-custom-column-id]'));
      const totalDataWidth = headerCells.reduce((sum, cell) => sum + cell.getBoundingClientRect().width, 0);
      const totalWidth = states.reduce((sum, column) => sum + column.width, 0);
      if (!state || !totalDataWidth || !totalWidth) return;

      const startX = event.clientX;
      const startWidths = states.map((column) => column.width);
      const pixelsPerWidthUnit = totalDataWidth / totalWidth;
      const pointerId = event.pointerId;
      handle.classList.add('is-resizing');
      document.body.classList.add('resizing-table-column');
      handle.setPointerCapture?.(pointerId);

      const onPointerMove = (moveEvent) => {
        if (moveEvent.pointerId !== pointerId) return;
        const deltaWidth = (moveEvent.clientX - startX) / pixelsPerWidthUnit;
        adjustCustomTableColumns(table, columnIndex, deltaWidth, startWidths);
      };
      const stopResize = (endEvent) => {
        if (endEvent.pointerId !== pointerId) return;
        handle.classList.remove('is-resizing');
        document.body.classList.remove('resizing-table-column');
        handle.removeEventListener('pointermove', onPointerMove);
        handle.removeEventListener('pointerup', stopResize);
        handle.removeEventListener('pointercancel', stopResize);
        handle.releasePointerCapture?.(pointerId);
      };

      handle.addEventListener('pointermove', onPointerMove);
      handle.addEventListener('pointerup', stopResize);
      handle.addEventListener('pointercancel', stopResize);
    });

    handle.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
      event.preventDefault();
      adjustCustomTableColumns(table, columnIndex, event.key === 'ArrowLeft' ? -2 : 2);
    });
  });
}

function readFlexibleTable(block, options = {}) {
  const columns = Array.from(block.querySelectorAll('[data-custom-column-name]')).map((input) => ({
    id: input.dataset.columnId,
    nome: input.value.trim(),
    largura: Number(input.dataset.columnWidth)
  }));
  const rows = Array.from(block.querySelectorAll('.custom-table tbody tr')).map((row) => {
    const values = {};
    row.querySelectorAll('[data-custom-cell]').forEach((input) => {
      values[input.dataset.columnId] = input.value.trim();
    });
    return { valores: values };
  });
  return {
    colunas: normalizeFlexibleTableColumns(columns),
    linhas: options.keepEmptyRows ? rows : rows.filter((row) => Object.values(row.valores).some(Boolean))
  };
}

function collectFlexibleBlock(block) {
  const type = block.dataset.flexType;
  const result = {
    id: block.dataset.flexId,
    tipo: type,
    titulo: block.querySelector('[data-flex-title]')?.value.trim() || ''
  };
  if (type === 'texto') {
    result.observacoes = collectTopicObservations(block.querySelector('[data-topic-observations]'));
    result.subtopicos = collectFlexibleSubtopics(block.querySelector('.flexible-block-body > [data-flex-subtopics]'));
  } else if (type === 'lista') {
    result.linhas = Array.from(block.querySelectorAll('[data-flex-list] tr'))
      .map((row) => ({
        descricao: row.querySelector('[data-flex-list-field="descricao"]').value.trim(),
        numero_documento: row.querySelector('[data-flex-list-field="numero_documento"]').value.trim(),
        data: row.querySelector('[data-flex-list-field="data"]').value.trim()
      }))
      .filter((row) => row.descricao || row.numero_documento || row.data);
  } else if (type === 'tabela') {
    Object.assign(result, readFlexibleTable(block));
  } else if (type === 'preco') {
    const topicsContainer = block.querySelector('[data-flex-price-topics]');
    result.topicos_preco = collectPriceTopics(topicsContainer);
    ['preco_total_numero', 'preco_total_extenso', 'moeda', 'validade_proposta', 'pagamento', 'prazo_entrega', 'frete', 'impostos'].forEach((field) => {
      const input = block.querySelector(`[data-flex-price-field="${field}"]`);
      result[field] = field === 'preco_total_numero' ? readNumber(input?.value) : (input?.value.trim() || '');
    });
  }
  return result;
}

function collectFlexibleBlocks() {
  return Array.from(flexibleBlocks.querySelectorAll('.flexible-block'))
    .map(collectFlexibleBlock)
    .filter((block) => {
      if (block.tipo === 'quebra_pagina') return true;
      if (block.titulo) return true;
      if (block.tipo === 'texto') return block.observacoes.length > 0 || block.subtopicos.length > 0;
      if (block.tipo === 'lista') return block.linhas.length > 0;
      if (block.tipo === 'preco') return true;
      return block.linhas.length > 0;
    });
}

function collectTopicObservations(container) {
  if (!container) return [];
  return Array.from(container.querySelectorAll(':scope > .topic-observation-row [data-topic-observation]'))
    .map((input) => input.value.trim())
    .filter(Boolean);
}

function collectFlexibleSubtopics(container) {
  if (!container) return [];
  return Array.from(container.children)
    .filter((element) => element.classList.contains('flex-subtopic'))
    .map((subtopic) => {
      const nestedContainer = Array.from(subtopic.children)
        .find((element) => element.hasAttribute?.('data-flex-subtopics'));
      return {
        id: subtopic.dataset.subtopicId,
        titulo: subtopic.querySelector(':scope > .flex-subtopic-heading [data-flex-subtopic-title]').value.trim(),
        observacoes: collectTopicObservations(subtopic.querySelector(':scope > [data-subtopic-observations]')),
        subtopicos: collectFlexibleSubtopics(nestedContainer)
      };
    })
    .filter((subtopic) => subtopic.titulo || subtopic.observacoes.length || subtopic.subtopicos.length);
}

function updateFlexibleEmptyState() {
  flexibleEmptyState.hidden = flexibleBlocks.querySelector('.flexible-block') !== null;
}

function renumberFlexibleTopics() {
  let sectionNumber = 1;
  Array.from(flexibleBlocks.children)
    .filter((section) => section.classList.contains('proposal-section') && !section.hidden)
    .forEach((section) => {
    if (section.dataset.flexType === 'quebra_pagina') {
      delete section.dataset.sectionNumber;
      return;
    }
    section.dataset.sectionNumber = String(sectionNumber);
    const fixedNumberTarget = section.querySelector('[data-section-number]');
    if (fixedNumberTarget) fixedNumberTarget.textContent = `${sectionNumber}.`;
    const numberTarget = section.querySelector('[data-flex-section-number]');
    if (numberTarget) numberTarget.textContent = `${sectionNumber}.`;
    section.querySelectorAll('[data-scope-subsection]').forEach((target) => {
      target.textContent = `${sectionNumber}.${target.dataset.scopeSubsection}`;
    });
    renumberNestedSubtopics(section.querySelector('.flexible-block-body > [data-flex-subtopics]'), String(sectionNumber));
    sectionNumber += 1;
    });
}

function renumberNestedSubtopics(container, parentNumber) {
  if (!container) return;
  Array.from(container.children)
    .filter((element) => element.classList.contains('flex-subtopic'))
    .forEach((subtopic, index) => {
      const number = `${parentNumber}.${index + 1}`;
      const numberTarget = subtopic.querySelector(':scope > .flex-subtopic-heading [data-flex-subtopic-number]');
      if (numberTarget) numberTarget.textContent = number;
      const nestedContainer = Array.from(subtopic.children)
        .find((element) => element.hasAttribute?.('data-flex-subtopics'));
      renumberNestedSubtopics(nestedContainer, number);
    });
}

function clearFlexibleBlocks() {
  flexibleBlocks.querySelectorAll(':scope > .flexible-block').forEach((block) => block.remove());
  updateFlexibleEmptyState();
}

function restoreDefaultFixedSectionOrder() {
  DEFAULT_FIXED_SECTION_ORDER.forEach((sectionId) => {
    const section = flexibleBlocks.querySelector(`:scope > [data-proposal-section="${sectionId}"]`);
    if (section) flexibleBlocks.appendChild(section);
  });
  renumberFlexibleTopics();
}

function collectProposalSectionOrder() {
  return Array.from(flexibleBlocks.children)
    .filter((section) => section.classList.contains('proposal-section') && !section.hidden)
    .map((section) => section.dataset.proposalSection)
    .filter(Boolean);
}

function applyProposalSectionOrder(order) {
  if (!Array.isArray(order) || !order.length) {
    renumberFlexibleTopics();
    return;
  }
  const sections = new Map(Array.from(flexibleBlocks.children)
    .filter((section) => section.classList.contains('proposal-section'))
    .map((section) => [section.dataset.proposalSection, section]));
  order.forEach((sectionId) => {
    const section = sections.get(sectionId);
    if (!section) return;
    flexibleBlocks.appendChild(section);
    sections.delete(sectionId);
  });
  sections.forEach((section) => flexibleBlocks.appendChild(section));
  renumberFlexibleTopics();
}

function stripAutomaticNumber(value) {
  return String(value || '').replace(/^\s*(?:\d+\.\s+|\d+(?:\.\d+)+\s+)/, '').trim();
}

function createFlexibleId(prefix) {
  const random = globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  return `${prefix}-${random}`;
}

function addPriceTopic(topic = {}, withDefaultItem = false, targetContainer) {
  if (!targetContainer) return null;
  const defaults = DEFAULT_PRICE_TOPICS.find((item) => item.tipo === topic.tipo) || {};
  const title = topic.titulo || defaults.titulo || 'Novo t\u00f3pico';
  const type = topic.tipo || 'personalizado';
  const section = document.createElement('div');
  section.className = 'price-topic';
  section.dataset.topicType = type;
  section.dataset.defaultNcm = topic.ncm || defaults.ncm || '-----';
  section.dataset.defaultUn = topic.un ?? defaults.un ?? 'PC';
  section.innerHTML = `
    <div class="topic-heading">
      <input class="topic-title-input" data-topic-title value="${escapeHtml(title)}" placeholder="Nome do t\u00f3pico">
      <div class="topic-actions">
        <button class="small-action" type="button" data-add-topic-item>Adicionar linha</button>
        <button class="danger-action" type="button" data-remove-topic title="Remover t\u00f3pico">x</button>
      </div>
    </div>
    <div class="item-table"></div>
  `;

  section.querySelector('[data-add-topic-item]').addEventListener('click', () => addItem(section));
  section.querySelector('[data-remove-topic]').addEventListener('click', () => {
    section.remove();
    renumberItemCodes(targetContainer);
    recalculate();
  });

  targetContainer.appendChild(section);
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
  const topicsContainer = topicElement.closest('.price-topics');
  if (!container.querySelector('.item-head')) {
    container.appendChild(createItemHeader());
  }

  const row = document.createElement('div');
  row.className = 'item-row';
  row.dataset.itemType = topicElement.dataset.topicType || 'personalizado';
  row.innerHTML = `
    <input data-field="item" value="${escapeHtml(values.item || nextItemCode(topicsContainer))}">
    <input data-field="descricao" value="${escapeHtml(values.descricao || '')}">
    <input data-field="ncm" value="${escapeHtml(values.ncm || topicElement.dataset.defaultNcm || '-----')}">
    <input data-field="quant" type="text" inputmode="decimal" value="${escapeHtml(values.quant ?? 1)}">
    <input data-field="un" value="${escapeHtml(values.un || topicElement.dataset.defaultUn || 'PC')}">
    <input data-field="valor_unit" type="text" inputmode="decimal" value="${escapeHtml(values.valor_unit ?? 0)}">
    <input data-field="valor_total" readonly value="0,00">
    <button class="danger-action" type="button" title="Remover">x</button>
  `;
  row.querySelector('button').addEventListener('click', () => {
    row.remove();
    renumberItemCodes(topicsContainer);
    recalculate();
  });
  row.querySelector('[data-field="ncm"]').addEventListener('blur', (event) => {
    if (!event.currentTarget.value.trim()) event.currentTarget.value = '-----';
  });
  row.querySelectorAll('input:not([readonly])').forEach(bindSelectAllOnEntry);
  container.appendChild(row);
  renumberItemCodes(topicsContainer);
  recalculate();
}

function applyImportedData(result) {
  Object.entries(result.data || {}).forEach(([field, value]) => {
    if (form.elements[field] && value != null && !Array.isArray(value) && typeof value !== 'object') {
      form.elements[field].value = String(value);
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

  clearFlexibleBlocks();
  const migratedPrice = migrateLegacyPriceData(importedData, topics);
  migratedPrice.blocks.forEach((block) => addFlexibleBlock(block));
  applyProposalSectionOrder(migratedPrice.order);
  updateFlexibleEmptyState();

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

function migrateLegacyPriceData(data = {}, topics = []) {
  const blocks = Array.isArray(data.blocos_adicionais)
    ? data.blocos_adicionais.map((block) => ({ ...block }))
    : [];
  const originalOrder = Array.isArray(data.ordem_secoes) ? [...data.ordem_secoes] : [];
  if (blocks.some((block) => block.tipo === 'preco') || !Array.isArray(topics) || !topics.length) {
    return {
      blocks,
      order: originalOrder.filter((sectionId) => sectionId !== 'preco')
    };
  }

  const id = createFlexibleId('preco-migrado');
  blocks.push({
    id,
    tipo: 'preco',
    topicos_preco: topics,
    preco_total_numero: data.preco_total_numero || 0,
    preco_total_extenso: data.preco_total_extenso || '',
    moeda: data.moeda || 'Real R$',
    validade_proposta: data.validade_proposta || '',
    pagamento: data.pagamento || '',
    prazo_entrega: data.prazo_entrega || '',
    frete: data.frete || '',
    impostos: data.impostos || ''
  });
  const migratedSectionId = `flex:${id}`;
  const order = originalOrder.length
    ? originalOrder.map((sectionId) => sectionId === 'preco' ? migratedSectionId : sectionId)
    : [];
  if (order.length && !order.includes(migratedSectionId)) order.push(migratedSectionId);
  return { blocks, order };
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

function nextItemCode(container) {
  if (!container) return '0001';
  const numbers = Array.from(container.querySelectorAll('.item-row [data-field="item"]'))
    .map((input) => Number.parseInt(input.value, 10))
    .filter(Number.isFinite);
  const next = numbers.length ? Math.max(...numbers) + 1 : 1;
  return String(next).padStart(4, '0');
}

function renumberItemCodes(container) {
  if (!container) return;
  const rows = Array.from(container.querySelectorAll('.item-row'));
  rows.forEach((row, index) => {
    row.querySelector('[data-field="item"]').value = String(index + 1).padStart(4, '0');
  });
}

function bindSelectAllOnEntry(input) {
  let selectOnPointerUp = false;
  input.addEventListener('pointerdown', () => {
    selectOnPointerUp = document.activeElement !== input;
  });
  input.addEventListener('focus', () => input.select());
  input.addEventListener('pointerup', (event) => {
    if (!selectOnPointerUp) return;
    event.preventDefault();
    input.select();
    selectOnPointerUp = false;
  });
}

function recalculate() {
  document.querySelectorAll('[data-flex-price-editor]').forEach((editor) => {
    const editorTotal = recalculatePriceRows(editor.querySelector('[data-flex-price-topics]'));
    editor.closest('.flexible-block').querySelector('[data-flex-grand-total]').textContent = moneyFormatter.format(editorTotal);
    editor.querySelector('[data-flex-price-field="preco_total_numero"]').value = moneyFormatter.format(editorTotal);
    editor.querySelector('[data-flex-price-field="preco_total_extenso"]').value = moneyToWords(editorTotal);
  });
}

function recalculatePriceRows(container) {
  let total = 0;
  if (!container) return total;
  container.querySelectorAll('.item-row').forEach((row) => {
    const quant = readNumber(row.querySelector('[data-field="quant"]').value);
    const unit = readNumber(row.querySelector('[data-field="valor_unit"]').value);
    const rowTotal = quant * unit;
    row.querySelector('[data-field="valor_total"]').value = numberFormatter.format(rowTotal);
    total += rowTotal;
  });
  return total;
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
    showResult('Proposta pronta', result.outputPath);
    await loadHistory();
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
    ['responsavel_nome', 'Informe o responsável.']
  ];

  requiredFields.forEach(([name, message]) => {
    const field = form.elements[name];
    if (!field || field.disabled || String(field.value || '').trim()) return;
    field.classList.add('invalid');
    messages.push(message);
    if (!focusTarget) focusTarget = field;
  });

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

  data.objeto_observacoes = [];
  data.objeto = '';
  data.secoes_excluidas = ['objeto', 'escopo'];
  data.blocos_adicionais = collectFlexibleBlocks();
  data.ordem_secoes = collectProposalSectionOrder();
  data.topicos_preco = [];
  data.itens_servico = [];
  data.itens_consumiveis = [];
  data.preco_total_numero = data.blocos_adicionais
    .filter((block) => block.tipo === 'preco')
    .reduce((sum, block) => sum + Number(block.preco_total_numero || 0), 0);
  return data;
}

function collectPriceTopics(container) {
  if (!container) return [];
  return Array.from(container.querySelectorAll('.price-topic')).map((topicElement) => {
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
    ncm: readRowField(row, 'ncm') || '-----',
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

async function loadModels() {
  if (!window.supplyMarine?.listarModelos) return;
  state.models = await window.supplyMarine.listarModelos();
  renderModels();
}

function renderModels() {
  modelEmptyState.hidden = state.models.length > 0;
  modelList.innerHTML = state.models.map((model) => {
    const blocks = model.estrutura?.blocos_adicionais?.length || 0;
    const updated = formatModelDate(model.updatedAt || model.createdAt);
    return `
      <article class="model-card">
        <div>
          <h3>${escapeHtml(model.nome || 'Modelo sem nome')}</h3>
          <p>${escapeHtml(model.empresa || '-')}</p>
        </div>
        <div class="model-card-meta">
          <span>${blocks} ${blocks === 1 ? 'bloco adicional' : 'blocos adicionais'}</span>
          <span>•</span>
          <span>${escapeHtml(updated)}</span>
        </div>
        <div class="model-card-actions">
          <button class="primary-action" type="button" data-model-action="use" data-id="${escapeHtml(model.id)}">Usar modelo</button>
          <button class="danger-action" type="button" data-model-action="delete" data-id="${escapeHtml(model.id)}" title="Excluir modelo" aria-label="Excluir modelo">x</button>
        </div>
      </article>
    `;
  }).join('');
  modelList.querySelectorAll('[data-model-action]').forEach((button) => {
    button.addEventListener('click', () => handleModelAction(button.dataset.modelAction, button.dataset.id));
  });
}

function formatModelDate(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? '' : date.toLocaleDateString('pt-BR');
}

async function handleModelAction(action, id) {
  const model = state.models.find((item) => item.id === id);
  if (!model) return;
  if (action === 'use') {
    applyProposalModel(model);
    return;
  }
  if (action === 'delete') {
    const confirmed = window.confirm(`Excluir o modelo ${model.nome || ''}?`);
    if (!confirmed) return;
    await window.supplyMarine.excluirModelo(id);
    await loadModels();
  }
}

function applyProposalModel(model) {
  loadProposalIntoForm({ id: '', data: model.estrutura || {}, docxPath: '' }, { duplicate: true });
  document.querySelector('#form-title').textContent = `Nova proposta · ${model.nome || 'Modelo'}`;
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
  clearFlexibleBlocks();
  restoreDefaultFixedSectionOrder();
  resultPanel.hidden = true;
  state.lastOutputPath = options.duplicate ? '' : proposal.docxPath;

  fillFormFields({
    ...data,
    numero_documento: options.duplicate && (data.numero_documento || proposal.numero_documento)
      ? `${data.numero_documento || proposal.numero_documento}-COPIA`
      : data.numero_documento
  });
  const topics = Array.isArray(data.topicos_preco) && data.topicos_preco.length
    ? data.topicos_preco
    : legacyTopicsFromData(data);
  const migratedPrice = migrateLegacyPriceData(data, topics);
  migratedPrice.blocks.forEach((block) => addFlexibleBlock(block));
  applyProposalSectionOrder(migratedPrice.order);
  updateFlexibleEmptyState();

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
