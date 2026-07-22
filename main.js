const path = require('node:path');
const fs = require('node:fs');
const { app, BrowserWindow, ipcMain, shell, dialog, safeStorage } = require('electron');
const { autoUpdater } = require('electron-updater');
const { gerarDocx } = require('./src/lib/gerarDocx');
const { gerarPdf } = require('./src/lib/gerarPdf');
const {
  listarPropostas,
  salvarProposta,
  atualizarPdf,
  excluirProposta,
  listarModelos,
  salvarModelo,
  excluirModelo
} = require('./src/lib/armazenamento');
const { normalizarEstruturaModelo } = require('./src/lib/modeloProposta');
const { getPropostasDir } = require('./src/lib/paths');
const { importarDocx } = require('./src/lib/importarDocx');
const { createOpenAiConfig } = require('./src/lib/openAiConfig');
const {
  isCloudEnabled,
  listarPropostasCloud,
  salvarPropostaCloud,
  atualizarPdfCloud,
  excluirPropostaCloud,
  listarModelosCloud,
  salvarModeloCloud,
  excluirModeloCloud,
  resolveCloudFile
} = require('./src/lib/supabaseStorage');

const isSmoke = process.argv.includes('--smoke');
const isSmokeGenerate = process.argv.includes('--smoke-generate');
const isSmokePdf = process.argv.includes('--smoke-pdf');
const isSmokeInvalid = process.argv.includes('--smoke-invalid');
const isSmokeModels = process.argv.includes('--smoke-models');
const isAnySmoke = isSmoke || isSmokeGenerate || isSmokePdf || isSmokeInvalid || isSmokeModels;
const isPackagedRuntime = app.isPackaged && !isAnySmoke;

if (isAnySmoke) {
  app.setPath('userData', path.join(app.getPath('temp'), `supply-marine-propostas-smoke-${process.pid}`));
}

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 1180,
    height: 760,
    minWidth: 960,
    minHeight: 640,
    icon: path.join(__dirname, 'assets', 'logo.ico'),
    backgroundColor: '#f4f6f8',
    title: 'Supply Marine - Propostas',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, 'src', 'renderer', 'index.html'));

  if (isAnySmoke) {
    mainWindow.webContents.once('did-finish-load', async () => {
      if (isSmokeModels) {
        try {
          const result = await mainWindow.webContents.executeJavaScript(`
            new Promise((resolve, reject) => {
              document.querySelector('[data-view-button="form"]').click();
              document.querySelector('[name="empresa_cliente"]').value = 'Conteudo que nao deve voltar';
              document.querySelector('[data-add-flex-block="tabela"]').click();
              const table = document.querySelector('#proposal-sections .flexible-block:last-child');
              table.querySelector('[data-flex-title]').value = 'EQUIPAMENTOS';
              table.querySelector('[data-custom-column-name]').value = 'Equipamento';
              table.querySelector('[data-custom-cell]').value = 'Bomba preenchida';
              document.querySelector('#save-proposal-model').click();
              document.querySelector('#model-name').value = 'Modelo smoke';
              document.querySelector('#model-company').value = 'Empresa smoke';
              document.querySelector('#model-dialog-form').requestSubmit();
              const start = Date.now();
              const timer = setInterval(() => {
                const useButton = document.querySelector('[data-model-action="use"]');
                if (useButton) {
                  clearInterval(timer);
                  useButton.click();
                  const appliedTable = document.querySelector('#proposal-sections .flexible-block[data-flex-type="tabela"]');
                  const title = appliedTable?.querySelector('[data-flex-title]')?.value || '';
                  const column = appliedTable?.querySelector('[data-custom-column-name]')?.value || '';
                  const item = appliedTable?.querySelector('[data-custom-cell]')?.value || '';
                  const company = document.querySelector('[name="empresa_cliente"]').value;
                  if (title === 'EQUIPAMENTOS' && column === 'Equipamento' && !item && !company) {
                    document.querySelector('#preview-proposal').click();
                    const preview = document.querySelector('#preview-dialog');
                    const previewText = document.querySelector('#proposal-preview-content')?.textContent || '';
                    const generatedPanelHidden = document.querySelector('#result-panel').hidden;
                    if (preview.open && previewText.includes('EQUIPAMENTOS') && generatedPanelHidden) {
                      resolve('estrutura mantida, itens limpos e previa sem gerar proposta');
                    } else {
                      reject(new Error(JSON.stringify({ previewOpen: preview.open, previewText, generatedPanelHidden })));
                    }
                  } else {
                    reject(new Error(JSON.stringify({ title, column, item, company })));
                  }
                }
                if (Date.now() - start > 5000) {
                  clearInterval(timer);
                  reject(new Error('Modelo nao apareceu na biblioteca'));
                }
              }, 100);
            });
          `);
          console.log(`Smoke modelos: ${result}`);
          app.quit();
        } catch (error) {
          console.error(error);
          app.exit(1);
        }
        return;
      }
      if (isSmokeInvalid) {
        try {
          const validationText = await mainWindow.webContents.executeJavaScript(`
            new Promise((resolve, reject) => {
              document.querySelector('[data-view-button="form"]').click();
              document.querySelector('[name="empresa_cliente"]').value = '';
              document.querySelector('#proposal-form').requestSubmit();
              const start = Date.now();
              const timer = setInterval(() => {
                const text = document.querySelector('#form-alert')?.textContent || '';
                if (text.includes('Informe a empresa cliente')) {
                  clearInterval(timer);
                  resolve(text);
                }
                if (Date.now() - start > 5000) {
                  clearInterval(timer);
                  reject(new Error(text || 'Validação não apareceu'));
                }
              }, 100);
            });
          `);
          console.log(`Smoke validação: ${validationText}`);
          app.quit();
        } catch (error) {
          console.error(error);
          app.exit(1);
        }
        return;
      }

      if (isSmokeGenerate || isSmokePdf) {
        try {
          const outputPath = await mainWindow.webContents.executeJavaScript(`
            new Promise((resolve, reject) => {
              document.querySelector('[data-view-button="form"]').click();
              document.querySelector('[name="empresa_cliente"]').value = 'Cliente Smoke';
              document.querySelector('[name="numero_documento"]').value = 'SMOKE-001';
              document.querySelector('[name="responsavel_nome"]').value = 'Responsavel Smoke';
              document.querySelector('[name="objeto"]').value = 'Teste de geração automática';
              document.querySelector('[data-service-description]').value = 'Serviço de teste';
              document.querySelector('[data-add-flex-block="texto"]').click();
              const topic = document.querySelector('#proposal-sections .flexible-block:last-child');
              topic.querySelector('[data-flex-title]').value = 'CRONOGRAMA';
              topic.querySelector('[data-add-topic-observation]').click();
              topic.querySelector('[data-topic-observation]').value = 'Conteudo adicional do smoke test.';
              topic.querySelector('[data-add-subtopic]').click();
              topic.querySelector('[data-flex-subtopic-title]').value = 'Etapas';
              topic.querySelector('[data-add-subtopic-observation]').click();
              topic.querySelector('[data-subtopic-observations] [data-topic-observation]').value = 'Primeira etapa';
              const firstSubtopic = topic.querySelector('[data-flex-subtopics] > .flex-subtopic');
              firstSubtopic.querySelector('[data-add-nested-subtopic]').click();
              const nestedSubtopic = firstSubtopic.querySelector('[data-flex-subtopics] > .flex-subtopic');
              nestedSubtopic.querySelector('[data-flex-subtopic-title]').value = 'Detalhamento';
              nestedSubtopic.querySelector('[data-add-subtopic-observation]').click();
              nestedSubtopic.querySelector('[data-subtopic-observations] [data-topic-observation]').value = 'Etapa interna';
              document.querySelector('[data-add-flex-block="lista"]').click();
              const documentList = document.querySelector('#proposal-sections .flexible-block:last-child');
              documentList.querySelector('[data-flex-list-field="descricao"]').value = 'Solicitacao de cotacao';
              documentList.querySelector('[data-flex-list-field="numero_documento"]').value = 'SC35085';
              documentList.querySelector('[data-flex-list-field="data"]').value = '14072026';
              documentList.querySelector('[data-flex-list-field="data"]').dispatchEvent(new Event('input', { bubbles: true }));
              document.querySelector('[data-add-flex-block="preco"]').click();
              const additionalPrice = document.querySelector('#proposal-sections .flexible-block:last-child');
              additionalPrice.querySelector('[data-topic-title]').value = 'LOCACAO';
              additionalPrice.querySelector('[data-field="descricao"]').value = 'Equipamento adicional';
              additionalPrice.querySelector('[data-field="valor_unit"]').value = '250';
              additionalPrice.querySelector('[data-field="quant"]').value = '2';
              additionalPrice.querySelector('[data-flex-price-field="pagamento"]').value = '15 dias';
              additionalPrice.querySelector('[data-field="valor_unit"]').dispatchEvent(new Event('input', { bubbles: true }));
              document.querySelector('[data-add-flex-block="tabela"]').click();
              const customTable = document.querySelector('#proposal-sections .flexible-block:last-child');
              customTable.querySelector('[data-flex-title]').value = 'EQUIPAMENTOS';
              customTable.querySelectorAll('[data-custom-column-name]')[0].value = 'Equipamento';
              customTable.querySelectorAll('[data-custom-column-name]')[1].value = 'Quantidade';
              customTable.querySelectorAll('[data-custom-cell]')[0].value = 'Bomba de teste';
              customTable.querySelectorAll('[data-custom-cell]')[1].value = '2';
              topic.querySelector('[data-flex-action="up"]').click();
              topic.querySelector('[data-flex-action="up"]').click();
              document.querySelector('#proposal-form').requestSubmit();
              const start = Date.now();
              const timer = setInterval(() => {
                const path = document.querySelector('#result-path')?.textContent || '';
                if (path.endsWith('.docx')) {
                  clearInterval(timer);
                  resolve(path);
                }
                if (Date.now() - start > 10000) {
                  clearInterval(timer);
                  const alert = document.querySelector('#form-alert')?.textContent || '';
                  reject(new Error(alert || path || 'Tempo esgotado ao gerar DOCX'));
                }
              }, 200);
            });
          `);
          console.log(`Smoke gerou DOCX: ${outputPath}`);
          if (isSmokePdf) {
            const pdfStatus = await mainWindow.webContents.executeJavaScript(`
              new Promise((resolve, reject) => {
                document.querySelector('#export-pdf').click();
                const start = Date.now();
                const timer = setInterval(() => {
                  const status = document.querySelector('#pdf-status')?.textContent || '';
                  if (status.includes('PDF gerado:') || status.includes('Instale o LibreOffice') || status.includes('Microsoft Word')) {
                    clearInterval(timer);
                    resolve(status);
                  }
                  if (Date.now() - start > 30000) {
                    clearInterval(timer);
                    reject(new Error(status || 'Tempo esgotado ao exportar PDF'));
                  }
                }, 200);
              });
            `);
            console.log(`Smoke PDF: ${pdfStatus}`);
          }
          app.quit();
        } catch (error) {
          console.error(error);
          app.exit(1);
        }
        return;
      }

      setTimeout(() => app.quit(), 300);
    });
  }

  if (isPackagedRuntime) {
    configureAutoUpdate();
  }
}

app.whenReady().then(() => {
  configureStoragePath();
  const openAiConfig = createOpenAiConfig({ app, safeStorage });

  ipcMain.handle('app:version', () => app.getVersion());
  ipcMain.handle('proposta:gerar-docx', async (_event, data) => {
    const outputPath = gerarDocx(data, buildProposalPath(data));
    const payload = {
      id: data._historico_id,
      data: sanitizeProposalData(data),
      docxPath: outputPath
    };
    const localProposta = salvarProposta(payload);
    const proposta = await tryCloud(
      () => salvarPropostaCloud({ ...payload, id: localProposta.id }),
      localProposta
    );
    return { outputPath, proposta };
  });
  ipcMain.handle('proposta:exportar-pdf', async (_event, docxPath) => {
    const resolvedDocxPath = await tryCloud(() => resolveCloudFile(docxPath), docxPath);
    const result = await gerarPdf(resolvedDocxPath);
    if (!result.converted) {
      return {
        ...result,
        openedDocx: false,
        openError: ''
      };
    }
    atualizarPdf(resolvedDocxPath, result.pdfPath);
    await tryCloud(() => atualizarPdfCloud(docxPath, result.pdfPath), null);
    return {
      ...result,
      skipOpen: isSmokePdf
    };
  });
  ipcMain.handle('historico:listar', async () => {
    return tryCloud(listarPropostasCloud, listarPropostas());
  });
  ipcMain.handle('historico:excluir', async (_event, id) => {
    const localResult = excluirProposta(id);
    return tryCloud(() => excluirPropostaCloud(id), localResult);
  });
  ipcMain.handle('modelos:listar', () => tryCloud(listarModelosCloud, listarModelos()));
  ipcMain.handle('modelos:salvar', async (_event, payload = {}) => {
    const nome = String(payload.nome || '').trim();
    const empresa = String(payload.empresa || '').trim();
    if (!nome) throw new Error('Informe o nome do modelo.');
    if (!empresa) throw new Error('Informe a empresa do modelo.');
    const localModel = salvarModelo({
      nome,
      empresa,
      estrutura: normalizarEstruturaModelo(payload.estrutura)
    });
    return tryCloud(() => salvarModeloCloud(localModel), localModel);
  });
  ipcMain.handle('modelos:excluir', async (_event, id) => {
    const localResult = excluirModelo(id);
    return tryCloud(() => excluirModeloCloud(id), localResult);
  });
  ipcMain.handle('arquivo:abrir', async (_event, filePath) => {
    const resolvedPath = await tryCloud(() => resolveCloudFile(filePath), filePath);
    return shell.openPath(resolvedPath);
  });
  ipcMain.handle('docx:selecionar', async () => {
    const result = await dialog.showOpenDialog({
      title: 'Selecionar proposta',
      filters: [{ name: 'Documentos Word', extensions: ['docx'] }],
      properties: ['openFile']
    });
    return result.canceled ? null : result.filePaths[0];
  });
  ipcMain.handle('openai:status', () => openAiConfig.getStatus());
  ipcMain.handle('openai:salvar-chave', (_event, apiKey) => openAiConfig.saveApiKey(apiKey));
  ipcMain.handle('docx:importar', async (_event, filePath) => {
    return importarDocx(filePath, { apiKey: openAiConfig.getApiKey() });
  });
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

function buildProposalPath(data) {
  const parts = [
    data.numero_documento || 'proposta',
    data.empresa_cliente || 'cliente',
    data.data_documento || new Date().toISOString().slice(0, 10)
  ];
  const fileName = `${parts.map(safeFilePart).filter(Boolean).join('_')}.docx`;
  return nextAvailablePath(path.join(getPropostasDir(), fileName));
}

function configureStoragePath() {
  const baseDir = isAnySmoke
    ? app.getPath('userData')
    : app.getPath('documents');
  process.env.SUPPLY_MARINE_PROPOSTAS_DIR = path.join(baseDir, 'Supply Marine Propostas');
}

function configureAutoUpdate() {
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.logger = console;

  autoUpdater.on('update-downloaded', () => {
    dialog.showMessageBox({
      type: 'info',
      buttons: ['Reiniciar agora', 'Depois'],
      defaultId: 0,
      cancelId: 1,
      title: 'Atualizacao pronta',
      message: 'Uma nova versao do Supply Marine Propostas foi baixada.',
      detail: 'Reinicie o aplicativo para instalar a atualizacao.'
    }).then((result) => {
      if (result.response === 0) {
        autoUpdater.quitAndInstall(false, true);
      }
    });
  });

  autoUpdater.on('error', (error) => {
    console.warn(`Atualizacao indisponivel: ${error.message || error}`);
  });

  setTimeout(() => {
    autoUpdater.checkForUpdatesAndNotify().catch((error) => {
      console.warn(`Nao foi possivel verificar atualizacoes: ${error.message || error}`);
    });
  }, 3000);
}

function sanitizeProposalData(data) {
  const clean = { ...data };
  delete clean._historico_id;
  return clean;
}

async function tryCloud(operation, fallback) {
  if (isAnySmoke || !isCloudEnabled()) {
    return fallback;
  }

  try {
    return await operation();
  } catch (error) {
    console.warn(`Supabase indisponivel: ${error.message || error}`);
    return fallback;
  }
}

function nextAvailablePath(filePath) {
  if (!fs.existsSync(filePath)) {
    return filePath;
  }

  const directory = path.dirname(filePath);
  const extension = path.extname(filePath);
  const baseName = path.basename(filePath, extension);

  for (let index = 2; index < 1000; index += 1) {
    const candidate = path.join(directory, `${baseName}_${index}${extension}`);
    if (!fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return path.join(directory, `${baseName}_${Date.now()}${extension}`);
}

function safeFilePart(value) {
  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
