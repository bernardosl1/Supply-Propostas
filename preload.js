const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('supplyMarine', {
  getVersion: () => ipcRenderer.invoke('app:version'),
  gerarDocx: (data) => ipcRenderer.invoke('proposta:gerar-docx', data),
  exportarPdf: (docxPath) => ipcRenderer.invoke('proposta:exportar-pdf', docxPath),
  abrirArquivo: (filePath) => ipcRenderer.invoke('arquivo:abrir', filePath),
  listarPropostas: () => ipcRenderer.invoke('historico:listar'),
  excluirProposta: (id) => ipcRenderer.invoke('historico:excluir', id),
  listarModelos: () => ipcRenderer.invoke('modelos:listar'),
  salvarModelo: (payload) => ipcRenderer.invoke('modelos:salvar', payload),
  excluirModelo: (id) => ipcRenderer.invoke('modelos:excluir', id),
  selecionarDocx: () => ipcRenderer.invoke('docx:selecionar'),
  importarDocx: (filePath) => ipcRenderer.invoke('docx:importar', filePath),
  getOpenAiStatus: () => ipcRenderer.invoke('openai:status'),
  salvarOpenAiKey: (apiKey) => ipcRenderer.invoke('openai:salvar-chave', apiKey)
});
