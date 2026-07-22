const fs = require('node:fs');
const path = require('node:path');
const { getIndexPath, getPropostasDir } = require('./paths');

function indexPath() {
  return getIndexPath();
}

function ensureStorage() {
  const propostasDir = getPropostasDir();
  fs.mkdirSync(propostasDir, { recursive: true });
  if (!fs.existsSync(indexPath())) {
    fs.writeFileSync(indexPath(), JSON.stringify({ propostas: [], modelos: [] }, null, 2));
  }
}

function readIndex() {
  ensureStorage();
  try {
    const parsed = JSON.parse(fs.readFileSync(indexPath(), 'utf8'));
    return Array.isArray(parsed.propostas)
      ? { ...parsed, modelos: Array.isArray(parsed.modelos) ? parsed.modelos : [] }
      : { propostas: [], modelos: [] };
  } catch {
    return { propostas: [], modelos: [] };
  }
}

function writeIndex(index) {
  ensureStorage();
  fs.writeFileSync(indexPath(), JSON.stringify(index, null, 2));
}

function listarPropostas() {
  const index = readIndex();
  return index.propostas
    .slice()
    .sort((a, b) => String(b.updatedAt || b.createdAt).localeCompare(String(a.updatedAt || a.createdAt)));
}

function salvarProposta({ id, data, docxPath, pdfPath }) {
  const index = readIndex();
  const now = new Date().toISOString();
  const existingIndex = id ? index.propostas.findIndex((item) => item.id === id) : -1;
  const previous = existingIndex >= 0 ? index.propostas[existingIndex] : {};
  const proposta = {
    id: previous.id || createId(),
    numero_documento: data.numero_documento || '',
    empresa_cliente: data.empresa_cliente || '',
    unidade: data.unidade || '',
    data_documento: data.data_documento || '',
    preco_total_numero: Number(data.preco_total_numero || 0),
    docxPath: docxPath || previous.docxPath || '',
    pdfPath: pdfPath || previous.pdfPath || '',
    data,
    createdAt: previous.createdAt || now,
    updatedAt: now
  };

  if (existingIndex >= 0) {
    index.propostas[existingIndex] = proposta;
  } else {
    index.propostas.push(proposta);
  }

  writeIndex(index);
  return proposta;
}

function atualizarPdf(docxPath, pdfPath) {
  const index = readIndex();
  const proposta = index.propostas.find((item) => item.docxPath === docxPath);
  if (!proposta) {
    return null;
  }

  proposta.pdfPath = pdfPath;
  proposta.updatedAt = new Date().toISOString();
  writeIndex(index);
  return proposta;
}

function excluirProposta(id) {
  const index = readIndex();
  const proposta = index.propostas.find((item) => item.id === id);
  index.propostas = index.propostas.filter((item) => item.id !== id);
  writeIndex(index);

  if (proposta) {
    removeIfInsideStorage(proposta.docxPath);
    removeIfInsideStorage(proposta.pdfPath);
  }

  return { removed: Boolean(proposta) };
}

function listarModelos() {
  const index = readIndex();
  return index.modelos
    .slice()
    .sort((a, b) => String(b.updatedAt || b.createdAt).localeCompare(String(a.updatedAt || a.createdAt)));
}

function salvarModelo({ id, nome, empresa, estrutura }) {
  const index = readIndex();
  const now = new Date().toISOString();
  const existingIndex = id ? index.modelos.findIndex((item) => item.id === id) : -1;
  const previous = existingIndex >= 0 ? index.modelos[existingIndex] : {};
  const modelo = {
    id: previous.id || createId(),
    nome: String(nome || '').trim(),
    empresa: String(empresa || '').trim(),
    estrutura,
    createdAt: previous.createdAt || now,
    updatedAt: now
  };

  if (existingIndex >= 0) {
    index.modelos[existingIndex] = modelo;
  } else {
    index.modelos.push(modelo);
  }
  writeIndex(index);
  return modelo;
}

function excluirModelo(id) {
  const index = readIndex();
  const exists = index.modelos.some((item) => item.id === id);
  index.modelos = index.modelos.filter((item) => item.id !== id);
  writeIndex(index);
  return { removed: exists };
}

function removeIfInsideStorage(filePath) {
  if (!filePath) return;

  const resolved = path.resolve(filePath);
  const storageRoot = path.resolve(getPropostasDir());
  if (!resolved.startsWith(storageRoot)) return;

  try {
    if (fs.existsSync(resolved)) {
      fs.unlinkSync(resolved);
    }
  } catch {
    // Arquivo aberto/bloqueado: remove do indice e deixa o arquivo no disco.
  }
}

function createId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

module.exports = {
  listarPropostas,
  salvarProposta,
  atualizarPdf,
  excluirProposta,
  listarModelos,
  salvarModelo,
  excluirModelo,
  indexPath
};
