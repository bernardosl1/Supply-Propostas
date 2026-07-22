const fs = require('node:fs');
const path = require('node:path');
const { createClient } = require('@supabase/supabase-js');
const { getPropostasDir } = require('./paths');

const DEFAULT_BUCKET = 'propostas';
const REMOTE_PREFIX = 'supabase://';

let cachedConfig;
let cachedClient;

function getConfig() {
  if (cachedConfig !== undefined) {
    return cachedConfig;
  }

  const fromEnv = {
    url: process.env.SUPABASE_URL,
    anonKey: process.env.SUPABASE_ANON_KEY,
    bucket: process.env.SUPABASE_BUCKET || DEFAULT_BUCKET
  };

  if (fromEnv.url && fromEnv.anonKey) {
    cachedConfig = fromEnv;
    return cachedConfig;
  }

  const configPath = path.resolve(__dirname, '..', '..', 'config', 'supabase.json');
  if (!fs.existsSync(configPath)) {
    cachedConfig = null;
    return cachedConfig;
  }

  try {
    const parsed = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    cachedConfig = {
      url: parsed.url,
      anonKey: parsed.anonKey || parsed.anon_key || parsed.publishableKey,
      bucket: parsed.bucket || DEFAULT_BUCKET
    };
  } catch {
    cachedConfig = null;
  }

  if (!cachedConfig?.url || !cachedConfig?.anonKey) {
    cachedConfig = null;
  }

  return cachedConfig;
}

function getClient() {
  const config = getConfig();
  if (!config) {
    return null;
  }

  if (!cachedClient) {
    cachedClient = createClient(config.url, config.anonKey, {
      auth: {
        persistSession: false,
        autoRefreshToken: false
      }
    });
  }

  return cachedClient;
}

function isCloudEnabled() {
  return Boolean(getClient());
}

async function listarPropostasCloud() {
  const supabase = requireClient();
  const { data, error } = await supabase
    .from('propostas')
    .select('*')
    .order('updated_at', { ascending: false });

  if (error) throw error;
  return (data || []).map(rowToProposal);
}

async function salvarPropostaCloud({ id, data, docxPath, pdfPath }) {
  const supabase = requireClient();
  const now = new Date().toISOString();
  const proposalId = id || createId();
  const existing = await getProposalRow(proposalId);
  const docxStoragePath = docxPath
    ? await uploadProposalFile(proposalId, docxPath, 'docx')
    : existing?.docx_storage_path || '';
  const pdfStoragePath = pdfPath
    ? await uploadProposalFile(proposalId, pdfPath, 'pdf')
    : existing?.pdf_storage_path || '';

  const row = {
    id: proposalId,
    numero_documento: data.numero_documento || '',
    empresa_cliente: data.empresa_cliente || '',
    data_documento: data.data_documento || '',
    preco_total_numero: Number(data.preco_total_numero || 0),
    data,
    docx_path: docxPath || existing?.docx_path || '',
    pdf_path: pdfPath || existing?.pdf_path || '',
    docx_storage_path: docxStoragePath,
    pdf_storage_path: pdfStoragePath,
    created_at: existing?.created_at || now,
    updated_at: now
  };

  const { data: saved, error } = await supabase
    .from('propostas')
    .upsert(row)
    .select()
    .single();

  if (error) throw error;
  return rowToProposal(saved);
}

async function atualizarPdfCloud(docxPath, pdfPath) {
  const supabase = requireClient();
  const row = await findProposalByDocxPath(docxPath);
  if (!row) {
    return null;
  }

  const pdfStoragePath = await uploadProposalFile(row.id, pdfPath, 'pdf');
  const { data, error } = await supabase
    .from('propostas')
    .update({
      pdf_path: pdfPath,
      pdf_storage_path: pdfStoragePath,
      updated_at: new Date().toISOString()
    })
    .eq('id', row.id)
    .select()
    .single();

  if (error) throw error;
  return rowToProposal(data);
}

async function excluirPropostaCloud(id) {
  const supabase = requireClient();
  const row = await getProposalRow(id);
  const storagePaths = [row?.docx_storage_path, row?.pdf_storage_path].filter(Boolean);

  if (storagePaths.length) {
    await supabase.storage.from(getConfig().bucket).remove(storagePaths);
  }

  const { error } = await supabase
    .from('propostas')
    .delete()
    .eq('id', id);

  if (error) throw error;
  return { removed: Boolean(row) };
}

async function listarModelosCloud() {
  const supabase = requireClient();
  const { data, error } = await supabase
    .from('modelos_propostas')
    .select('*')
    .order('updated_at', { ascending: false });

  if (error) throw error;
  return (data || []).map(rowToModel);
}

async function salvarModeloCloud({ id, nome, empresa, estrutura }) {
  const supabase = requireClient();
  const now = new Date().toISOString();
  const modelId = id || createId();
  const { data: existing, error: existingError } = await supabase
    .from('modelos_propostas')
    .select('created_at')
    .eq('id', modelId)
    .maybeSingle();
  if (existingError) throw existingError;

  const { data, error } = await supabase
    .from('modelos_propostas')
    .upsert({
      id: modelId,
      nome: String(nome || '').trim(),
      empresa: String(empresa || '').trim(),
      estrutura,
      created_at: existing?.created_at || now,
      updated_at: now
    })
    .select()
    .single();

  if (error) throw error;
  return rowToModel(data);
}

async function excluirModeloCloud(id) {
  const supabase = requireClient();
  const { data, error } = await supabase
    .from('modelos_propostas')
    .delete()
    .eq('id', id)
    .select('id');
  if (error) throw error;
  return { removed: Boolean(data?.length) };
}

async function resolveCloudFile(filePath) {
  const remote = parseRemotePath(filePath);
  if (!remote) {
    return filePath;
  }

  const supabase = requireClient();
  const localPath = path.join(getPropostasDir(), 'cache', safeCacheName(remote.path));
  if (fs.existsSync(localPath)) {
    return localPath;
  }

  const { data, error } = await supabase.storage.from(remote.bucket).download(remote.path);
  if (error) throw error;

  const buffer = Buffer.from(await data.arrayBuffer());
  fs.mkdirSync(path.dirname(localPath), { recursive: true });
  fs.writeFileSync(localPath, buffer);
  return localPath;
}

async function uploadProposalFile(proposalId, filePath, extension) {
  const remote = parseRemotePath(filePath);
  if (remote) {
    return remote.path;
  }

  if (!filePath || !fs.existsSync(filePath)) {
    return '';
  }

  const supabase = requireClient();
  const bucket = getConfig().bucket;
  const fileName = path.basename(filePath);
  const storagePath = `${proposalId}/${Date.now()}-${safeStorageName(fileName)}`;
  const contentType = extension === 'pdf'
    ? 'application/pdf'
    : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  const { error } = await supabase.storage
    .from(bucket)
    .upload(storagePath, fs.readFileSync(filePath), {
      contentType,
      upsert: true
    });

  if (error) throw error;
  return storagePath;
}

async function getProposalRow(id) {
  if (!id) return null;
  const supabase = requireClient();
  const { data, error } = await supabase
    .from('propostas')
    .select('*')
    .eq('id', id)
    .maybeSingle();

  if (error) throw error;
  return data;
}

async function findProposalByDocxPath(docxPath) {
  const supabase = requireClient();
  const remote = parseRemotePath(docxPath);
  let query = supabase.from('propostas').select('*');

  if (remote) {
    query = query.eq('docx_storage_path', remote.path);
  } else {
    query = query.eq('docx_path', docxPath);
  }

  const { data, error } = await query.maybeSingle();
  if (error) throw error;
  return data;
}

function rowToProposal(row) {
  const bucket = getConfig()?.bucket || DEFAULT_BUCKET;
  return {
    id: row.id,
    numero_documento: row.numero_documento || '',
    empresa_cliente: row.empresa_cliente || '',
    unidade: row.data?.unidade || '',
    data_documento: row.data_documento || '',
    preco_total_numero: Number(row.preco_total_numero || 0),
    docxPath: row.docx_storage_path ? remotePath(bucket, row.docx_storage_path) : row.docx_path || '',
    pdfPath: row.pdf_storage_path ? remotePath(bucket, row.pdf_storage_path) : row.pdf_path || '',
    data: row.data || {},
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function rowToModel(row) {
  return {
    id: row.id,
    nome: row.nome || '',
    empresa: row.empresa || '',
    estrutura: row.estrutura || {},
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function parseRemotePath(filePath) {
  if (!filePath || !String(filePath).startsWith(REMOTE_PREFIX)) {
    return null;
  }

  const withoutPrefix = String(filePath).slice(REMOTE_PREFIX.length);
  const slashIndex = withoutPrefix.indexOf('/');
  if (slashIndex < 0) {
    return null;
  }

  return {
    bucket: withoutPrefix.slice(0, slashIndex),
    path: withoutPrefix.slice(slashIndex + 1)
  };
}

function remotePath(bucket, storagePath) {
  return `${REMOTE_PREFIX}${bucket}/${storagePath}`;
}

function requireClient() {
  const supabase = getClient();
  if (!supabase) {
    throw new Error('Supabase nao configurado. Crie config/supabase.json.');
  }
  return supabase;
}

function safeStorageName(value) {
  return String(value || 'arquivo')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'arquivo';
}

function safeCacheName(value) {
  return safeStorageName(value.replace(/[\\/]/g, '-'));
}

function createId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

module.exports = {
  isCloudEnabled,
  listarPropostasCloud,
  salvarPropostaCloud,
  atualizarPdfCloud,
  excluirPropostaCloud,
  listarModelosCloud,
  salvarModeloCloud,
  excluirModeloCloud,
  resolveCloudFile
};
