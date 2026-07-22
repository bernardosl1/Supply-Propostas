const fs = require('node:fs');
const path = require('node:path');

function createOpenAiConfig({ app, safeStorage }) {
  const configPath = path.join(app.getPath('userData'), 'openai-config.json');

  function getApiKey() {
    const environmentKey = String(process.env.OPENAI_API_KEY || '').trim();
    if (environmentKey) return environmentKey;
    if (!fs.existsSync(configPath)) return '';
    if (!safeStorage.isEncryptionAvailable()) return '';
    try {
      const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
      if (!config.encryptedKey) return '';
      return safeStorage.decryptString(Buffer.from(config.encryptedKey, 'base64')).trim();
    } catch {
      return '';
    }
  }

  function saveApiKey(value) {
    const apiKey = String(value || '').trim();
    if (!apiKey) {
      if (fs.existsSync(configPath)) fs.unlinkSync(configPath);
      return { configured: false, source: 'none' };
    }
    if (!apiKey.startsWith('sk-')) {
      throw new Error('A chave informada não parece ser uma chave válida da OpenAI.');
    }
    if (!safeStorage.isEncryptionAvailable()) {
      throw new Error('O Windows não disponibilizou armazenamento seguro para proteger a chave.');
    }
    fs.mkdirSync(path.dirname(configPath), { recursive: true });
    const encryptedKey = safeStorage.encryptString(apiKey).toString('base64');
    fs.writeFileSync(configPath, JSON.stringify({ encryptedKey }), { encoding: 'utf8', mode: 0o600 });
    return { configured: true, source: 'encrypted' };
  }

  function getStatus() {
    const fromEnvironment = Boolean(String(process.env.OPENAI_API_KEY || '').trim());
    return {
      configured: Boolean(getApiKey()),
      source: fromEnvironment ? 'environment' : (fs.existsSync(configPath) ? 'encrypted' : 'none'),
      model: 'gpt-5.6-luna'
    };
  }

  return { getApiKey, saveApiKey, getStatus };
}

module.exports = { createOpenAiConfig };
