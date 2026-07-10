const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { listarPropostas, salvarProposta, excluirProposta } = require('../src/lib/armazenamento');

const docxPath = path.resolve(__dirname, '..', 'propostas', 'teste-historico.docx');

fs.mkdirSync(path.dirname(docxPath), { recursive: true });
fs.writeFileSync(docxPath, 'teste');

const proposta = salvarProposta({
  data: {
    numero_documento: 'TESTE-HISTORICO',
    empresa_cliente: 'Cliente Teste',
    data_documento: '07/07/2026',
    preco_total_numero: 123.45,
    itens_servico: [],
    itens_consumiveis: [],
    servicos_descricao: []
  },
  docxPath
});

assert.ok(proposta.id);
assert.ok(listarPropostas().some((item) => item.id === proposta.id));

const removed = excluirProposta(proposta.id);
assert.equal(removed.removed, true);
assert.equal(listarPropostas().some((item) => item.id === proposta.id), false);

console.log('Historico validado.');
