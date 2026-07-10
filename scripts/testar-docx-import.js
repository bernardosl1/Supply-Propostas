const assert = require('node:assert/strict');
const path = require('node:path');
const { importarDocx } = require('../src/lib/importarDocx');

const filePath = path.resolve(__dirname, '..', 'propostas', '20260200-teste.docx');
const imported = importarDocx(filePath);

assert.equal(imported.data.empresa_cliente, 'Oceânica Engenharia E Consultoria Ltda');
assert.equal(imported.data.numero_documento, '20260200');
assert.equal(imported.data.responsavel_nome, 'André Luis Souza');
assert.equal(imported.data.objeto, 'Manutenção do ar-condicionado de BB & BE');
assert.equal(imported.data.local_servico, 'Maceió | AL');
assert.equal(imported.data.data_servico, '13/03/2026');
assert.equal(imported.data.prazo_execucao_dias, '03 (três) dias');
assert.equal(imported.servicos_descricao.length, 2);
assert.equal(imported.itens_servico.length, 4);
assert.equal(imported.itens_consumiveis.length, 1);
assert.equal(imported.itens_servico[0].descricao, 'Técnico HVAC-R');
assert.equal(imported.itens_servico[0].valor_unit, 310);
assert.equal(imported.itens_consumiveis[0].valor_total, 1624);

console.log('Importação DOCX validada.');
