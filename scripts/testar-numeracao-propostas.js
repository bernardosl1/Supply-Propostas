const assert = require('node:assert/strict');
const { resolverNumeroDocumento, separarNumeroDocumento } = require('../src/lib/numeracaoProposta');

const propostas = [
  proposta('original', '2026550'),
  proposta('revisao-1', '2026550-R01')
];

assert.equal(resolverNumeroDocumento('NOVO-100', propostas), 'NOVO-100');
assert.equal(resolverNumeroDocumento('2026550', [proposta('original', '2026550')]), '2026550-R01');
assert.equal(resolverNumeroDocumento('2026550', propostas), '2026550-R02');
assert.equal(resolverNumeroDocumento('2026550-R01', propostas), '2026550-R02');
assert.equal(resolverNumeroDocumento('2026550-COPIA', propostas), '2026550-R02');
assert.equal(resolverNumeroDocumento('2026550', propostas, 'original'), '2026550');
assert.equal(resolverNumeroDocumento('2026550-R01', propostas, 'revisao-1'), '2026550-R01');
assert.equal(
  resolverNumeroDocumento('2026550', [...propostas, proposta('revisao-9', '2026550-R09')]),
  '2026550-R10'
);
assert.equal(
  resolverNumeroDocumento('sm-10', [proposta('original', 'SM-10')]),
  'sm-10-R01'
);
assert.deepEqual(
  separarNumeroDocumento('2026550-R02-COPIA'),
  {
    base: '2026550',
    baseNormalizada: '2026550',
    revisao: 2,
    numeroSemCopia: '2026550-R02'
  }
);

console.log('Numera\u00e7\u00e3o de propostas validada.');

function proposta(id, numero_documento) {
  return { id, numero_documento, data: { numero_documento } };
}
