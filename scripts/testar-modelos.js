const assert = require('node:assert/strict');
const os = require('node:os');
const path = require('node:path');

process.env.SUPPLY_MARINE_PROPOSTAS_DIR = path.join(os.tmpdir(), `supply-modelos-${process.pid}`);

const { listarModelos, salvarModelo, excluirModelo } = require('../src/lib/armazenamento');
const { normalizarEstruturaModelo } = require('../src/lib/modeloProposta');

const estrutura = normalizarEstruturaModelo({
  _historico_id: 'nao-deve-ser-copiado',
  empresa_cliente: 'Cliente completo',
  numero_documento: 'PROP-123',
  objeto: 'Objeto integral da proposta',
  servicos_descricao: ['Inspeção completa'],
  equipe_tecnica_itens: ['Técnico responsável'],
  secoes_excluidas: ['escopo'],
  ordem_secoes: ['dados_comerciais', 'objeto', 'flex:topico-1', 'flex:tabela-1', 'flex:preco-1'],
  blocos_adicionais: [
    {
      id: 'topico-1',
      tipo: 'texto',
      titulo: 'CRONOGRAMA',
      observacoes: ['Conteúdo integral do modelo'],
      subtopicos: [{ titulo: 'Etapas', observacoes: ['Item preservado'] }]
    },
    {
      id: 'tabela-1',
      tipo: 'tabela',
      titulo: 'EQUIPAMENTOS',
      colunas: [{ id: 'equipamento', nome: 'Equipamento' }],
      linhas: [{ valores: { equipamento: 'Bomba preenchida' } }]
    },
    {
      id: 'preco-1',
      tipo: 'preco',
      topicos_preco: [{ titulo: 'LOCAÇÃO', tipo: 'personalizado', itens: [{ descricao: 'Item', valor: 500 }] }],
      preco_total_numero: 500
    }
  ]
});

assert.deepEqual(estrutura.secoes_excluidas, ['escopo']);
assert.equal(estrutura.empresa_cliente, 'Cliente completo');
assert.equal(estrutura.numero_documento, 'PROP-123');
assert.equal(estrutura.objeto, 'Objeto integral da proposta');
assert.deepEqual(estrutura.servicos_descricao, ['Inspeção completa']);
assert.deepEqual(estrutura.equipe_tecnica_itens, ['Técnico responsável']);
assert.equal(Object.hasOwn(estrutura, '_historico_id'), false);
assert.equal(estrutura.blocos_adicionais[0].titulo, 'CRONOGRAMA');
assert.deepEqual(estrutura.blocos_adicionais[0].observacoes, ['Conteúdo integral do modelo']);
assert.equal(estrutura.blocos_adicionais[0].subtopicos[0].titulo, 'Etapas');
assert.deepEqual(estrutura.blocos_adicionais[0].subtopicos[0].observacoes, ['Item preservado']);
assert.equal(estrutura.blocos_adicionais[1].linhas[0].valores.equipamento, 'Bomba preenchida');
assert.equal(estrutura.blocos_adicionais[1].colunas[0].nome, 'Equipamento');
assert.equal(estrutura.blocos_adicionais[2].topicos_preco[0].itens[0].descricao, 'Item');
assert.equal(estrutura.blocos_adicionais[2].preco_total_numero, 500);

const modelo = salvarModelo({ nome: 'Padrão offshore', empresa: 'Cliente Teste', estrutura });
assert.ok(modelo.id);
assert.equal(listarModelos()[0].nome, 'Padrão offshore');
assert.equal(excluirModelo(modelo.id).removed, true);
assert.equal(listarModelos().length, 0);

console.log('Modelos de propostas validados.');
