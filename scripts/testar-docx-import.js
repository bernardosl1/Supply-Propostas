const assert = require('node:assert/strict');
const path = require('node:path');
const {
  OPENAI_MODEL,
  importarDocx,
  extrairConteudoDocx
} = require('../src/lib/importarDocx');

async function run() {
  const filePath = path.resolve(__dirname, '..', 'propostas', '20260200-teste.docx');
  const document = extrairConteudoDocx(filePath);
  assert.ok(document.blocks.some((block) => block.type === 'paragraph'));
  assert.ok(document.blocks.some((block) => block.type === 'table'));

  let sentRequest;
  const block = (values) => ({
    tipo: 'texto', titulo: '', observacoes: [], subtopicos: [], linhas_documentacao: [],
    topicos_preco: [], preco_total_numero: 0, preco_total_extenso: '', moeda: '',
    validade_proposta: '', pagamento: '', prazo_entrega: '', frete: '', impostos: '',
    colunas_tabela: [], linhas_tabela: [], ...values
  });
  const modelOutput = {
    empresa_cliente: 'Oceânica Engenharia E Consultoria Ltda',
    unidade: 'Sub VIII',
    processo_fluig: '',
    solicitante_nome_cargo: '',
    contato_email: '',
    contato_telefone: '',
    numero_documento: '20260200',
    data_documento: '13/03/2026',
    responsavel_nome: 'André Luis Souza',
    responsavel_email: '',
    responsavel_telefone: '',
    objeto: 'Manutenção do ar-condicionado de BB & BE',
    servicos_descricao: ['Manutenção preventiva'],
    equipe_tecnica_itens: ['Técnico HVAC-R'],
    local_servico: 'Maceió | AL',
    data_servico: '13/03/2026',
    prazo_execucao_dias: '03 (três) dias',
    blocos_adicionais: [
      block({
        tipo: 'texto', titulo: 'Observações', observacoes: ['Texto preservado'],
        subtopicos: [{
          titulo: 'Prazo', observacoes: ['Três dias'],
          subtopicos: [{ titulo: 'Detalhe', observacoes: ['Após mobilização'] }]
        }]
      }),
      block({
        tipo: 'lista', titulo: 'Documentação',
        linhas_documentacao: [{ descricao: 'Solicitação', numero_documento: 'SC-1', data: '13/03/2026' }]
      }),
      block({
        tipo: 'preco', titulo: 'Preço',
        topicos_preco: [{
          tipo: 'servico', titulo: 'Serviços', itens: [{
            item: '0001', descricao: 'Técnico HVAC-R', ncm: '-----', quant: 1,
            un: 'HH', valor_unit: 310, valor_total: 310
          }]
        }],
        preco_total_numero: 310, preco_total_extenso: 'trezentos e dez reais',
        moeda: 'Real R$', validade_proposta: '30 dias', pagamento: '30 dias',
        prazo_entrega: 'A combinar', frete: 'FOB', impostos: 'inclusos no preço'
      }),
      block({
        tipo: 'tabela', titulo: 'Equipamentos',
        colunas_tabela: ['Equipamento', 'Quantidade'],
        linhas_tabela: [['Bomba', '2']]
      }),
      block({ tipo: 'quebra_pagina' })
    ],
    campos_duvidosos: []
  };

  const fetchImpl = async (_url, request) => {
    sentRequest = JSON.parse(request.body);
    return {
      ok: true,
      status: 200,
      json: async () => ({
        model: OPENAI_MODEL,
        output_text: JSON.stringify(modelOutput),
        usage: { input_tokens: 1000, output_tokens: 500, total_tokens: 1500 }
      })
    };
  };

  const imported = await importarDocx(filePath, { apiKey: 'sk-test-only', fetchImpl });
  assert.equal(sentRequest.model, 'gpt-5.6-luna');
  assert.equal(sentRequest.reasoning.effort, 'low');
  assert.equal(sentRequest.text.format.type, 'json_schema');
  sentRequest.text.format.schema.required.forEach((field) => {
    assert.ok(Object.prototype.hasOwnProperty.call(modelOutput, field), `Campo obrigatório ausente no mock: ${field}`);
  });
  const blockSchema = sentRequest.text.format.schema.properties.blocos_adicionais.items;
  modelOutput.blocos_adicionais.forEach((item) => {
    blockSchema.required.forEach((field) => {
      assert.ok(Object.prototype.hasOwnProperty.call(item, field), `Campo obrigatório ausente no bloco: ${field}`);
    });
  });
  assert.equal(imported.data.empresa_cliente, modelOutput.empresa_cliente);
  assert.equal(imported.data.numero_documento, '20260200');
  assert.equal(imported.data.blocos_adicionais.length, 5);
  assert.deepEqual(imported.data.blocos_adicionais.map((item) => item.tipo), [
    'texto', 'lista', 'preco', 'tabela', 'quebra_pagina'
  ]);
  assert.equal(imported.data.blocos_adicionais[0].subtopicos[0].subtopicos[0].titulo, 'Detalhe');
  assert.equal(imported.data.blocos_adicionais[1].linhas[0].numero_documento, 'SC-1');
  assert.equal(imported.data.blocos_adicionais[3].linhas[0].valores.coluna_1, 'Bomba');
  assert.equal(imported.itens_servico[0].valor_unit, 310);
  assert.equal(imported.ai.total_tokens, 1500);

  await assert.rejects(
    importarDocx(filePath, { apiKey: '' }),
    /Configure a chave da API OpenAI/
  );

  console.log('Importação DOCX via GPT-5.6 Luna validada com resposta simulada.');
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
