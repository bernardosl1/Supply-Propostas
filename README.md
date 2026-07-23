# Supply-Propostas

## Importação de propostas com GPT-5.6 Luna

O botão **Selecionar .docx** extrai localmente os parágrafos e as tabelas do
documento e envia esse conteúdo à Responses API da OpenAI. O único modelo usado
para interpretar e preencher o formulário é `gpt-5.6-luna`.

Além dos dados comerciais, objeto e escopo, o modelo reconstrói o restante do
documento na ordem original usando as opções disponíveis na parte inferior do
formulário:

- **Tópico:** textos, observações, cláusulas e subtópicos;
- **Documentação:** descrição, número do documento e data;
- **Preço:** serviços, consumíveis, outros grupos de preço e condições comerciais;
- **Tabela:** qualquer tabela que não seja documentação nem preço;
- **Quebra de página:** marcadores explícitos de página encontrados no DOCX.

Na primeira utilização:

1. Crie uma chave de API no painel da OpenAI e habilite faturamento para o projeto.
2. Abra o formulário e cole a chave no campo `Chave OpenAI (sk-...)`.
3. Clique em **Salvar chave** e depois em **Selecionar .docx**.
4. Revise os dados preenchidos antes de gerar a nova proposta.

A chave é criptografada pelo `safeStorage` do Electron e fica no diretório de
dados do usuário. Ela não é incluída no repositório nem enviada ao renderer após
ser salva. Como alternativa para desenvolvimento, defina `OPENAI_API_KEY` no
ambiente antes de iniciar o aplicativo; a variável de ambiente tem prioridade.

O aplicativo envia apenas o texto e as tabelas extraídos do DOCX. Imagens e
detalhes puramente visuais do Word não são reconstruídos pelo modelo, portanto a
revisão humana continua obrigatória.

O teste automatizado não chama a API e não gera custo:

```powershell
$env:ELECTRON_RUN_AS_NODE='1'
& '.\node_modules\electron\dist\electron.exe' scripts\testar-docx-import.js
```

## Modelos compartilhados no Supabase

Os modelos guardam uma cópia completa da proposta, incluindo campos, itens,
valores e blocos personalizados. Eles são sincronizados pela tabela
`public.modelos_propostas` para ficarem visíveis em todas as máquinas. Antes de distribuir uma versão com
essa funcionalidade, abra o **SQL Editor** do projeto Supabase, cole o conteúdo
de `config/supabase-modelos.sql` e execute uma vez.

O aplicativo não possui login individual. Por isso, as políticas do arquivo SQL
permitem que clientes configurados com a chave pública listem, criem, alterem e
excluam modelos. Para restringir alterações por usuário será necessário adicionar
autenticação ao aplicativo e substituir essas políticas.
