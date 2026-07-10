# Configuracao Supabase

Este app pode salvar propostas em nuvem usando Supabase Database + Supabase Storage.

## 1. Criar o projeto

1. Abra o Supabase e crie um projeto.
2. No painel do projeto, abra **SQL Editor**.
3. Cole e execute o arquivo [`supabase/schema.sql`](./supabase/schema.sql).

Isso cria:

- tabela `public.propostas`;
- bucket privado `propostas`;
- policies iniciais para permitir uso pela anon key do app.

## 2. Criar a configuracao local

Copie:

```text
config/supabase.example.json
```

para:

```text
config/supabase.json
```

Preencha com os dados do Supabase:

```json
{
  "url": "https://seu-projeto.supabase.co",
  "anonKey": "sua-anon-public-key",
  "bucket": "propostas"
}
```

Onde encontrar:

- `url`: Project Settings > API > Project URL.
- `anonKey`: Project Settings > API > Project API keys > anon/public key.

## 3. Como o app se comporta

- Com `config/supabase.json`: lista, salva e exclui propostas no Supabase.
- Sem `config/supabase.json`: continua usando o historico local.
- Se a internet/Supabase falhar: o app tenta manter o fluxo local.
- DOCX/PDF sao enviados para o bucket `propostas`.
- Em outro computador, ao listar historico, o app baixa o arquivo remoto quando o usuario clicar em abrir/exportar.

## 4. Seguranca

A configuracao inicial usa policies para `anon` para facilitar o primeiro uso sem tela de login.
Isso funciona para um app interno, mas a anon key fica embutida/disponivel no app instalado.

Para producao, o recomendado e adicionar login e trocar as policies para `authenticated`.
