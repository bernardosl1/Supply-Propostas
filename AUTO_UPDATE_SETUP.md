# Auto-update com GitHub Releases

O app ja tem `electron-updater` configurado em `main.js`.

O `package.json` esta configurado para publicar em:

```text
https://github.com/bernardosl1/Supply-Propostas
```

Depois disso:

1. Gere uma versao maior em `package.json`, por exemplo `1.0.1`.
2. Rode o build com publicacao:

```powershell
$env:PATH="C:\projetos\propostas andre\.tools\node;$env:PATH"
$env:GH_TOKEN="SEU_TOKEN_DO_GITHUB"
node .\node_modules\electron-builder\cli.js --win nsis portable --publish always
```

O GitHub Release precisa conter pelo menos:

- `Supply-Marine-Propostas-Setup-x.y.z.exe`
- `Supply-Marine-Propostas-Setup-x.y.z.exe.blockmap`
- `latest.yml`

Quando o usuario abrir o app instalado, ele verifica atualizacao automaticamente.
Ao baixar uma nova versao, o app pergunta se deve reiniciar para instalar.
