const fs = require('node:fs');
const path = require('node:path');
const { spawn } = require('node:child_process');

const commonSofficePaths = [
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  '/Applications/LibreOffice.app/Contents/MacOS/soffice',
  '/usr/bin/soffice',
  '/usr/local/bin/soffice'
];

function findSoffice() {
  if (process.env.SOFFICE_PATH && fs.existsSync(process.env.SOFFICE_PATH)) {
    return process.env.SOFFICE_PATH;
  }

  const pathNames = process.env.PATH ? process.env.PATH.split(path.delimiter) : [];
  const executableNames = process.platform === 'win32'
    ? ['soffice.exe', 'soffice.com', 'soffice']
    : ['soffice'];

  for (const directory of pathNames) {
    for (const executable of executableNames) {
      const candidate = path.join(directory, executable);
      if (fs.existsSync(candidate)) {
        return candidate;
      }
    }
  }

  return commonSofficePaths.find((candidate) => fs.existsSync(candidate)) || '';
}

function gerarPdf(docxPath, outputDir = path.dirname(docxPath)) {
  if (!docxPath || !fs.existsSync(docxPath)) {
    return Promise.reject(new Error(`Arquivo DOCX não encontrado: ${docxPath || ''}`));
  }

  fs.mkdirSync(outputDir, { recursive: true });

  const soffice = findSoffice();
  if (soffice) {
    return convertWithLibreOffice(soffice, docxPath, outputDir).catch((error) => {
      if (process.platform === 'win32') {
        return convertWithWord(docxPath, outputDir).catch(() => Promise.reject(error));
      }
      return Promise.reject(error);
    });
  }

  if (process.platform === 'win32') {
    return convertWithWord(docxPath, outputDir).catch(() => ({
      converted: false,
      reason: 'LibreOffice ou Microsoft Word não encontrado para converter PDF',
      docxPath
    }));
  }

  return Promise.resolve({
    converted: false,
    reason: 'LibreOffice não encontrado',
    docxPath
  });
}

function convertWithLibreOffice(soffice, docxPath, outputDir) {
  return new Promise((resolve, reject) => {
    const child = spawn(soffice, [
      '--headless',
      '--convert-to',
      'pdf',
      '--outdir',
      outputDir,
      docxPath
    ], {
      windowsHide: true
    });

    let stderr = '';
    child.stderr.on('data', (chunk) => {
      stderr += chunk.toString();
    });

    child.on('error', (error) => {
      reject(error);
    });

    child.on('exit', (code) => {
      if (code !== 0) {
        reject(new Error(stderr || `LibreOffice saiu com código ${code}`));
        return;
      }

      const pdfPath = path.join(outputDir, `${path.basename(docxPath, path.extname(docxPath))}.pdf`);
      if (!fs.existsSync(pdfPath)) {
        reject(new Error('LibreOffice concluiu, mas o PDF não foi encontrado.'));
        return;
      }

      resolve({
        converted: true,
        method: 'libreoffice',
        pdfPath,
        docxPath
      });
    });
  });
}

function convertWithWord(docxPath, outputDir) {
  return new Promise((resolve, reject) => {
    const pdfPath = path.join(outputDir, `${path.basename(docxPath, path.extname(docxPath))}.pdf`);
    const script = `
$ErrorActionPreference = 'Stop'
$docxPath = ${toPowerShellString(path.resolve(docxPath))}
$pdfPath = ${toPowerShellString(path.resolve(pdfPath))}
$word = $null
$document = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $document = $word.Documents.Open($docxPath, $false, $true)
  $document.ExportAsFixedFormat($pdfPath, 17)
}
finally {
  if ($document -ne $null) { $document.Close($false) | Out-Null }
  if ($word -ne $null) { $word.Quit() | Out-Null }
}
`;

    const child = spawn('powershell.exe', [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-Command',
      script
    ], {
      windowsHide: true
    });

    let stderr = '';
    child.stderr.on('data', (chunk) => {
      stderr += chunk.toString();
    });

    child.on('error', reject);
    child.on('exit', (code) => {
      if (code !== 0) {
        reject(new Error(stderr || `Microsoft Word saiu com código ${code}`));
        return;
      }

      if (!fs.existsSync(pdfPath)) {
        reject(new Error('Microsoft Word concluiu, mas o PDF não foi encontrado.'));
        return;
      }

      resolve({
        converted: true,
        method: 'word',
        pdfPath,
        docxPath
      });
    });
  });
}

function toPowerShellString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

module.exports = {
  gerarPdf,
  findSoffice
};

if (require.main === module) {
  const docxPath = process.argv[2];
  gerarPdf(docxPath)
    .then((result) => console.log(JSON.stringify(result, null, 2)))
    .catch((error) => {
      console.error(error);
      process.exit(1);
    });
}
