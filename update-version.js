const fs = require('fs');
const path = require('path');

// 1. Lê a versão atualizada do package.json
const packageJson = JSON.parse(fs.readFileSync('./package.json', 'utf8'));
const novaVersao = `v${packageJson.version}`;

// 2. Caminho do arquivo Code.js
const codeJsPath = path.join(__dirname, 'src', 'server', 'Code.js');

if (fs.existsSync(codeJsPath)) {
  let conteudo = fs.readFileSync(codeJsPath, 'utf8');

  // 3. Substitui o valor da constante VERSAO_SISTEMA pelo novo número (Suporta variações de espaços e aspas)
  const conteudoAtualizado = conteudo.replace(
    /const VERSAO_SISTEMA\s*=\s*['"].*?['"];/,
    `const VERSAO_SISTEMA = 'V ${packageJson.version}';`
  );

  fs.writeFileSync(codeJsPath, conteudoAtualizado, 'utf8');
  console.log(`\n✅ Versão atualizada com sucesso no Code.js para: ${novaVersao}\n`);
} else {
  console.error(`\n❌ Arquivo Code.js não encontrado em: ${codeJsPath}\n`);
}