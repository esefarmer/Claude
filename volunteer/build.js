// 使い方: node build.js  → 同じフォルダに Word ファイル3点を出力
const fs = require('fs');
const path = require('path');
const { Packer } = require('docx');

const outputs = {
  '01_京都翔英高等学校ボランティア受入規程.docx': require('./build_kitei'),
  '02_ボランティア活動の手引き.docx': require('./build_tebiki'),
  '03_ボランティア受入様式集.docx': require('./build_yoshiki'),
};

(async () => {
  for (const [name, d] of Object.entries(outputs)) {
    fs.writeFileSync(path.join(__dirname, name), await Packer.toBuffer(d));
    console.log('wrote', name);
  }
})();
