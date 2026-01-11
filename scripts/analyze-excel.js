import XLSX from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function analyzeExcel(filePath) {
  let workbook;
  try {
    workbook = XLSX.readFile(filePath);
  } catch (e) {
    console.log(`\n파일 읽기 실패: ${path.basename(filePath)}`);
    console.log(`에러: ${e.message}`);
    return;
  }

  console.log(`\n${'='.repeat(70)}`);
  console.log(`파일: ${path.basename(filePath)}`);
  console.log(`시트 목록: ${workbook.SheetNames.join(', ')}`);
  console.log(`${'='.repeat(70)}`);

  workbook.SheetNames.forEach((sheetName, idx) => {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    console.log(`\n[시트 ${idx + 1}] ${sheetName}`);
    console.log(`행 수: ${data.length}`);

    console.log('\n--- 처음 20행 미리보기 ---');
    for (let i = 0; i < Math.min(20, data.length); i++) {
      const row = data[i];
      if (row && row.length > 0) {
        const values = row.slice(0, 10).map(v => {
          let str = String(v || '');
          if (str.length > 25) str = str.substring(0, 22) + '...';
          return str.padEnd(25);
        });
        console.log(`  행${String(i + 1).padStart(2)}: ${values.join('|')}`);
      }
    }
  });
}

function main() {
  const importDir = path.join(__dirname, '..', 'import');
  const files = fs.readdirSync(importDir).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));

  for (const file of files) {
    analyzeExcel(path.join(importDir, file));
  }
}

main();
