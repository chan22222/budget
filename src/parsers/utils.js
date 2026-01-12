import XLSX from 'xlsx';
import XlsxPopulate from 'xlsx-populate';
import { execSync } from 'child_process';
import fs from 'fs';
import path from 'path';
import os from 'os';

// 자동 시도할 비밀번호 목록
export const AUTO_PASSWORDS = ['891117', '19891117'];

/**
 * Python msoffcrypto로 복호화 시도
 */
function tryPythonDecrypt(filePath, password) {
  const tempDir = os.tmpdir();
  const decryptedPath = path.join(tempDir, `decrypted_${Date.now()}.xlsx`);

  try {
    const scriptPath = path.join(process.cwd(), 'decrypt.py');
    const cmd = `python3 "${scriptPath}" "${filePath}" "${decryptedPath}" "${password}"`;

    console.log(`Python 복호화 시도: ${password.substring(0, 2)}***`);
    const result = execSync(cmd, { encoding: 'utf8', timeout: 30000 });

    if (result.includes('SUCCESS:')) {
      console.log('Python 복호화 성공');
      return decryptedPath;
    }
  } catch (e) {
    console.log(`Python 복호화 실패 (${password.substring(0, 2)}***): ${e.message}`);
  }

  return null;
}

/**
 * 엑셀 파일 열기 (비밀번호 자동 시도)
 * 1. xlsx-populate 시도
 * 2. xlsx 시도
 * 3. Python msoffcrypto로 복호화 후 시도
 */
export async function openExcelFile(filePath, password = '') {
  // 사용자 입력 비밀번호 + 자동 비밀번호 모두 시도
  const passwords = password ? [password, ...AUTO_PASSWORDS] : ['', ...AUTO_PASSWORDS];
  let lastError = null;

  for (const pw of passwords) {
    // 1. xlsx-populate로 시도
    try {
      const options = pw ? { password: pw } : {};
      const workbook = await XlsxPopulate.fromFileAsync(filePath, options);
      console.log(`파일 열기 성공 (xlsx-populate)${pw ? ` (비밀번호: ${pw.substring(0,2)}***)` : ''}`);

      const buffer = await workbook.outputAsync();
      const xlsxWorkbook = XLSX.read(buffer, { type: 'buffer' });
      return xlsxWorkbook;
    } catch (e) {
      console.log(`xlsx-populate 시도 실패 (${pw || '없음'}): ${e.message}`);
      lastError = e;
    }

    // 2. xlsx로 시도
    try {
      const options = pw ? { password: pw } : {};
      const workbook = XLSX.readFile(filePath, options);
      console.log(`파일 열기 성공 (xlsx)${pw ? ` (비밀번호: ${pw.substring(0,2)}***)` : ''}`);
      return workbook;
    } catch (e) {
      console.log(`xlsx 시도 실패 (${pw || '없음'}): ${e.message}`);
      lastError = e;
    }

    // 3. Python msoffcrypto로 복호화 시도
    if (pw) {
      const decryptedPath = tryPythonDecrypt(filePath, pw);
      if (decryptedPath && fs.existsSync(decryptedPath)) {
        try {
          const workbook = XLSX.readFile(decryptedPath);
          console.log(`파일 열기 성공 (Python 복호화)${pw ? ` (비밀번호: ${pw.substring(0,2)}***)` : ''}`);

          // 임시 파일 삭제
          try { fs.unlinkSync(decryptedPath); } catch (e) {}

          return workbook;
        } catch (e) {
          console.log(`Python 복호화 파일 읽기 실패: ${e.message}`);
          lastError = e;
          try { fs.unlinkSync(decryptedPath); } catch (e) {}
        }
      }
    }
  }

  // 모든 비밀번호 실패
  console.log('모든 비밀번호 실패:', lastError?.message);

  const err = new Error('NEED_PASSWORD');
  err.originalError = lastError?.message;
  throw err;
}
