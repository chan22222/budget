import XLSX from 'xlsx';
import XlsxPopulate from 'xlsx-populate';

// 자동 시도할 비밀번호 목록
export const AUTO_PASSWORDS = ['891117', '19891117'];

/**
 * 엑셀 파일 열기 (비밀번호 자동 시도)
 * 각 비밀번호에 대해 xlsx-populate와 xlsx 모두 시도
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

      // xlsx-populate -> xlsx 형식으로 변환
      const buffer = await workbook.outputAsync();
      const xlsxWorkbook = XLSX.read(buffer, { type: 'buffer' });
      return xlsxWorkbook;
    } catch (e) {
      console.log(`xlsx-populate 시도 실패 (${pw || '없음'}): ${e.message}`);
      lastError = e;
    }

    // 2. xlsx로 시도 (같은 비밀번호로)
    try {
      const options = pw ? { password: pw } : {};
      const workbook = XLSX.readFile(filePath, options);
      console.log(`파일 열기 성공 (xlsx)${pw ? ` (비밀번호: ${pw.substring(0,2)}***)` : ''}`);
      return workbook;
    } catch (e) {
      console.log(`xlsx 시도 실패 (${pw || '없음'}): ${e.message}`);
      lastError = e;
    }
  }

  // 모든 비밀번호 실패
  console.log('모든 비밀번호 실패:', lastError?.message);

  const err = new Error('NEED_PASSWORD');
  err.originalError = lastError?.message;
  throw err;
}
