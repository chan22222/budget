import XLSX from 'xlsx';

// 자동 시도할 비밀번호 목록
export const AUTO_PASSWORDS = ['891117', '19891117'];

/**
 * 엑셀 파일 열기 (비밀번호 자동 시도)
 */
export function openExcelFile(filePath, password = '') {
  // 사용자 입력 비밀번호 + 자동 비밀번호 모두 시도
  const passwords = password ? [password, ...AUTO_PASSWORDS] : ['', ...AUTO_PASSWORDS];
  let lastError = null;

  for (const pw of passwords) {
    try {
      const options = pw ? { password: pw } : {};
      const workbook = XLSX.readFile(filePath, options);
      console.log(`파일 열기 성공${pw ? ` (비밀번호: ${pw.substring(0,2)}***)` : ''}`);
      return workbook;
    } catch (e) {
      lastError = e;
      console.log(`비밀번호 시도 실패 (${pw || '없음'}): ${e.message}`);
      // 비밀번호 관련 오류면 다음 시도
      if (e.message.includes('password') ||
          e.message.includes('encrypt') ||
          e.message.includes('CFB') ||
          e.message.includes('Unsupported') ||
          e.message.includes('corrupted')) {
        continue;
      }
      // 다른 오류면 그대로 던짐
      throw e;
    }
  }

  // 모든 비밀번호 실패
  console.log('모든 비밀번호 실패:', lastError?.message);
  console.log('마지막 오류 전체:', lastError);

  // 실제 오류 메시지 전달 (디버깅용)
  const err = new Error('NEED_PASSWORD');
  err.originalError = lastError?.message;
  throw err;
}
