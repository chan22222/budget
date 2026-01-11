import XLSX from 'xlsx';

// 자동 시도할 비밀번호 목록
export const AUTO_PASSWORDS = ['891117', '19891117'];

/**
 * 엑셀 파일 열기 (비밀번호 자동 시도)
 */
export function openExcelFile(filePath, password = '') {
  const passwords = password ? [password] : ['', ...AUTO_PASSWORDS];

  for (const pw of passwords) {
    try {
      const options = pw ? { password: pw } : {};
      const workbook = XLSX.readFile(filePath, options);
      return workbook;
    } catch (e) {
      // 비밀번호 오류면 다음 시도
      if (e.message.includes('password') || e.message.includes('encrypt') || e.message.includes('CFB')) {
        continue;
      }
      // 다른 오류면 그대로 던짐
      throw e;
    }
  }

  // 모든 비밀번호 실패
  throw new Error('NEED_PASSWORD');
}
