import { parseTossBank } from './tossbank.js';
import { parseEeum } from './eeum.js';
import XLSX from 'xlsx';
import path from 'path';

/**
 * 파일 유형 감지 및 적절한 파서 선택
 */
export function parseExcelFile(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // 파일 내용으로 유형 판단
  const content = data.flat().join(' ');

  if (content.includes('토스뱅크')) {
    return parseTossBank(filePath);
  }

  if (content.includes('e음') || content.includes('인천')) {
    return parseEeum(filePath);
  }

  // 파일명으로 판단
  const filename = path.basename(filePath).toLowerCase();
  if (filename.includes('토스') || filename.includes('toss')) {
    return parseTossBank(filePath);
  }

  throw new Error(`알 수 없는 파일 형식: ${path.basename(filePath)}`);
}

/**
 * 가계부 형식으로 변환
 */
export function toBudgetFormat(transactions) {
  return transactions.map(tx => {
    // 날짜에서 일만 추출 (2026-01-15 -> 15)
    let day = tx.date;
    if (tx.date && tx.date.includes('-')) {
      const parts = tx.date.split('-');
      day = parseInt(parts[2], 10).toString(); // 01 -> 1
    }

    return {
      날짜: day,
      대분류: tx.category,
      소분류: tx.subcategory,
      내용: tx.description,
      수입금액: tx.incomeAmount || '',
      지출금액: tx.expenseAmount || '',
      지출수단: tx.paymentMethod,
      지출성격: tx.expenseType,
      비고: tx.memo,
      _fullDate: tx.date  // 월 필터링용 원본 날짜
    };
  });
}

export { parseTossBank, parseEeum };
