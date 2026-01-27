import XLSX from 'xlsx';
import { openExcelFile, guessCategory } from './utils.js';

/**
 * 토스뱅크 거래내역 파싱
 * 컬럼: 거래 일시, 적요, 거래 유형, 거래 기관, 계좌번호, 거래 금액, 거래 후 잔액, 메모
 */
export async function parseTossBank(filePath, password = '') {
  const workbook = await openExcelFile(filePath, password);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // 헤더 찾기 (거래 일시가 있는 행)
  let headerIdx = data.findIndex(row =>
    row.some(cell => String(cell).includes('거래 일시'))
  );

  if (headerIdx === -1) {
    throw new Error('토스뱅크 형식을 인식할 수 없습니다.');
  }

  const transactions = [];

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    const dateStr = String(row[0]); // 2026.01.08 12:49:10
    const description = String(row[1] || '');
    const txType = String(row[2] || '');
    const institution = String(row[3] || '');
    const amount = Number(row[5]) || 0;
    const balance = Number(row[6]) || 0;
    const memo = String(row[7] || '');

    // 날짜 파싱
    const dateParts = dateStr.match(/(\d{4})\.(\d{2})\.(\d{2})/);
    if (!dateParts) continue;

    const date = `${dateParts[1]}-${dateParts[2]}-${dateParts[3]}`;

    // 수입/지출 판단
    const isIncome = amount > 0;
    const incomeAmount = isIncome ? Math.abs(amount) : 0;
    const expenseAmount = !isIncome ? Math.abs(amount) : 0;

    // 지출수단 결정 (현금, 체크카드, 신용카드만 가능)
    let paymentMethod = '체크카드';

    // 대분류 추정 (수입 여부 전달)
    let category = guessCategory(description, memo, isIncome);

    // 내용: 적요 + 메모 병합
    let content = description;
    if (memo && memo !== description) {
      content = memo ? `${description} (${memo})` : description;
    }

    transactions.push({
      date,
      category: category.main,
      subcategory: category.sub,
      description: content,
      incomeAmount,
      expenseAmount,
      paymentMethod,
      expenseType: isIncome ? '' : '변동지출',
      memo: '쀼카드',
      source: '토스뱅크',
      balance
    });
  }

  return transactions;
}
