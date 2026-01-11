import XLSX from 'xlsx';
import { openExcelFile } from './utils.js';

/**
 * 인천e음 카드 거래내역 파싱
 * 컬럼: 거래일시, 카드번호, 결제처, 거래방식, 승인번호, 거래금액, 총 결제금액, 충전잔액, 내 캐시, 공급가액
 */
export function parseEeum(filePath, password = '') {
  const workbook = openExcelFile(filePath, password);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // 헤더 찾기 (거래일시가 있는 행)
  let headerIdx = data.findIndex(row =>
    row.some(cell => String(cell).includes('거래일시'))
  );

  if (headerIdx === -1) {
    throw new Error('인천e음 형식을 인식할 수 없습니다.');
  }

  const transactions = [];

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    const dateStr = String(row[0]); // 2025/12/19 21:14:46
    const txType = String(row[3] || ''); // 충전/결제
    const amount = Number(row[5]) || 0;
    const balance = Number(row[7]) || 0;

    // 날짜 파싱
    const dateParts = dateStr.match(/(\d{4})\/(\d{2})\/(\d{2})/);
    if (!dateParts) continue;

    const date = `${dateParts[1]}-${dateParts[2]}-${dateParts[3]}`;

    // 충전은 제외 (수입으로 처리하지 않음 - 카드 충전이므로)
    if (txType === '충전') continue;

    transactions.push({
      date,
      category: '기타지출',
      subcategory: '기타지출',
      description: `인천e음 결제`,
      incomeAmount: 0,
      expenseAmount: amount,
      paymentMethod: '인천e음',
      expenseType: '변동',
      memo: '인천e음',
      source: '인천e음',
      balance
    });
  }

  return transactions;
}
