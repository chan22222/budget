import XLSX from 'xlsx';
import { openExcelFile, guessCategory } from './utils.js';

/**
 * 신용카드 이용내역 파싱 (카드이용내역_YYYYMMDD_YYYYMMDD.xls)
 * 컬럼: 이용일, 이용시간, 이용고객명, 이용카드명, 이용하신곳, 국내이용금액(원),
 *       해외이용금액($), 결제방법, 가맹점정보, 할인금액, 적립포인트, 상태, 결제예정일, 승인번호
 */
export async function parseCreditCard(filePath, password = '') {
  const workbook = await openExcelFile(filePath, password);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // 헤더 찾기 (이용일과 이용하신곳이 모두 있는 행)
  let headerIdx = data.findIndex(row =>
    row.some(cell => String(cell).includes('이용일') && !String(cell).includes('이용일시')) &&
    row.some(cell => String(cell).includes('이용하신곳'))
  );

  if (headerIdx === -1) {
    throw new Error('신용카드 이용내역 형식을 인식할 수 없습니다.');
  }

  const header = data[headerIdx];
  const dateIdx = header.findIndex(cell => String(cell).trim() === '이용일');
  const cardNameIdx = header.findIndex(cell => String(cell).includes('이용카드명'));
  const merchantIdx = header.findIndex(cell => String(cell).includes('이용하신곳'));
  const amountIdx = header.findIndex(cell => String(cell).includes('국내이용금액'));
  const overseasIdx = header.findIndex(cell => String(cell).includes('해외이용금액'));
  const payMethodIdx = header.findIndex(cell => String(cell).includes('결제방법'));
  const statusIdx = header.findIndex(cell => String(cell).includes('상태'));

  const transactions = [];

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row[dateIdx]) continue;

    const dateStr = String(row[dateIdx]); // 2026-05-01
    const cardName = cardNameIdx >= 0 ? String(row[cardNameIdx] || '') : '';
    const merchant = merchantIdx >= 0 ? String(row[merchantIdx] || '') : '';
    const domesticAmount = amountIdx >= 0 ? (Number(String(row[amountIdx]).replace(/,/g, '')) || 0) : 0;
    const overseasAmount = overseasIdx >= 0 ? (Number(String(row[overseasIdx]).replace(/,/g, '')) || 0) : 0;
    const payMethod = payMethodIdx >= 0 ? String(row[payMethodIdx] || '') : '';
    const status = statusIdx >= 0 ? String(row[statusIdx] || '') : '';

    // 취소된 거래 제외
    if (status.includes('취소')) continue;

    const amount = domesticAmount > 0 ? domesticAmount : overseasAmount;
    if (amount === 0) continue;

    // 날짜 파싱 (YYYY-MM-DD)
    const dateParts = dateStr.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (!dateParts) continue;

    const date = `${dateParts[1]}-${dateParts[2]}-${dateParts[3]}`;

    // 가맹점명 기반 카테고리 추정
    const category = guessCategory(merchant, cardName, false);

    // 할부 정보를 비고에 포함
    const installment = payMethod && payMethod !== '일시불' ? payMethod : '';

    transactions.push({
      date,
      category: category.main,
      subcategory: category.sub,
      description: merchant || '카드결제',
      incomeAmount: 0,
      expenseAmount: amount,
      paymentMethod: '신용카드',
      expenseType: '변동지출',
      memo: cardName.replace(/^\([^)]+\)\s*/, '').trim() || '신용카드',
      source: '신용카드',
      installment
    });
  }

  return transactions;
}
