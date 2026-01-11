import XLSX from 'xlsx';

/**
 * 토스뱅크 거래내역 파싱
 * 컬럼: 거래 일시, 적요, 거래 유형, 거래 기관, 계좌번호, 거래 금액, 거래 후 잔액, 메모
 */
export function parseTossBank(filePath) {
  const workbook = XLSX.readFile(filePath);
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
    const incomeAmount = isIncome ? amount : 0;
    const expenseAmount = isIncome ? 0 : Math.abs(amount);

    // 지출수단 결정
    let paymentMethod = '체크카드';
    if (txType.includes('송금') || txType.includes('출금')) {
      paymentMethod = '이체';
    } else if (txType.includes('입금')) {
      paymentMethod = '입금';
    }

    // 대분류 추정
    let category = guessCategory(description, memo);

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
      expenseType: isIncome ? '' : '변동',
      memo: '[토스뱅크]',
      source: '토스뱅크',
      balance
    });
  }

  return transactions;
}

function guessCategory(description, memo) {
  const text = `${description} ${memo}`.toLowerCase();

  // 식비
  if (text.includes('쿠팡이츠') || text.includes('배민') || text.includes('요기요') || text.includes('배달')) {
    return { main: '식비', sub: '외식배달' };
  }
  if (text.includes('마트') || text.includes('롯데') || text.includes('이마트') || text.includes('홈플러스') || text.includes('식자재')) {
    return { main: '식비', sub: '식자재' };
  }
  if (text.includes('스타벅스') || text.includes('카페') || text.includes('커피') || text.includes('빵')) {
    return { main: '식비', sub: '음료간식' };
  }

  // 생활용품
  if (text.includes('쿠팡') || text.includes('네이버페이') || text.includes('11번가') || text.includes('지마켓')) {
    return { main: '생활용품', sub: '생활용품' };
  }
  if (text.includes('다이소') || text.includes('올리브영')) {
    return { main: '생활용품', sub: '생활용품' };
  }

  // 차량교통
  if (text.includes('주유') || text.includes('기름') || text.includes('gs칼텍스') || text.includes('sk에너지')) {
    return { main: '차량교통', sub: '주유비' };
  }
  if (text.includes('택시') || text.includes('카카오t')) {
    return { main: '차량교통', sub: '택시비' };
  }
  if (text.includes('버스') || text.includes('지하철') || text.includes('교통')) {
    return { main: '차량교통', sub: '대중교통비' };
  }
  if (text.includes('주차') || text.includes('톨게이트') || text.includes('하이패스')) {
    return { main: '차량교통', sub: '주차톨게비' };
  }

  // 주거통신
  if (text.includes('skt') || text.includes('kt') || text.includes('lg u+') || text.includes('통신')) {
    return { main: '주거통신', sub: '이동통신' };
  }
  if (text.includes('가스') || text.includes('도시가스')) {
    return { main: '주거통신', sub: '도시가스' };
  }

  // 금융지출
  if (text.includes('보험') || text.includes('삼성화재') || text.includes('현대해상')) {
    return { main: '금융지출', sub: '보험' };
  }

  // 건강문화
  if (text.includes('병원') || text.includes('약국') || text.includes('의원')) {
    return { main: '건강문화', sub: '병원/약' };
  }
  if (text.includes('헬스') || text.includes('gym') || text.includes('피트니스')) {
    return { main: '건강문화', sub: '운동취미' };
  }
  if (text.includes('넷플릭스') || text.includes('왓챠') || text.includes('영화') || text.includes('cgv') || text.includes('롯데시네마')) {
    return { main: '건강문화', sub: '문화생활' };
  }

  // 경조회비
  if (text.includes('축의금') || text.includes('부의금') || text.includes('경조사')) {
    return { main: '경조회비', sub: '경조사비' };
  }

  return { main: '기타지출', sub: '기타지출' };
}
