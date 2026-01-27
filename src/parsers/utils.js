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

/**
 * 가맹점명/설명을 기반으로 카테고리 추정
 */
export function guessCategory(description, memo = '', isIncome = false) {
  const text = `${description} ${memo}`.toLowerCase();

  // 수입인 경우
  if (isIncome) {
    if (text.includes('급여') || text.includes('월급') || text.includes('급료')) {
      return { main: '주수입', sub: '급여' };
    }
    if (text.includes('인센티브') || text.includes('보너스') || text.includes('상여')) {
      return { main: '주수입', sub: '인센티브' };
    }
    if (text.includes('이자') || text.includes('캐시백') || text.includes('리워드')) {
      return { main: '부수입', sub: '이자캐시백' };
    }
    if (text.includes('포인트')) {
      return { main: '부수입', sub: '포인트적립' };
    }
    return { main: '부수입', sub: '부업' };
  }

  // 편의점
  if (text.includes('gs25') || text.includes('cu ') || text.includes('세븐일레븐') ||
      text.includes('이마트24') || text.includes('미니스톱')) {
    return { main: '식비', sub: '음료간식' };
  }

  // 외식 (음식점)
  if (text.includes('삼겹살') || text.includes('마라탕') || text.includes('곱창') ||
      text.includes('회뜨는') || text.includes('고깃집') || text.includes('식당') ||
      text.includes('치킨') || text.includes('피자') || text.includes('햄버거') ||
      text.includes('한우') || text.includes('소고기') || text.includes('돼지') ||
      text.includes('구이') || text.includes('탕') || text.includes('찌개')) {
    return { main: '식비', sub: '외식배달' };
  }

  // 배달앱
  if (text.includes('쿠팡이츠') || text.includes('배민') || text.includes('요기요') || text.includes('배달')) {
    return { main: '식비', sub: '외식배달' };
  }

  // 식자재/마트
  if (text.includes('마트') || text.includes('롯데') || text.includes('이마트') ||
      text.includes('홈플러스') || text.includes('식자재') || text.includes('농협')) {
    return { main: '식비', sub: '식자재' };
  }

  // 카페
  if (text.includes('스타벅스') || text.includes('카페') || text.includes('커피') ||
      text.includes('빵') || text.includes('베이커리')) {
    return { main: '식비', sub: '음료간식' };
  }

  // 온라인쇼핑
  if (text.includes('쿠팡') || text.includes('네이버페이') || text.includes('11번가') || text.includes('지마켓')) {
    return { main: '생활용품', sub: '생활용품' };
  }
  if (text.includes('다이소') || text.includes('올리브영')) {
    return { main: '생활용품', sub: '생활용품' };
  }

  // 주유소
  if (text.includes('주유') || text.includes('기름') || text.includes('gs칼텍스') ||
      text.includes('sk에너지') || text.includes('농협주유') || text.includes('오일')) {
    return { main: '차량교통', sub: '주유비' };
  }

  // 교통
  if (text.includes('택시') || text.includes('카카오t')) {
    return { main: '차량교통', sub: '택시비' };
  }
  if (text.includes('버스') || text.includes('지하철') || text.includes('교통')) {
    return { main: '차량교통', sub: '대중교통비' };
  }
  if (text.includes('주차') || text.includes('톨게이트') || text.includes('하이패스')) {
    return { main: '차량교통', sub: '주차톨게비' };
  }

  // 통신
  if (text.includes('skt') || text.includes('kt') || text.includes('lg u+') || text.includes('통신')) {
    return { main: '주거통신', sub: '이동통신' };
  }
  if (text.includes('가스') || text.includes('도시가스')) {
    return { main: '주거통신', sub: '도시가스' };
  }

  // 금융
  if (text.includes('보험') || text.includes('삼성화재') || text.includes('현대해상')) {
    return { main: '금융지출', sub: '보험' };
  }

  // 건강/의료
  if (text.includes('병원') || text.includes('약국') || text.includes('의원')) {
    return { main: '건강문화', sub: '병원/약' };
  }
  if (text.includes('헬스') || text.includes('gym') || text.includes('피트니스')) {
    return { main: '건강문화', sub: '운동취미' };
  }

  // 문화
  if (text.includes('넷플릭스') || text.includes('왓챠') || text.includes('영화') ||
      text.includes('cgv') || text.includes('롯데시네마')) {
    return { main: '건강문화', sub: '문화생활' };
  }

  // 경조사
  if (text.includes('축의금') || text.includes('부의금') || text.includes('경조사')) {
    return { main: '경조회비', sub: '경조사비' };
  }

  return { main: '기타지출', sub: '기타지출' };
}
