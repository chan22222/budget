import { google } from 'googleapis';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = path.join(__dirname, '..', 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, '..', 'credentials.json');

// 스프레드시트 설정 (환경변수 또는 기본값)
const DEFAULT_SPREADSHEET_ID = '1MK725XvdpkWESa8WlvgNNYGC_XK5UqNf76S8rihkmGM';
let currentSpreadsheetId = process.env.SPREADSHEET_ID || DEFAULT_SPREADSHEET_ID;

// 메모리에 토큰 저장 (서버 재시작시 초기화됨)
let memoryToken = null;

// 스프레드시트 ID getter/setter
export function getSpreadsheetId() {
  return currentSpreadsheetId;
}

export function setSpreadsheetId(input) {
  // URL에서 ID 추출 또는 그대로 사용
  let id = input;
  if (input && input.includes('docs.google.com/spreadsheets')) {
    const match = input.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match) {
      id = match[1];
    }
  }
  currentSpreadsheetId = id || DEFAULT_SPREADSHEET_ID;
  return currentSpreadsheetId;
}
const MONTHS = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];

let oAuth2Client = null;
let clientCredentials = null;

/**
 * credentials 로드
 */
function loadCredentials() {
  if (clientCredentials) return clientCredentials;

  let client_id, client_secret;

  // GOOGLE_CREDENTIALS 환경변수 (JSON) 우선
  if (process.env.GOOGLE_CREDENTIALS) {
    try {
      const credStr = process.env.GOOGLE_CREDENTIALS.replace(/[\r\n]+/g, '').trim();
      const credentials = JSON.parse(credStr);
      const creds = credentials?.web || credentials?.installed;
      if (creds) {
        client_id = creds.client_id;
        client_secret = creds.client_secret;
      }
    } catch (e) {
      console.error('GOOGLE_CREDENTIALS 파싱 실패:', e.message);
    }
  }

  // 개별 환경변수
  if (!client_id) client_id = process.env.GOOGLE_CLIENT_ID;
  if (!client_secret) client_secret = process.env.GOOGLE_CLIENT_SECRET;

  // credentials.json 파일
  if (!client_id || !client_secret) {
    if (fs.existsSync(CREDENTIALS_PATH)) {
      const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf8'));
      const creds = credentials.web || credentials.installed;
      client_id = client_id || creds.client_id;
      client_secret = client_secret || creds.client_secret;
    }
  }

  if (client_id && client_secret) {
    clientCredentials = { client_id, client_secret };
  }
  return clientCredentials;
}

/**
 * OAuth2 클라이언트 생성 (동적 redirect_uri 지원)
 */
export function getOAuth2Client(customRedirectUri) {
  const creds = loadCredentials();
  if (!creds) {
    console.log('Google 인증 정보가 없습니다.');
    return null;
  }

  const redirectUri = customRedirectUri || 'http://localhost:8080/api/google/callback';
  return new google.auth.OAuth2(creds.client_id, creds.client_secret, redirectUri);
}

/**
 * OAuth2 클라이언트 초기화 (기존 호환성)
 */
export function initOAuth2Client() {
  const creds = loadCredentials();
  if (!creds) {
    console.log('Google 인증 정보가 없습니다.');
    return null;
  }

  oAuth2Client = new google.auth.OAuth2(
    creds.client_id,
    creds.client_secret,
    'http://localhost:8080/api/google/callback'
  );

  // 토큰 우선순위: 메모리 > 환경변수 > 로컬파일
  if (memoryToken) {
    oAuth2Client.setCredentials(memoryToken);
  } else if (process.env.GOOGLE_TOKEN) {
    try {
      const tokenStr = process.env.GOOGLE_TOKEN.replace(/[\r\n]+/g, '').trim();
      const token = JSON.parse(tokenStr);
      oAuth2Client.setCredentials(token);
    } catch (e) {
      console.error('GOOGLE_TOKEN 파싱 실패:', e.message);
    }
  } else if (fs.existsSync(TOKEN_PATH)) {
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
    oAuth2Client.setCredentials(token);
  }

  return oAuth2Client;
}

/**
 * 메모리 토큰 설정
 */
export function setMemoryToken(tokens) {
  memoryToken = tokens;
  if (oAuth2Client) {
    oAuth2Client.setCredentials(tokens);
  }
}

/**
 * 인증 URL 생성 (동적 redirect_uri 지원)
 */
export function getAuthUrl(customRedirectUri) {
  const oauth2Client = customRedirectUri
    ? getOAuth2Client(customRedirectUri)
    : (oAuth2Client || initOAuth2Client());

  if (!oauth2Client) return null;

  return oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent'  // 항상 refresh_token 받기
  });
}

/**
 * 인증 코드로 토큰 발급 (동적 redirect_uri 지원)
 */
export async function getTokenFromCode(code, customRedirectUri) {
  const oauth2Client = customRedirectUri
    ? getOAuth2Client(customRedirectUri)
    : (oAuth2Client || initOAuth2Client());

  const { tokens } = await oauth2Client.getToken(code);

  // 메모리에 토큰 저장
  memoryToken = tokens;

  // oAuth2Client도 업데이트
  if (oAuth2Client) {
    oAuth2Client.setCredentials(tokens);
  }

  // 파일에도 저장 (로컬 개발용)
  try {
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
  } catch (e) {
    console.log('토큰 파일 저장 실패 (Railway에서는 정상):', e.message);
  }

  return tokens;
}

/**
 * 인증 상태 확인
 */
export function isAuthenticated() {
  // 메모리 토큰 체크
  if (memoryToken && memoryToken.access_token) return true;

  // 환경변수 토큰 체크
  if (process.env.GOOGLE_TOKEN) {
    try {
      const token = JSON.parse(process.env.GOOGLE_TOKEN.replace(/[\r\n]+/g, '').trim());
      if (token.access_token) return true;
    } catch (e) {}
  }

  // oAuth2Client 체크
  if (!oAuth2Client) initOAuth2Client();
  return oAuth2Client && oAuth2Client.credentials && oAuth2Client.credentials.access_token;
}

/**
 * 인증 소스 확인
 */
export function getAuthSource() {
  if (process.env.GOOGLE_TOKEN) return 'environment';
  if (memoryToken) return 'memory';
  if (fs.existsSync(TOKEN_PATH)) return 'file';
  return 'none';
}

/**
 * 스프레드시트에서 데이터 읽기
 */
export async function readSheetData(month) {
  if (!isAuthenticated()) {
    throw new Error('Google 인증이 필요합니다.');
  }

  const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
  const sheetName = month || '1월';

  // B11:L999 범위 읽기
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: currentSpreadsheetId,
    range: `${sheetName}!B11:L999`,
  });

  const rows = response.data.values || [];

  // 가계부 형식으로 변환 (F, G 열은 빈칸이므로 건너뜀)
  // B:날짜, C:대분류, D:소분류, E:내용, F:빈칸, G:빈칸, H:수입금액, I:지출금액, J:지출수단, K:지출성격, L:비고
  return rows
    .filter(row => row[0]) // 날짜가 있는 행만
    .map(row => ({
      날짜: row[0] || '',
      대분류: row[1] || '',
      소분류: row[2] || '',
      내용: row[3] || '',
      수입금액: row[6] || '',  // H열 (index 6)
      지출금액: row[7] || '',  // I열 (index 7)
      지출수단: row[8] || '',  // J열 (index 8)
      지출성격: row[9] || '',  // K열 (index 9)
      비고: row[10] || ''      // L열 (index 10)
    }));
}

/**
 * 스프레드시트에 데이터 쓰기 (추가) - 중복 제거
 */
export async function appendToSheet(month, data) {
  if (!isAuthenticated()) {
    throw new Error('Google 인증이 필요합니다.');
  }

  const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
  const sheetName = month || '1월';

  // 기존 데이터 읽기 (중복 체크용)
  // B:날짜, C:대분류, D:소분류, E:내용, F:빈칸, G:빈칸, H:수입금액, I:지출금액, J:지출수단, K:지출성격, L:비고
  let existingData = [];
  try {
    const existing = await sheets.spreadsheets.values.get({
      spreadsheetId: currentSpreadsheetId,
      range: `${sheetName}!B11:L999`,
    });
    existingData = existing.data.values || [];
  } catch (e) {
    console.log('기존 데이터 읽기 실패:', e.message);
  }

  // 기존 데이터 파싱 (느슨한 중복 체크용)
  // B:날짜(0), C:대분류(1), D:소분류(2), E:내용(3), F:빈칸(4), G:빈칸(5), H:수입(6), I:지출(7), J:수단(8), K:성격(9), L:비고(10)
  const existingRows = existingData
    .filter(row => row[0])
    .map(row => ({
      날짜: String(row[0] || ''),
      내용: String(row[3] || ''),
      수입금액: String(row[6] || ''),
      지출금액: String(row[7] || ''),
      비고: String(row[10] || '')
    }));

  // 금액을 숫자로 변환 (쉼표 제거)
  function toNumber(val) {
    if (!val && val !== 0) return 0;
    return Number(String(val).replace(/,/g, '')) || 0;
  }

  // 2글자 이상 겹치는지 체크
  function hasCommonChars(str1, str2, minChars = 2) {
    if (!str1 || !str2) return false;
    const s1 = str1.toLowerCase();
    const s2 = str2.toLowerCase();
    // 2글자씩 잘라서 비교
    for (let i = 0; i <= s1.length - minChars; i++) {
      const substr = s1.substring(i, i + minChars);
      if (s2.includes(substr)) return true;
    }
    return false;
  }

  // 중복 체크: 날짜 + 금액 + 내용 2글자 이상 겹침
  console.log(`중복 체크: 기존 ${existingRows.length}건, 신규 ${data.length}건`);

  const duplicates = [];  // 중복 항목 저장
  const newData = data.filter(row => {
    const matchedExisting = existingRows.find(existing => {
      const sameDate = String(row.날짜) === existing.날짜;
      // 금액은 숫자로 변환해서 비교 (쉼표 제거)
      const sameIncome = toNumber(row.수입금액) === toNumber(existing.수입금액);
      const sameExpense = toNumber(row.지출금액) === toNumber(existing.지출금액);
      const sameAmount = sameIncome && sameExpense;
      const contentOverlap = hasCommonChars(row.내용, existing.내용, 2);

      return sameDate && sameAmount && contentOverlap;
    });

    if (matchedExisting) {
      duplicates.push({
        new: `[${row.날짜}일] ${row.내용} (${row.지출금액 || row.수입금액}원)`,
        existing: `[${matchedExisting.날짜}일] ${matchedExisting.내용} (${matchedExisting.지출금액 || matchedExisting.수입금액}원)`
      });
      console.log(`중복 발견: ${row.내용} vs ${matchedExisting.내용}`);
      return false;
    }
    return true;
  });

  console.log(`중복 ${duplicates.length}건, 신규 ${newData.length}건`);

  if (newData.length === 0) {
    return {
      updatedRows: 0,
      message: '모든 데이터가 이미 존재합니다.',
      duplicates: duplicates,
      existingCount: existingRows.length
    };
  }

  // 데이터를 배열 형식으로 변환 (F, G 열은 빈칸)
  // B:날짜, C:대분류, D:소분류, E:내용, F:빈칸, G:빈칸, H:수입금액, I:지출금액, J:지출수단, K:지출성격, L:비고
  const values = newData.map(row => [
    row.날짜,
    row.대분류,
    row.소분류,
    row.내용,
    '',  // F열 빈칸
    '',  // G열 빈칸
    row.수입금액 || '',
    row.지출금액 || '',
    row.지출수단 || '',
    row.지출성격 || '',
    row.비고 || ''
  ]);

  // 기존 데이터 다음 빈 행에 추가
  const response = await sheets.spreadsheets.values.append({
    spreadsheetId: currentSpreadsheetId,
    range: `${sheetName}!B11:L`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values }
  });

  return {
    ...response.data,
    addedCount: newData.length,
    skippedCount: data.length - newData.length,
    duplicates: duplicates,
    existingCount: existingRows.length
  };
}

/**
 * 스프레드시트 특정 범위 업데이트
 */
export async function updateSheet(month, startRow, data) {
  if (!isAuthenticated()) {
    throw new Error('Google 인증이 필요합니다.');
  }

  const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
  const sheetName = month || '1월';

  // F, G 열은 빈칸
  const values = data.map(row => [
    row.날짜,
    row.대분류,
    row.소분류,
    row.내용,
    '',  // F열 빈칸
    '',  // G열 빈칸
    row.수입금액 || '',
    row.지출금액 || '',
    row.지출수단 || '',
    row.지출성격 || '',
    row.비고 || ''
  ]);

  const response = await sheets.spreadsheets.values.update({
    spreadsheetId: currentSpreadsheetId,
    range: `${sheetName}!B${startRow}:L${startRow + values.length - 1}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values }
  });

  return response.data;
}

/**
 * 시트 목록 가져오기
 */
export async function getSheetList() {
  if (!isAuthenticated()) {
    throw new Error('Google 인증이 필요합니다.');
  }

  const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

  const response = await sheets.spreadsheets.get({
    spreadsheetId: currentSpreadsheetId,
  });

  return response.data.sheets.map(s => s.properties.title);
}

export { currentSpreadsheetId, MONTHS };
