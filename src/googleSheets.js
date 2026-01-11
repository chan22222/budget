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

  // 가계부 형식으로 변환
  return rows
    .filter(row => row[0]) // 날짜가 있는 행만
    .map(row => ({
      날짜: row[0] || '',
      대분류: row[1] || '',
      소분류: row[2] || '',
      내용: row[3] || '',
      수입금액: row[4] || '',
      지출금액: row[5] || '',
      지출수단: row[6] || '',
      지출성격: row[7] || '',
      비고: row[8] || ''
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
  let existingData = [];
  try {
    const existing = await sheets.spreadsheets.values.get({
      spreadsheetId: currentSpreadsheetId,
      range: `${sheetName}!B11:J999`,
    });
    existingData = existing.data.values || [];
  } catch (e) {
    console.log('기존 데이터 읽기 실패:', e.message);
  }

  // 기존 데이터 키 생성 (날짜 + 내용 + 수입금액 + 지출금액)
  const existingKeys = new Set();
  existingData.forEach(row => {
    if (row[0]) {
      const key = `${row[0]}|${row[3] || ''}|${row[4] || ''}|${row[5] || ''}`;
      existingKeys.add(key);
    }
  });

  // 중복 제거
  const newData = data.filter(row => {
    const key = `${row.날짜}|${row.내용 || ''}|${row.수입금액 || ''}|${row.지출금액 || ''}`;
    return !existingKeys.has(key);
  });

  if (newData.length === 0) {
    return { updatedRows: 0, message: '모든 데이터가 이미 존재합니다.' };
  }

  // 데이터를 배열 형식으로 변환
  const values = newData.map(row => [
    row.날짜,
    row.대분류,
    row.소분류,
    row.내용,
    row.수입금액 || '',
    row.지출금액 || '',
    row.지출수단 || '',
    row.지출성격 || '',
    row.비고 || ''
  ]);

  // 기존 데이터 다음 빈 행에 추가
  const response = await sheets.spreadsheets.values.append({
    spreadsheetId: currentSpreadsheetId,
    range: `${sheetName}!B11:J`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values }
  });

  return {
    ...response.data,
    addedCount: newData.length,
    skippedCount: data.length - newData.length
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

  const values = data.map(row => [
    row.날짜,
    row.대분류,
    row.소분류,
    row.내용,
    row.수입금액 || '',
    row.지출금액 || '',
    row.지출수단 || '',
    row.지출성격 || '',
    row.비고 || ''
  ]);

  const response = await sheets.spreadsheets.values.update({
    spreadsheetId: currentSpreadsheetId,
    range: `${sheetName}!B${startRow}:J${startRow + values.length - 1}`,
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
