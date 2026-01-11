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
const REDIRECT_URI = process.env.REDIRECT_URI || 'http://localhost:8080/api/google/callback';

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

/**
 * OAuth2 클라이언트 초기화
 */
export function initOAuth2Client() {
  let client_id, client_secret, redirect_uri;

  // GOOGLE_CREDENTIALS 환경변수 (JSON) 우선
  if (process.env.GOOGLE_CREDENTIALS) {
    try {
      const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
      const creds = credentials?.web || credentials?.installed;
      if (creds) {
        client_id = creds.client_id;
        client_secret = creds.client_secret;
        redirect_uri = creds.redirect_uris?.[0] || REDIRECT_URI;
      }
    } catch (e) {
      console.error('GOOGLE_CREDENTIALS 파싱 실패:', e.message);
    }
  }

  // 개별 환경변수
  if (!client_id) client_id = process.env.GOOGLE_CLIENT_ID;
  if (!client_secret) client_secret = process.env.GOOGLE_CLIENT_SECRET;
  if (!redirect_uri) redirect_uri = process.env.REDIRECT_URI || REDIRECT_URI;

  // credentials.json 파일
  if (!client_id || !client_secret) {
    if (fs.existsSync(CREDENTIALS_PATH)) {
      const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf8'));
      const creds = credentials.web || credentials.installed;
      client_id = client_id || creds.client_id;
      client_secret = client_secret || creds.client_secret;
      redirect_uri = redirect_uri || creds.redirect_uris?.[0];
    }
  }

  if (!client_id || !client_secret) {
    console.log('Google 인증 정보가 없습니다.');
    return null;
  }

  oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uri);

  // 환경변수에서 토큰 로드 또는 파일에서 로드
  if (process.env.GOOGLE_TOKEN) {
    try {
      const token = JSON.parse(process.env.GOOGLE_TOKEN);
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
 * 인증 URL 생성
 */
export function getAuthUrl() {
  if (!oAuth2Client) initOAuth2Client();
  if (!oAuth2Client) return null;

  return oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
}

/**
 * 인증 코드로 토큰 발급
 */
export async function getTokenFromCode(code) {
  if (!oAuth2Client) initOAuth2Client();

  const { tokens } = await oAuth2Client.getToken(code);
  oAuth2Client.setCredentials(tokens);

  // 토큰 저장
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));

  return tokens;
}

/**
 * 인증 상태 확인
 */
export function isAuthenticated() {
  if (!oAuth2Client) initOAuth2Client();
  return oAuth2Client && oAuth2Client.credentials && oAuth2Client.credentials.access_token;
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
 * 스프레드시트에 데이터 쓰기 (추가)
 */
export async function appendToSheet(month, data) {
  if (!isAuthenticated()) {
    throw new Error('Google 인증이 필요합니다.');
  }

  const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
  const sheetName = month || '1월';

  // 데이터를 배열 형식으로 변환
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

  // 기존 데이터 다음 행에 추가 (행 삽입 없이)
  const response = await sheets.spreadsheets.values.append({
    spreadsheetId: currentSpreadsheetId,
    range: `${sheetName}!B11:J`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'OVERWRITE',
    requestBody: { values }
  });

  return response.data;
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
