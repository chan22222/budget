import express from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';
import multer from 'multer';
import { parseExcelFile, toBudgetFormat } from './src/parsers/index.js';
import * as gsheets from './src/googleSheets.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const IMPORT_DIR = path.join(__dirname, 'import');
const EXPORT_DIR = path.join(__dirname, 'export');

// export 폴더 생성
if (!fs.existsSync(EXPORT_DIR)) {
  fs.mkdirSync(EXPORT_DIR, { recursive: true });
}

// import 폴더 생성
if (!fs.existsSync(IMPORT_DIR)) {
  fs.mkdirSync(IMPORT_DIR, { recursive: true });
}

// multer 설정 (파일 업로드)
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, IMPORT_DIR),
  filename: (req, file, cb) => {
    const name = Buffer.from(file.originalname, 'latin1').toString('utf8');
    cb(null, name);
  }
});
const upload = multer({ storage });

// 카테고리 정의 (Google Sheets 가계부 형식)
const CATEGORIES = {
  주수입: ['급여', '인센티브'],
  부수입: ['이자캐시백', '부업', '포인트적립'],
  식비: ['식자재', '외식배달', '음료간식', '술/유흥', '업무식사'],
  주거통신: ['임대료', '관리비', '도시가스', '이동통신', 'TV인터넷'],
  생활용품: ['가구가전', '주방욕실', '생활용품'],
  의복미용: ['의류잡화', '헤어뷰티', '세탁수선'],
  건강문화: ['운동취미', '문화생활', '병원/약', '멤버쉽', '교육비'],
  육아교육: ['육아용품', '육아병원비'],
  차량교통: ['주유비', '차량유지비', '대중교통비', '주차톨게비', '택시비'],
  경조회비: ['경조사비', '모임회비', '선물비'],
  금융지출: ['보험', '대출', '주식'],
  기타지출: ['세금과태료', '수수료', '기타지출', '국내여행', '해외여행'],
  저축: ['적금', '예금', '주택청약']
};

// import 폴더 비우기
app.delete('/api/files', (req, res) => {
  try {
    const files = fs.readdirSync(IMPORT_DIR);
    for (const file of files) {
      fs.unlinkSync(path.join(IMPORT_DIR, file));
    }
    console.log(`import 폴더 비움: ${files.length}개 파일 삭제`);
    res.json({ success: true, deleted: files.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// import 폴더의 파일 목록
app.get('/api/files', (req, res) => {
  try {
    const files = fs.readdirSync(IMPORT_DIR)
      .filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'))
      .map(f => ({
        name: f,
        path: path.join(IMPORT_DIR, f),
        size: fs.statSync(path.join(IMPORT_DIR, f)).size
      }));
    res.json(files);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 파일 파싱 및 거래내역 반환
app.get('/api/parse/:filename', async (req, res) => {
  try {
    const filePath = path.join(IMPORT_DIR, req.params.filename);
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: '파일을 찾을 수 없습니다.' });
    }

    const transactions = await parseExcelFile(filePath);
    const budgetData = toBudgetFormat(transactions);

    res.json({
      filename: req.params.filename,
      count: budgetData.length,
      data: budgetData
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 모든 파일 파싱
app.get('/api/parse-all', async (req, res) => {
  try {
    const files = fs.readdirSync(IMPORT_DIR)
      .filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));

    let allTransactions = [];

    for (const file of files) {
      try {
        const filePath = path.join(IMPORT_DIR, file);
        const transactions = await parseExcelFile(filePath);
        allTransactions = allTransactions.concat(transactions);
      } catch (e) {
        console.error(`파일 파싱 실패: ${file}`, e.message);
      }
    }

    // 날짜순 정렬
    allTransactions.sort((a, b) => a.date.localeCompare(b.date));

    const budgetData = toBudgetFormat(allTransactions);

    res.json({
      count: budgetData.length,
      data: budgetData
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 카테고리 목록
app.get('/api/categories', (req, res) => {
  res.json(CATEGORIES);
});

// 파일 업로드 및 파싱
app.post('/api/upload', upload.array('files'), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: '파일이 없습니다.' });
    }

    const password = req.body.password || '';
    let allTransactions = [];
    const fileResults = [];

    const fileDataMap = {};  // 파일별 데이터

    for (const file of req.files) {
      try {
        const transactions = await parseExcelFile(file.path, password);
        const budgetData = toBudgetFormat(transactions);
        allTransactions = allTransactions.concat(transactions);
        fileResults.push({ name: file.originalname, status: 'success', count: transactions.length });
        fileDataMap[file.originalname] = budgetData;
      } catch (e) {
        console.error(`파일 파싱 실패: ${file.originalname}`, e.message);
        if (e.message === 'NEED_PASSWORD') {
          fileResults.push({ name: file.originalname, status: 'need_password', path: file.path });
        } else {
          fileResults.push({ name: file.originalname, status: 'error', error: e.message });
        }
      }
    }

    // 비밀번호 필요한 파일이 있는 경우
    const needPasswordFiles = fileResults.filter(f => f.status === 'need_password');
    if (needPasswordFiles.length > 0 && allTransactions.length === 0) {
      return res.json({
        error: 'NEED_PASSWORD',
        message: '엑셀 비밀번호를 입력해주세요.',
        files: fileResults
      });
    }

    // 날짜순 정렬
    allTransactions.sort((a, b) => a.date.localeCompare(b.date));

    const budgetData = toBudgetFormat(allTransactions);

    res.json({
      count: budgetData.length,
      files: fileResults,
      data: budgetData,
      fileData: fileDataMap,  // 파일별 데이터
      needPasswordFiles: needPasswordFiles.map(f => f.name)
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 단일 파일 파싱 (비밀번호 지원)
app.post('/api/parse-file', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '파일이 없습니다.' });
    }

    const password = req.body.password || '';
    const fileName = Buffer.from(req.file.originalname, 'latin1').toString('utf8');

    try {
      const transactions = await parseExcelFile(req.file.path, password);
      const budgetData = toBudgetFormat(transactions);

      res.json({
        success: true,
        name: fileName,
        count: transactions.length,
        data: budgetData
      });
    } catch (e) {
      if (e.message === 'NEED_PASSWORD') {
        return res.json({
          success: false,
          needPassword: true,
          name: fileName,
          filePath: req.file.path  // 서버 경로 저장
        });
      }
      throw e;
    }
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 저장된 파일 재파싱 (비밀번호로)
app.post('/api/parse-saved', express.json(), async (req, res) => {
  try {
    const { filePath, password, fileName } = req.body;

    if (!filePath || !fs.existsSync(filePath)) {
      return res.status(400).json({ error: '파일을 찾을 수 없습니다.' });
    }

    try {
      const transactions = await parseExcelFile(filePath, password || '');
      const budgetData = toBudgetFormat(transactions);

      res.json({
        success: true,
        name: fileName,
        count: transactions.length,
        data: budgetData
      });
    } catch (e) {
      console.log('parse-saved 오류:', e.message, e.originalError);
      if (e.message === 'NEED_PASSWORD') {
        return res.json({
          success: false,
          needPassword: true,
          name: fileName,
          filePath: filePath,
          originalError: e.originalError  // 원본 오류 메시지
        });
      }
      throw e;
    }
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 거래 업데이트 (카테고리 수정 등)
app.post('/api/update', (req, res) => {
  // TODO: 수정된 데이터 저장
  res.json({ success: true });
});

// 가계부 형식으로 Excel 내보내기
app.post('/api/export', (req, res) => {
  try {
    const { data, month } = req.body;

    const workbook = XLSX.utils.book_new();

    // 헤더 추가
    const headers = ['날짜', '대분류', '소분류', '내용', '수입금액', '지출금액', '지출수단', '지출성격', '비고'];
    const sheetData = [headers, ...data.map(row => [
      row.날짜,
      row.대분류,
      row.소분류,
      row.내용,
      row.수입금액,
      row.지출금액,
      row.지출수단,
      row.지출성격,
      row.비고
    ])];

    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, month || '가계부');

    const filename = `가계부_${month || new Date().toISOString().slice(0, 10)}.xlsx`;
    const exportPath = path.join(EXPORT_DIR, filename);

    XLSX.writeFile(workbook, exportPath);

    res.json({ success: true, filename, path: exportPath });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ============ Google Sheets API ============

// 실제 redirect_uri 계산
function getActualRedirectUri(req) {
  const protocol = req.headers['x-forwarded-proto'] || req.protocol;
  const host = req.headers['x-forwarded-host'] || req.get('host');
  return `${protocol}://${host}/api/google/callback`;
}

// Google 인증 상태 확인
app.get('/api/google/status', (req, res) => {
  try {
    const authenticated = gsheets.isAuthenticated();
    const redirectUri = getActualRedirectUri(req);
    const authUrl = !authenticated ? gsheets.getAuthUrl(redirectUri) : null;
    const spreadsheetId = gsheets.getSpreadsheetId();
    const authSource = gsheets.getAuthSource();
    res.json({ authenticated, authUrl, spreadsheetId, authSource, redirectUri });
  } catch (e) {
    res.json({ authenticated: false, authUrl: null, spreadsheetId: gsheets.getSpreadsheetId(), error: e.message });
  }
});

// 스프레드시트 ID 설정
app.post('/api/google/spreadsheet-id', (req, res) => {
  const { spreadsheetId } = req.body;
  const newId = gsheets.setSpreadsheetId(spreadsheetId);
  res.json({ success: true, spreadsheetId: newId });
});

// Google 인증 시작
app.get('/api/google/auth', (req, res) => {
  const redirectUri = getActualRedirectUri(req);
  console.log('Auth redirect_uri:', redirectUri);
  const authUrl = gsheets.getAuthUrl(redirectUri);
  if (authUrl) {
    res.redirect(authUrl);
  } else {
    res.status(500).json({ error: 'GOOGLE_CREDENTIALS 환경변수가 필요합니다.' });
  }
});

// Google OAuth 콜백
app.get('/api/google/callback', async (req, res) => {
  try {
    const code = req.query.code;
    const redirectUri = getActualRedirectUri(req);
    console.log('Callback redirect_uri:', redirectUri);

    const tokens = await gsheets.getTokenFromCode(code, redirectUri);

    // 토큰을 화면에 표시 (환경변수에 복사할 수 있도록)
    res.send(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>토큰 발급 완료</title>
        <meta charset="utf-8">
        <style>
          body { font-family: Arial, sans-serif; padding: 30px; max-width: 800px; margin: 0 auto; }
          h1 { color: #4CAF50; }
          textarea { width: 100%; height: 120px; font-size: 12px; font-family: monospace; }
          .success { background: #e8f5e9; padding: 15px; border-radius: 8px; margin: 20px 0; }
          .info { background: #e3f2fd; padding: 15px; border-radius: 8px; margin: 20px 0; }
          a { color: #1976D2; }
          button { background: #4CAF50; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; }
          button:hover { background: #388E3C; }
        </style>
      </head>
      <body>
        <h1>Google 인증 완료!</h1>
        <div class="success">
          <strong>메모리에 토큰이 저장되었습니다.</strong><br>
          지금 바로 사용할 수 있습니다.
        </div>

        <div class="info">
          <strong>영구 저장하려면:</strong><br>
          아래 토큰을 Railway 환경변수 <code>GOOGLE_TOKEN</code>에 저장하세요.<br>
          (서버 재시작 후에도 인증 유지)
        </div>

        <textarea id="token" readonly>${JSON.stringify(tokens)}</textarea>
        <br><br>
        <button onclick="navigator.clipboard.writeText(document.getElementById('token').value); alert('복사됨!')">토큰 복사</button>
        <br><br>
        <a href="/">← 메인 페이지로 이동</a>
      </body>
      </html>
    `);
  } catch (e) {
    console.error('OAuth callback error:', e);
    res.status(500).send(`
      <!DOCTYPE html>
      <html>
      <head><title>인증 오류</title><meta charset="utf-8"></head>
      <body style="font-family: Arial; padding: 30px;">
        <h1 style="color: red;">인증 실패</h1>
        <p>${e.message}</p>
        <p>Google Cloud Console에서 redirect_uri를 확인하세요:</p>
        <code>${getActualRedirectUri(req)}</code>
        <br><br>
        <a href="/">← 메인 페이지</a>
      </body>
      </html>
    `);
  }
});

// 스프레드시트에서 데이터 읽기
app.get('/api/google/read/:month', async (req, res) => {
  try {
    const data = await gsheets.readSheetData(req.params.month);
    res.json({ month: req.params.month, count: data.length, data });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 스프레드시트에 데이터 추가
app.post('/api/google/append', async (req, res) => {
  try {
    const { month, data } = req.body;
    console.log(`[append] 월: ${month}, 데이터: ${data?.length || 0}건`);
    if (data && data.length > 0) {
      console.log('[append] 첫 번째 항목:', JSON.stringify(data[0]));
    }
    const result = await gsheets.appendToSheet(month, data);
    console.log('[append] 결과:', JSON.stringify(result));
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('[append] 오류:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// 시트 목록 가져오기
app.get('/api/google/sheets', async (req, res) => {
  try {
    const sheets = await gsheets.getSheetList();
    res.json(sheets);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`가계부 서버 실행 중: http://localhost:${PORT}`);
  console.log(`import 폴더: ${IMPORT_DIR}`);
  gsheets.initOAuth2Client();
});
