import express from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';
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
app.get('/api/parse/:filename', (req, res) => {
  try {
    const filePath = path.join(IMPORT_DIR, req.params.filename);
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: '파일을 찾을 수 없습니다.' });
    }

    const transactions = parseExcelFile(filePath);
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
app.get('/api/parse-all', (req, res) => {
  try {
    const files = fs.readdirSync(IMPORT_DIR)
      .filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));

    let allTransactions = [];

    for (const file of files) {
      try {
        const filePath = path.join(IMPORT_DIR, file);
        const transactions = parseExcelFile(filePath);
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

// Google 인증 상태 확인
app.get('/api/google/status', (req, res) => {
  const authenticated = gsheets.isAuthenticated();
  const authUrl = !authenticated ? gsheets.getAuthUrl() : null;
  res.json({ authenticated, authUrl });
});

// Google 인증 시작
app.get('/api/google/auth', (req, res) => {
  const authUrl = gsheets.getAuthUrl();
  if (authUrl) {
    res.redirect(authUrl);
  } else {
    res.status(500).json({ error: 'credentials.json 파일이 필요합니다.' });
  }
});

// Google OAuth 콜백
app.get('/api/google/callback', async (req, res) => {
  try {
    const code = req.query.code;
    await gsheets.getTokenFromCode(code);
    res.redirect('/?google=connected');
  } catch (e) {
    res.status(500).json({ error: e.message });
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
    const result = await gsheets.appendToSheet(month, data);
    res.json({ success: true, result });
  } catch (e) {
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`가계부 서버 실행 중: http://localhost:${PORT}`);
  console.log(`import 폴더: ${IMPORT_DIR}`);
  gsheets.initOAuth2Client();
});
