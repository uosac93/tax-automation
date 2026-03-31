/**
 * 기준시가 계산기 백엔드 서버
 * - 건축물대장 API 프록시 (API 키 서버 보관)
 * - 공시지가 조회
 * - 정적 파일 서빙
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const https = require('https');

const PORT = 3100;
const API_KEY = process.env.DATA_GO_KR_API_KEY || 'YOUR_DATA_GO_KR_API_KEY';
const BUILDING_API = 'http://apis.data.go.kr/1613000/BldRgstHubService';
const LAND_API = 'http://apis.data.go.kr/1160100/service/GetLandInfoService';

// ============================================================
// 유틸
// ============================================================

function fetchExternal(url) {
  return new Promise((resolve, reject) => {
    const mod = url.startsWith('https') ? https : http;
    mod.get(url, { timeout: 10000 }, (res) => {
      let body = '';
      res.on('data', (chunk) => body += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(body)); }
        catch { resolve(body); }
      });
    }).on('error', reject);
  });
}

function parseItems(data) {
  const items = data?.response?.body?.items?.item;
  if (!items) return [];
  return Array.isArray(items) ? items : [items];
}

function sendJson(res, obj, status = 200) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(obj));
}

function serveStatic(res, filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const mimeTypes = {
    '.html': 'text/html', '.js': 'application/javascript',
    '.css': 'text/css', '.json': 'application/json',
    '.xml': 'application/xml', '.png': 'image/png',
    '.ico': 'image/x-icon',
  };
  const contentType = mimeTypes[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not Found');
      return;
    }
    res.writeHead(200, { 'Content-Type': contentType + '; charset=utf-8' });
    res.end(data);
  });
}

// ============================================================
// API 핸들러
// ============================================================

async function handleSearch(query, res) {
  const { addr, sigunguCd, bjdongCd, bun, ji, ho, ledgerType, ledgerKind } = query;

  if (!sigunguCd && !addr) {
    return sendJson(res, { error: '주소 또는 시군구코드를 입력하세요.' }, 400);
  }

  const endpoint = ledgerKind === 'chongwal'
    ? '/getBrRecapTitleInfo'
    : '/getBrTitleInfo';

  const params = new URLSearchParams({
    serviceKey: API_KEY,
    sigunguCd: sigunguCd || '',
    bjdongCd: bjdongCd || '',
    numOfRows: '50',
    pageNo: '1',
    type: 'json',
  });
  if (bun) { params.set('platGbCd', '0'); params.set('bun', bun.padStart(4, '0')); }
  if (ji) params.set('ji', ji.padStart(4, '0'));

  try {
    const data = await fetchExternal(`${BUILDING_API}${endpoint}?${params.toString()}`);
    const items = parseItems(data);

    if (items.length === 0) {
      return sendJson(res, { error: '조회 결과가 없습니다.' });
    }

    // 호수 필터링
    let filtered = items;
    if (ho) {
      filtered = items.filter(i =>
        String(i.hoNm || '').includes(ho) || String(i.rnum || '') === ho
      );
    }

    const buildings = (filtered.length > 0 ? filtered : items).map(item => ({
      name: item.bldNm || '',
      address: item.platPlc || item.newPlatPlc || '',
      usage: item.mainPurpsCdNm || '',
      structure: item.strctCdNm || '',
      totArea: parseFloat(item.totArea) || 0,
      platArea: parseFloat(item.platArea) || 0,
      buildYear: parseInt(String(item.useAprDay || '').substring(0, 4)) || 0,
      flrCnt: parseInt(item.grndFlrCnt) || 0,
      ugrndFlrCnt: parseInt(item.ugrndFlrCnt) || 0,
      pnu: item.sigunguCd && item.bjdongCd
        ? `${item.sigunguCd}${item.bjdongCd}1${(item.bun || '0000').padStart(4,'0')}${(item.ji || '0000').padStart(4,'0')}`
        : '',
    }));

    sendJson(res, { type: 'building_list', buildings, totals: [] });
  } catch (e) {
    sendJson(res, { error: `API 조회 실패: ${e.message}` }, 500);
  }
}

async function handleLandprice(query, res) {
  const { pnu } = query;
  if (!pnu) return sendJson(res, { error: 'PNU가 필요합니다.' }, 400);

  const params = new URLSearchParams({
    serviceKey: API_KEY,
    pnu,
    numOfRows: '10',
    pageNo: '1',
    type: 'json',
    stdrYear: new Date().getFullYear().toString(),
  });

  try {
    const data = await fetchExternal(`${LAND_API}/getLandInfoItem?${params.toString()}`);
    const items = parseItems(data);
    sendJson(res, { indvdLandPrices: { field: items } });
  } catch (e) {
    sendJson(res, { error: `공시지가 조회 실패: ${e.message}` }, 500);
  }
}

// ============================================================
// 서버 시작
// ============================================================

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://localhost:${PORT}`);
  const pathname = url.pathname;

  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');

  if (req.method === 'HEAD' && pathname === '/') {
    res.writeHead(200);
    return res.end();
  }

  if (pathname === '/api/health') {
    return sendJson(res, { status: 'ok' });
  }

  if (pathname === '/api/search') {
    const query = Object.fromEntries(url.searchParams);
    return handleSearch(query, res);
  }


  if (pathname === '/api/parse-pdf' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY || '';
        if (!ANTHROPIC_KEY) return sendJson(res, { error: 'ANTHROPIC_API_KEY 환경변수를 설정하세요.' }, 500);
        const options = {
          hostname: 'api.anthropic.com',
          path: '/v1/messages',
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': ANTHROPIC_KEY,
            'anthropic-version': '2023-06-01',
          },
        };
        const proxyReq = https.request(options, (proxyRes) => {
          let data = '';
          proxyRes.on('data', chunk => data += chunk);
          proxyRes.on('end', () => sendJson(res, JSON.parse(data)));
        });
        proxyReq.on('error', e => sendJson(res, { error: e.message }, 500));
        proxyReq.write(body);
        proxyReq.end();
      } catch (e) { sendJson(res, { error: e.message }, 500); }
    });
    return;
  }

  if (pathname === '/api/landprice') {
    const query = Object.fromEntries(url.searchParams);
    return handleLandprice(query, res);
  }

  // 정적 파일 서빙
  let filePath = pathname === '/' ? '/index.html' : pathname;
  filePath = path.join(__dirname, filePath);
  serveStatic(res, filePath);
});

server.listen(PORT, () => {
  console.log(`서버 시작: http://localhost:${PORT}`);
});

module.exports = server;
