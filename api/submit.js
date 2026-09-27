// 出差旅費報告表 → Ragic 轉送站（Vercel Function）
// 金鑰放在 Vercel 環境變數 RAGIC_API_KEY，絕不寫進程式（本 repo 是公開的）。
const RAGIC_URL = 'https://ap12.ragic.com/kumimi/travel-expense-system2/2';

// 主表
const F = {
  type: 1006047, company: 1006048, taxId: 1006049, traveler: 1006064, title: 1006055,
  dateStart: 1006052, dateEnd: 1006053, days: 1006054, visitTarget: 1006056,
  exRate: 1006059, exCur: 1006060, planItem: 1006058, ntdTotal: 1006063, fxTotal: 1006062,
  stampManager: 1006061, stampTraveler: 1006050,
};
// 國內費用明細
const D = { summary: 1006066, date: 1006067, city: 1006068, transport: 1006069, lodging: 1006070,
  meals: 1006071, parking: 1006072, other: 1006073, subtotal: 1006074 };
// 國內交通明細
const S = { date: 1006078, mode: 1006077, from: 1006079, to: 1006080, method: 1006081, km: 1006082,
  fuel: 1006083, etc: 1006084, fare: 1006085, amount: 1006086 };
// 國外費用明細（幣別／金額編號由老闆 9/27 親自點出）
const I = { date: 1006089, from: 1006088, dest: 1006090,
  transportCur: 1006091, transportAmt: 1006092, lodgingCur: 1006093, lodgingAmt: 1006094,
  mealsCur: 1006095, mealsAmt: 1006096, officeCur: 1006097, officeAmt: 1006098 };
// 報帳憑證（備註編號由老闆 9/27 點出）
const V = { serial: 1006101, kind: 1006102, count: 1006103, ticketCur: 1006104, ticketAmt: 1006105,
  advanceCur: 1006106, advanceAmt: 1006107, cardCur: 1006108, cardAmt: 1006109, note: 1006111 };

const TYPE_NAME = { domestic: '國內出差', intl: '國外出差' };
const STAMP_FIELD = { manager: F.stampManager, applicant: F.stampTraveler, traveler: F.stampTraveler };

const ymd = s => (s ? String(s).replace(/-/g, '/') : '');

function buildForm(p) {
  const fd = new FormData();
  const put = (id, val) => {
    if (id == null || val === '' || val == null) return;
    fd.append(String(id), String(val));
  };
  let rowId = 0;
  const row = (map, obj) => {
    rowId -= 1;
    for (const [k, id] of Object.entries(map)) {
      if (id == null || obj[k] === '' || obj[k] == null) continue;
      fd.append(`${id}_${rowId}`, String(k === 'date' ? ymd(obj[k]) : obj[k]));
    }
  };

  const c = p.common || {};
  put(F.type, TYPE_NAME[p.type]);
  put(F.company, c.companyName);
  put(F.taxId, c.taxId);
  put(F.traveler, c.filler);
  put(F.title, c.title);
  put(F.dateStart, ymd(c.dateStart));
  put(F.dateEnd, ymd(c.dateEnd));
  put(F.days, c.days);

  if (p.type === 'domestic') {
    put(F.visitTarget, c.visitTarget);
    put(F.ntdTotal, c.total);
    for (const r of p.rows || []) {
      row(D, r);
      for (const s of r.segments || []) row(S, { ...s, date: r.date });
    }
  } else {
    const x = p.extra || {};
    put(F.exRate, c.exRate);
    put(F.exCur, c.exCur);
    put(F.planItem, c.planItem);
    put(F.ntdTotal, x.ntdTotal);
    put(F.fxTotal, x.fxTotal);
    for (const r of p.rows || []) row(I, r);
    for (const v of x.vouchers || []) row(V, v);
  }

  for (const [key, st] of Object.entries(p.stamps || {})) {
    const id = STAMP_FIELD[key];
    if (!id || !st || !st.base64) continue;
    const ext = st.mimeType === 'image/png' ? 'png' : 'jpg';
    const name = `stamp_${key}.${ext}`;
    fd.append(String(id), new Blob([Buffer.from(st.base64, 'base64')], { type: st.mimeType }), name);
  }
  return fd;
}

async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: '只接受 POST' });
  const key = process.env.RAGIC_API_KEY;
  if (!key) return res.status(500).json({ ok: false, error: '伺服器尚未設定 Ragic 金鑰' });

  const p = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
  if (!p || !TYPE_NAME[p.type] || !p.common || !p.common.companyName) {
    return res.status(400).json({ ok: false, error: '資料不完整' });
  }

  try {
    const r = await fetch(`${RAGIC_URL}?api&v=3&doFormula=true&doDefaultValue=true`, {
      method: 'POST',
      headers: { Authorization: `Basic ${key}` },
      body: buildForm(p),
    });
    const data = await r.json().catch(() => ({}));
    if (!r.ok || data.status !== 'SUCCESS') {
      return res.status(502).json({ ok: false, error: 'Ragic 沒有收下資料', detail: data.msg || data.status || r.status });
    }
    return res.status(200).json({ ok: true, id: data.rv });
  } catch (e) {
    return res.status(502).json({ ok: false, error: '連不上 Ragic', detail: String(e) });
  }
}

module.exports = handler;
module.exports.buildForm = buildForm;
