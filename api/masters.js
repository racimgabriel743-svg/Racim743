const { getMasters } = require('../lib/sheets_masters');

module.exports = async function handler(req, res) {
  if (req.method === 'OPTIONS') return res.status(200).end();
  try {
    const dados = await getMasters();
    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ ok: true, ...dados });
  } catch (err) {
    console.error('[API MASTERS] Erro:', err);
    return res.status(500).json({ ok: false, msg: err.message });
  }
};
