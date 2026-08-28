// api/metas.js
const sheetsMetas = require('../lib/sheets_metas');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const action = req.method === 'POST' ? req.body?.action : req.query?.action;

    if (!action) {
      return res.status(400).json({
        ok: false,
        msg: 'Action é obrigatória',
        acoesDisponiveis: ['listar', 'criar', 'lancar', 'historico', 'excluir']
      });
    }

    switch (action) {
      case 'listar': {
        const usuario = req.method === 'POST' ? req.body?.usuario : req.query?.usuario;
        const resultado = await sheetsMetas.listarMetas(usuario);
        return res.status(resultado.ok ? 200 : 400).json(resultado);
      }

      case 'criar': {
        const { usuario, atividade, tipo, metaTotal, descricao } = req.body || {};
        const resultado = await sheetsMetas.criarMeta(usuario, atividade, tipo, metaTotal, descricao);
        return res.status(resultado.ok ? 200 : 400).json(resultado);
      }

      case 'lancar': {
        const { usuario, id, quantidade, observacao } = req.body || {};
        const resultado = await sheetsMetas.lancarProgresso(usuario, id, quantidade, observacao);
        return res.status(resultado.ok ? 200 : 400).json(resultado);
      }

      case 'historico': {
        const usuario = req.method === 'POST' ? req.body?.usuario : req.query?.usuario;
        const id = req.method === 'POST' ? req.body?.id : req.query?.id;
        const resultado = await sheetsMetas.obterHistorico(usuario, id);
        return res.status(resultado.ok ? 200 : 400).json(resultado);
      }

      case 'excluir': {
        const { usuario, id } = req.body || {};
        const resultado = await sheetsMetas.excluirMeta(usuario, id);
        return res.status(resultado.ok ? 200 : 400).json(resultado);
      }

      default:
        return res.status(400).json({
          ok: false,
          msg: 'Ação inválida: ' + action,
          acoesDisponiveis: ['listar', 'criar', 'lancar', 'historico', 'excluir']
        });
    }
  } catch (error) {
    console.error('[API METAS] Erro:', error);
    return res.status(500).json({
      ok: false,
      msg: 'Erro interno: ' + error.message
    });
  }
};
