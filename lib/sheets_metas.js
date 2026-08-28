// lib/sheets_metas.js
// Serviço de Metas & Atividades - usa a MESMA planilha base (GOOGLE_SHEETS_ID)
// já utilizada pelo login, para ficar interligado aos usuários existentes na aba "Usuarios".
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

function parsePrivateKey(raw) {
  let key = raw.replace(/^["']+|["']+$/g, '').trim();
  key = key.replace(/\\n/g, '\n');
  const BEGIN = '-----BEGIN PRIVATE KEY-----';
  const END = '-----END PRIVATE KEY-----';
  const bi = key.indexOf(BEGIN);
  const ei = key.indexOf(END);
  if (bi !== -1 && ei !== -1) {
    const body = key.substring(bi + BEGIN.length, ei).replace(/\s+/g, '');
    const lines = body.match(/.{1,64}/g) || [];
    key = `${BEGIN}\n${lines.join('\n')}\n${END}\n`;
  }
  return key;
}

const TIPOS_VALIDOS = ['hora', 'numero', 'valor'];

class SheetsMetasService {
  constructor() {
    this.doc = null;
    this.initialized = false;
  }

  async init() {
    if (this.initialized) return this.doc;
    try {
      const clientEmail = process.env.GOOGLE_SHEETS_CLIENT_EMAIL;
      const rawKey = process.env.GOOGLE_SHEETS_PRIVATE_KEY;
      const sheetId = process.env.GOOGLE_SHEETS_ID;

      if (!clientEmail) throw new Error('Variável GOOGLE_SHEETS_CLIENT_EMAIL não configurada');
      if (!rawKey) throw new Error('Variável GOOGLE_SHEETS_PRIVATE_KEY não configurada');
      if (!sheetId) throw new Error('Variável GOOGLE_SHEETS_ID não configurada');

      const privateKey = parsePrivateKey(rawKey);
      const serviceAccountAuth = new JWT({
        email: clientEmail,
        key: privateKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      this.doc = new GoogleSpreadsheet(sheetId, serviceAccountAuth);
      await this.doc.loadInfo();
      this.initialized = true;
      console.log(`[METAS] ✓ Conectado: ${this.doc.title}`);
      return this.doc;
    } catch (error) {
      console.error('[METAS] Erro na conexão:', error);
      throw new Error('Falha na conexão com a planilha de metas: ' + error.message);
    }
  }

  async getSheetMetas() {
    await this.init();
    let sheet = this.doc.sheetsByTitle['Metas'];
    if (!sheet) {
      console.log('[METAS] Criando aba Metas...');
      sheet = await this.doc.addSheet({
        title: 'Metas',
        headerValues: ['ID', 'Usuario', 'Atividade', 'Tipo', 'MetaTotal', 'Realizado', 'Status', 'Descricao', 'DataCriacao', 'AtribuidoPor']
      });
    }
    // Garante a coluna AtribuidoPor mesmo em planilhas criadas antes desta versão
    await sheet.loadHeaderRow();
    if (!sheet.headerValues.includes('AtribuidoPor')) {
      await sheet.setHeaderRow([...sheet.headerValues, 'AtribuidoPor']);
    }
    return sheet;
  }

  async getSheetLancamentos() {
    await this.init();
    let sheet = this.doc.sheetsByTitle['MetasLancamentos'];
    if (!sheet) {
      console.log('[METAS] Criando aba MetasLancamentos...');
      sheet = await this.doc.addSheet({
        title: 'MetasLancamentos',
        headerValues: ['DataHora', 'Usuario', 'MetaId', 'Atividade', 'Quantidade', 'Observacao']
      });
    }
    return sheet;
  }

  agoraFormatado() {
    return new Date().toLocaleString('pt-BR', {
      timeZone: 'America/Manaus',
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit', second: '2-digit'
    });
  }

  /**
   * Confirma que o usuário existe na aba "Usuarios" da base (mesma usada no login).
   */
  async usuarioExiste(usuario) {
    await this.init();
    const sheet = this.doc.sheetsByTitle['Usuarios'];
    if (!sheet) return false;
    const rows = await sheet.getRows();
    return rows.some(r => String(r.get('Usuario') || '').trim() === String(usuario || '').trim());
  }

  /**
   * Lista todos os usuários existentes na aba "Usuarios" (mesma base do login),
   * para que uma atividade possa ser atribuída a qualquer um deles.
   */
  async listarUsuarios() {
    try {
      await this.init();
      const sheet = this.doc.sheetsByTitle['Usuarios'];
      if (!sheet) return { ok: false, msg: 'Aba Usuarios não encontrada', dados: [] };

      const rows = await sheet.getRows();
      const usuarios = [...new Set(
        rows.map(r => String(r.get('Usuario') || '').trim()).filter(Boolean)
      )].sort((a, b) => a.localeCompare(b, 'pt-BR'));

      return { ok: true, dados: usuarios };
    } catch (error) {
      console.error('[METAS] Erro ao listar usuários:', error);
      return { ok: false, msg: 'Erro ao listar usuários: ' + error.message, dados: [] };
    }
  }

  formatarMeta(row) {
    const metaTotal = parseFloat(String(row.get('MetaTotal') || '0').replace(',', '.')) || 0;
    const realizado = parseFloat(String(row.get('Realizado') || '0').replace(',', '.')) || 0;
    const restante = Math.max(metaTotal - realizado, 0);
    return {
      id: String(row.get('ID') || ''),
      usuario: String(row.get('Usuario') || ''),
      atividade: String(row.get('Atividade') || ''),
      tipo: String(row.get('Tipo') || 'numero'),
      metaTotal,
      realizado,
      restante,
      percentual: metaTotal > 0 ? Math.min(Math.round((realizado / metaTotal) * 100), 100) : 0,
      status: restante <= 0 && metaTotal > 0 ? 'Concluída' : (String(row.get('Status') || 'Ativa')),
      descricao: String(row.get('Descricao') || ''),
      dataCriacao: String(row.get('DataCriacao') || ''),
      atribuidoPor: String(row.get('AtribuidoPor') || '')
    };
  }

  /**
   * Lista as metas/atividades de um usuário existente na base.
   */
  async listarMetas(usuario) {
    try {
      if (!usuario) return { ok: false, msg: 'Usuário não informado' };
      const existe = await this.usuarioExiste(usuario);
      if (!existe) return { ok: false, msg: 'Usuário não encontrado na base' };

      const sheet = await this.getSheetMetas();
      const rows = await sheet.getRows();
      const dados = rows
        .filter(r => String(r.get('Usuario') || '').trim() === String(usuario).trim())
        .map(r => this.formatarMeta(r));

      const resumo = {
        totalAtividades: dados.length,
        concluidas: dados.filter(d => d.status === 'Concluída').length,
        emAndamento: dados.filter(d => d.status !== 'Concluída').length,
        totalRealizado: dados.reduce((s, d) => s + d.realizado, 0),
        totalRestante: dados.reduce((s, d) => s + d.restante, 0)
      };

      return { ok: true, dados, resumo };
    } catch (error) {
      console.error('[METAS] Erro ao listar:', error);
      return { ok: false, msg: 'Erro ao listar metas: ' + error.message };
    }
  }

  /**
   * Cria uma nova atividade/meta e ATRIBUI a um usuário específico existente na base.
   * usuario: usuário responsável pela atividade (quem vai cumprir a meta)
   * atribuidoPor: usuário logado que está cadastrando/atribuindo a atividade
   * tipo: 'hora' | 'numero' | 'valor'
   */
  async criarMeta(usuario, atividade, tipo, metaTotal, descricao, atribuidoPor) {
    try {
      if (!usuario) return { ok: false, msg: 'Selecione o usuário responsável pela atividade' };
      if (!atividade) return { ok: false, msg: 'Nome da atividade é obrigatório' };
      if (!TIPOS_VALIDOS.includes(tipo)) return { ok: false, msg: 'Tipo inválido (use hora, numero ou valor)' };

      const meta = parseFloat(String(metaTotal).replace(',', '.'));
      if (!meta || meta <= 0) return { ok: false, msg: 'Meta deve ser maior que zero' };

      const existe = await this.usuarioExiste(usuario);
      if (!existe) return { ok: false, msg: 'Usuário responsável não encontrado na base' };

      const sheet = await this.getSheetMetas();
      const id = `M${Date.now()}${Math.floor(Math.random() * 1000)}`;

      await sheet.addRow({
        ID: id,
        Usuario: usuario,
        Atividade: atividade,
        Tipo: tipo,
        MetaTotal: meta,
        Realizado: 0,
        Status: 'Ativa',
        Descricao: descricao || '',
        DataCriacao: this.agoraFormatado(),
        AtribuidoPor: atribuidoPor || usuario
      });

      console.log(`[METAS] ✓ Meta criada para ${usuario} (por ${atribuidoPor || usuario}): ${atividade} (${tipo}) - ${meta}`);
      return { ok: true, msg: `Atividade atribuída a ${usuario} com sucesso!`, id };
    } catch (error) {
      console.error('[METAS] Erro ao criar meta:', error);
      return { ok: false, msg: 'Erro ao criar meta: ' + error.message };
    }
  }

  /**
   * Lança o que foi realizado no dia - subtrai da meta restante (soma em Realizado).
   */
  async lancarProgresso(usuario, id, quantidade, observacao) {
    try
    {
      if (!usuario || !id) return { ok: false, msg: 'Usuário e ID da atividade são obrigatórios' };
      const qtd = parseFloat(String(quantidade).replace(',', '.'));
      if (!qtd || qtd <= 0) return { ok: false, msg: 'Informe uma quantidade válida maior que zero' };

      const sheet = await this.getSheetMetas();
      const rows = await sheet.getRows();
      const row = rows.find(r => String(r.get('ID') || '') === String(id) && String(r.get('Usuario') || '').trim() === String(usuario).trim());

      if (!row) return { ok: false, msg: 'Atividade não encontrada para este usuário' };

      const metaTotal = parseFloat(String(row.get('MetaTotal') || '0').replace(',', '.')) || 0;
      const realizadoAtual = parseFloat(String(row.get('Realizado') || '0').replace(',', '.')) || 0;
      const novoRealizado = realizadoAtual + qtd;

      row.set('Realizado', novoRealizado);
      row.set('Status', novoRealizado >= metaTotal ? 'Concluída' : 'Ativa');
      await row.save();

      const lancamentos = await this.getSheetLancamentos();
      await lancamentos.addRow({
        DataHora: this.agoraFormatado(),
        Usuario: usuario,
        MetaId: id,
        Atividade: String(row.get('Atividade') || ''),
        Quantidade: qtd,
        Observacao: observacao || ''
      });

      console.log(`[METAS] ✓ Progresso lançado: ${usuario} - ${id} - +${qtd}`);
      return { ok: true, msg: 'Progresso lançado com sucesso!', meta: this.formatarMeta(row) };
    } catch (error) {
      console.error('[METAS] Erro ao lançar progresso:', error);
      return { ok: false, msg: 'Erro ao lançar progresso: ' + error.message };
    }
  }

  /**
   * Retorna o histórico de lançamentos (diários) de uma atividade específica.
   */
  async obterHistorico(usuario, id) {
    try {
      if (!usuario) return { ok: false, msg: 'Usuário não informado' };
      const sheet = await this.getSheetLancamentos();
      const rows = await sheet.getRows();
      let dados = rows
        .filter(r => String(r.get('Usuario') || '').trim() === String(usuario).trim())
        .map(r => ({
          dataHora: String(r.get('DataHora') || ''),
          metaId: String(r.get('MetaId') || ''),
          atividade: String(r.get('Atividade') || ''),
          quantidade: parseFloat(String(r.get('Quantidade') || '0').replace(',', '.')) || 0,
          observacao: String(r.get('Observacao') || '')
        }));

      if (id) dados = dados.filter(d => d.metaId === String(id));
      dados.reverse(); // mais recentes primeiro

      return { ok: true, dados };
    } catch (error) {
      console.error('[METAS] Erro ao obter histórico:', error);
      return { ok: false, msg: 'Erro ao obter histórico: ' + error.message };
    }
  }

  /**
   * Remove uma atividade/meta do usuário.
   */
  async excluirMeta(usuario, id) {
    try {
      if (!usuario || !id) return { ok: false, msg: 'Usuário e ID são obrigatórios' };
      const sheet = await this.getSheetMetas();
      const rows = await sheet.getRows();
      const row = rows.find(r => String(r.get('ID') || '') === String(id) && String(r.get('Usuario') || '').trim() === String(usuario).trim());

      if (!row) return { ok: false, msg: 'Atividade não encontrada para este usuário' };

      await row.delete();
      console.log(`[METAS] ✓ Meta excluída: ${usuario} - ${id}`);
      return { ok: true, msg: 'Atividade removida com sucesso!' };
    } catch (error) {
      console.error('[METAS] Erro ao excluir meta:', error);
      return { ok: false, msg: 'Erro ao excluir meta: ' + error.message };
    }
  }
}

module.exports = new SheetsMetasService();
