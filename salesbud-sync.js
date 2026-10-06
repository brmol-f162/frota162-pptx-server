// salesbud-sync.js — Sincronização via API REST da Salesbud (OAuth2 client credentials).
//
// Consulta as reuniões concluídas e entrega cada uma ao MESMO pipeline do webhook
// (processarPayloadSalesbud). Serve como gatilho principal e como rede de segurança:
// a API lista toda reunião concluída, inclusive as que o webhook nunca disparou
// (caso Baterax: único convidado externo com Gmail, "Nenhuma conta externa encontrada").
//
// Pontos de projeto (vêm da documentação da API):
//  - Não existe sincronização incremental: uma reunião que vira "concluída" depois do ponto já
//    lido só aparece na próxima varredura. Por isso a janela é ROLANTE (48h) e a deduplicação é
//    feita por marcador, nunca por "desde a última execução".
//  - Loop de paginação sempre em has_more/next_cursor, nunca no tamanho de data (página curta
//    ou vazia com has_more=true é normal).
//  - Token expira em 1h e não há refresh token: renova-se chamando /oauth/token de novo.
//  - Limites: 120 req/min por credencial, 60/min em transcrição, 10/min em /oauth/token.
//
// Segurança de operação: NADA acontece enquanto SALESBUD_SYNC_DESDE não estiver definida
// (evita republicar no Slack tudo que o webhook já processou antes da virada).

const BASE = process.env.SALESBUD_API_BASE || 'https://api.salesbud.com.br';
const JANELA_MS = 48 * 60 * 60 * 1000;
const MAX_POR_CICLO = 3;                                   // evita rajada de chamadas à Anthropic
const PAUSA_ENTRE_MS = Number(process.env.SYNC_PAUSA_MS ?? 8000);
const MAX_TENTATIVAS_ERRO = 2;

// E-mail do dono da reunião na Salesbud -> executivo (mesmos 6 do webhook)
const EMAIL_PARA_EXECUTIVO = {
  'bruno@frota162.com.br': 'Bruno Pereira',
  'ravila@frota162.com.br': 'Rávila Silva',
  'julio@frota162.com.br': 'Julio Mazzetti',
  'thais.cristina@frota162.com.br': 'Thais Cristina',
  'palloma.santos@frota162.com.br': 'Palloma Santos',
  'william@frota162.com.br': 'William Duarte',
};

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// ── Título ────────────────────────────────────────────────────────────────
function decodificarTitulo(t) {
  return String(t || '').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&amp;/gi, '&').replace(/\s+/g, ' ').trim();
}
// Mesmo padrão do webhook: "Frota162 ><", "Frota162 <>", sem espaço, ou "Frota 162".
function tituloNoPadrao(t) {
  const l = decodificarTitulo(t).toLowerCase();
  return l.includes('frota162 ><') || l.includes('frota162 <>') || l.includes('frota162><') || l.includes('frota162<>') || l.includes('frota 162');
}
// Contém "frota162" mas fora do padrão (ex.: "Acompanhamento Genesis Group <> Frota162"): só é logado.
function tituloQuasePadrao(t) {
  const n = decodificarTitulo(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
  return n.includes('frota162') && !tituloNoPadrao(t);
}

// ── Autenticação ──────────────────────────────────────────────────────────
let tokenCache = { valor: null, expiraEm: 0 };

async function obterToken(forcar = false) {
  if (!forcar && tokenCache.valor && Date.now() < tokenCache.expiraEm) return tokenCache.valor;
  const id = process.env.SALESBUD_CLIENT_ID, secret = process.env.SALESBUD_CLIENT_SECRET;
  if (!id || !secret) throw new Error('SALESBUD_CLIENT_ID / SALESBUD_CLIENT_SECRET não configurados');

  // O endpoint exige application/x-www-form-urlencoded (JSON devolve 415).
  const r = await fetch(`${BASE}/oauth/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ grant_type: 'client_credentials', client_id: id, client_secret: secret }),
  });
  const texto = await r.text();
  let j = null; try { j = JSON.parse(texto); } catch (e) {}
  if (!r.ok || !j || !j.access_token) {
    throw new Error(`Salesbud /oauth/token ${r.status}: ${(j && (j.error_description || j.error)) || texto.slice(0, 150)}`);
  }
  // Guarda um pouco menos que expires_in; também renova ao receber 401.
  tokenCache = { valor: j.access_token, expiraEm: Date.now() + Math.max(60, (j.expires_in || 3600) - 120) * 1000 };
  return tokenCache.valor;
}

async function apiGet(caminho, params = {}) {
  const url = new URL(BASE + caminho);
  for (const [k, v] of Object.entries(params)) {
    if (v !== undefined && v !== null && v !== '') url.searchParams.set(k, String(v));
  }
  let renovou = false;
  for (let tentativa = 1; tentativa <= 4; tentativa++) {
    const token = await obterToken();
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });

    if (r.status === 401 && !renovou) { renovou = true; await obterToken(true); continue; }   // credencial revogada/expirada
    if (r.status === 429) {                                                                    // respeitar Retry-After
      const ra = Number(r.headers.get('retry-after')) || 5;
      await sleep(ra * 1000 + Math.random() * 500); continue;
    }
    if (r.status === 503) { await sleep(2000 * tentativa + Math.random() * 500); continue; }   // backoff com jitter

    if (!r.ok) {
      const corpo = await r.text();
      let codigo = '', reqId = '';
      try { const e = JSON.parse(corpo).error || {}; codigo = e.code || ''; reqId = e.request_id || ''; } catch (e) {}
      const err = new Error(`Salesbud ${caminho} ${r.status} ${codigo}${reqId ? ' (' + reqId + ')' : ''}`);
      err.status = r.status; err.code = codigo;
      throw err;
    }
    return r.json();
  }
  throw new Error(`Salesbud ${caminho}: tentativas esgotadas`);
}

// ── Listagem (pagina em has_more, nunca no tamanho de data) ───────────────
async function listarReunioes(desdeISO) {
  const todas = [];
  let cursor = null, paginas = 0;
  do {
    const pagina = await apiGet('/v1/meetings', { meeting_after: desdeISO, limit: 100, cursor });
    todas.push(...(pagina.data || []));
    const pg = pagina.pagination || {};
    cursor = pg.has_more && pg.next_cursor ? pg.next_cursor : null;
    paginas++;
  } while (cursor && paginas < 30);
  return todas;
}

// ── Conversão API -> formato do payload do webhook ────────────────────────
function montarTranscricao(tr) {
  return (tr.utterances || []).map(u => `${u.speaker ? u.speaker + ': ' : ''}${u.text}`).join('\n');
}

function montarPayload(m, tr, avaliacao, executivo) {
  const partic = (m.participants || []).map(p => {
    const contato = p.email || p.phone || '';
    return p.display_name ? `${p.display_name}${contato ? ' <' + contato + '>' : ''}` : contato;
  }).filter(Boolean);
  const conta = (m.accounts || [])[0];
  return {
    id: m.id,
    title: decodificarTitulo(m.title),
    status: 3,                         // a API só devolve reuniões concluídas
    isExternal: true,                  // a decisão de incluir é nossa: não filtramos por audience
    audienceReal: m.audience,
    meetingAt: m.meeting_at,
    transcription: montarTranscricao(tr),
    transcriptionIsPlain: true,
    customerName: partic.slice(0, 3).join(', '),
    company: conta && conta.domain ? conta.domain : '',
    context: {
      competitorMentions: (m.competitors || []).map(c => c.name).filter(Boolean),
      tags: (m.tags || []).map(t => t.name).filter(Boolean),
    },
    analytics: {
      score: avaliacao && avaliacao.score != null ? avaliacao.score : null,
      justification: avaliacao ? avaliacao.justification : null,
    },
    _origemApi: true,
    _executivo: executivo,
  };
}

// ── Rota ──────────────────────────────────────────────────────────────────
function registrarRotaSalesbudSync(app, deps) {
  const { processar, getDriveClient, isProcessed, isMarked, markGeneric, chaveReuniao } = deps;
  const memoria = new Set();        // reuniões já resolvidas neste processo (evita consultar o Drive à toa)
  const tentativas = new Map();     // callId -> falhas transitórias
  const avisadasForaPadrao = new Set();
  let rodando = false;

  async function executar({ dry, tituloFiltro }) {
    const resumo = { dry: !!dry, janela_desde: null, total_na_api: 0, a_processar: [], ignoradas: [], processadas: [] };

    const desdeEnv = process.env.SALESBUD_SYNC_DESDE ? new Date(process.env.SALESBUD_SYNC_DESDE) : null;
    if ((!desdeEnv || isNaN(desdeEnv.getTime())) && !tituloFiltro) {
      console.log('[SalesbudSync] SALESBUD_SYNC_DESDE não configurada — sincronização desligada (nada foi consultado).');
      resumo.erro = 'SALESBUD_SYNC_DESDE não configurada';
      return resumo;
    }
    let desde = new Date(Date.now() - JANELA_MS);
    if (!tituloFiltro && desdeEnv > desde) desde = desdeEnv;   // com ?titulo= a janela de 48h ignora o corte
    resumo.janela_desde = desde.toISOString();

    const reunioes = await listarReunioes(resumo.janela_desde);
    resumo.total_na_api = reunioes.length;
    const drive = getDriveClient();
    const candidatas = [];

    for (const m of reunioes) {
      const titulo = decodificarTitulo(m.title);
      if (tituloFiltro && !titulo.toLowerCase().includes(String(tituloFiltro).toLowerCase())) continue;

      if (!tituloNoPadrao(titulo)) {
        if (tituloQuasePadrao(titulo) && !avisadasForaPadrao.has(m.id)) {
          avisadasForaPadrao.add(m.id);
          console.log(`[SalesbudSync] TÍTULO FORA DO PADRÃO — id:${m.id} titulo:"${titulo}" dono:${m.owner && m.owner.email} (não será processada)`);
        }
        if (tituloQuasePadrao(titulo)) resumo.ignoradas.push({ id: m.id, title: titulo, motivo: 'titulo_fora_do_padrao' });
        continue;
      }
      const email = String((m.owner && m.owner.email) || '').toLowerCase();
      const executivo = EMAIL_PARA_EXECUTIVO[email];
      if (!executivo) { resumo.ignoradas.push({ id: m.id, title: titulo, motivo: 'dono_nao_mapeado', dono: email }); continue; }
      if (m.no_show) { resumo.ignoradas.push({ id: m.id, title: titulo, motivo: 'no_show' }); continue; }

      const callId = `sb_${m.id}`;
      const chave = chaveReuniao(titulo, m.meeting_at);
      if (memoria.has(chave)) continue;
      if (await isProcessed(drive, chave) || await isMarked(drive, `sbfinal_${callId}`)) { memoria.add(chave); continue; }
      if (!m.transcript || !m.transcript.available) {
        resumo.ignoradas.push({ id: m.id, title: titulo, motivo: 'transcricao_indisponivel_ainda' });
        continue;
      }
      candidatas.push({ m, titulo, executivo, callId, chave });
    }

    resumo.a_processar = candidatas.map(c => ({ id: c.m.id, title: c.titulo, dono: c.executivo, meeting_at: c.m.meeting_at, audience: c.m.audience }));
    console.log(`[SalesbudSync] ciclo — desde ${resumo.janela_desde}: ${resumo.total_na_api} reunião(ões) na API, ${candidatas.length} a processar${dry ? ' (dry-run)' : ''}`);

    // Freio: ?titulo= ignora o corte SALESBUD_SYNC_DESDE (serve para recuperar UMA reunião específica,
    // como a Baterax). Reuniões processadas por versões anteriores ao v29 não têm o marcador de
    // duplicidade, então um filtro amplo republicaria o que já saiu no Slack. Acima de 3 casamentos
    // não processa nada: o dry-run lista as reuniões para você escolher um título mais específico.
    const amplo = !!tituloFiltro && candidatas.length > 3;
    if (amplo) {
      resumo.erro = `?titulo=${tituloFiltro} casa com ${candidatas.length} reuniões — use um título mais específico. Nada foi processado.`;
      console.log(`[SalesbudSync] RECUSADO — ${resumo.erro}`);
    }
    if (dry || amplo) return resumo;

    let feitas = 0;
    for (const c of candidatas) {
      if (feitas >= MAX_POR_CICLO) break;
      let tr;
      try { tr = (await apiGet(`/v1/meetings/${c.m.id}/transcript`)).data; }
      catch (e) { console.log(`[SalesbudSync] falha ao ler transcrição de ${c.m.id}: ${e.message}`); continue; }
      if (!tr || !tr.available || !(tr.utterances || []).length) {
        console.log(`[SalesbudSync] transcrição de ${c.m.id} ainda indisponível (status:${tr && tr.status}) — tenta no próximo ciclo`);
        continue;
      }

      let avaliacao = null;
      try {
        const ev = (await apiGet(`/v1/meetings/${c.m.id}/evaluations/overall`)).data;
        if (ev && ev.status === 'completed') avaliacao = ev;
      } catch (e) { console.log(`[SalesbudSync] sem avaliação para ${c.m.id} (não bloqueante): ${e.message}`); }

      console.log(`[SalesbudSync] PROCESSANDO — id:${c.m.id} titulo:"${c.titulo}" dono:${c.executivo} audience:${c.m.audience} meeting_at:${c.m.meeting_at}`);
      const status = await processar(montarPayload(c.m, tr, avaliacao, c.executivo));
      feitas++;
      resumo.processadas.push({ id: c.m.id, title: c.titulo, status });

      if (status === 'ok') memoria.add(c.chave);
      else if (status === 'ja_processada') {
        // Só é definitivo se existe o marcador de "processada" (o webhook pode ter concluído).
        // Se não existe, era uma reserva de tentativa anterior/concorrente: tenta de novo depois.
        if (await isProcessed(drive, c.chave)) memoria.add(c.chave);
        else console.log(`[SalesbudSync] ${c.m.id} estava reservada por outra tentativa — será retomada no próximo ciclo`);
      } else if (status === 'descartada') {          // decisão final (ex.: empresa não identificada): não repetir a cada ciclo
        await markGeneric(drive, `sbfinal_${c.callId}`);
        memoria.add(c.chave);
      } else if (status === 'erro') {
        const n = (tentativas.get(c.callId) || 0) + 1;
        tentativas.set(c.callId, n);
        if (n >= MAX_TENTATIVAS_ERRO) {
          console.log(`[SalesbudSync] desistindo de ${c.m.id} após ${n} falhas — reprocesse manualmente`);
          await markGeneric(drive, `sbfinal_${c.callId}`);
          memoria.add(c.chave);
        }
      }
      if (feitas < MAX_POR_CICLO && PAUSA_ENTRE_MS > 0) await sleep(PAUSA_ENTRE_MS);
    }
    return resumo;
  }

  const handler = async (req, res) => {
    const chave = process.env.SYNC_KEY;
    if (chave && req.query.key !== chave) return res.status(403).json({ ok: false, erro: 'chave inválida' });

    const dry = req.query.dry === '1';
    const tituloFiltro = req.query.titulo ? String(req.query.titulo) : '';
    if (rodando) {
      console.log('[SalesbudSync] ciclo anterior ainda em andamento, pulando esta execução.');
      return res.json({ ok: true, pulado: true });
    }
    rodando = true;

    if (dry) {   // dry-run responde com o resultado (só lê a API e o Drive, não processa nada)
      try { return res.json({ ok: true, ...(await executar({ dry, tituloFiltro })) }); }
      catch (e) { return res.status(500).json({ ok: false, erro: e.message }); }
      finally { rodando = false; }
    }

    res.json({ ok: true, status: 'processing' });
    executar({ dry: false, tituloFiltro })
      .catch(e => console.error('[SalesbudSync] erro no ciclo:', e.message))
      .finally(() => { rodando = false; });
  };

  app.get('/cron/salesbud-sync', handler);
  app.post('/cron/salesbud-sync', handler);
}

module.exports = {
  registrarRotaSalesbudSync,
  // exportados para teste
  _t: { tituloNoPadrao, tituloQuasePadrao, montarPayload, montarTranscricao, listarReunioes, apiGet, obterToken, decodificarTitulo,
        reset: () => { tokenCache = { valor: null, expiraEm: 0 }; } },
};
