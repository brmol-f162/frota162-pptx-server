// followup-plan.js — Geração e envio da sugestão de follow-up pós-call.
//
// Módulo ISOLADO do pipeline principal (mesmo padrão do super-briefing.js):
// chamado depois que o PPTX estratégico já foi gerado e postado com sucesso.
// Qualquer falha aqui é só logada — nunca deve derrubar ou re-processar a call
// principal, que já foi entregue.
//
// Fonte primária de conteúdo: a TRANSCRIÇÃO da call (objeção literal, promessa
// feita, próximo passo combinado). Campos do HubSpot (motivo_da_dor,
// aderente_sne, frota_pj_ou_pf, frota_propria, origem, uf_de_atuacao,
// super_briefing) são contexto de apoio, nunca o gatilho principal.
// Temperatura e status de POC foram removidos por decisão do Bruno — não
// entram em nenhuma camada desta lógica.

const https = require('https');

// ─── Faixas reais de economia SNE por porte ──────────────────────────────
// Fonte: SIA (gold_roi_analitico), amostra de 500 linhas conta-mês entre
// mai/out/dez de 2025, 411 contas aderentes ao SNE. Percentis p25/p75 (banda
// "típica") e p90/máximo observado (banda "topo"). Usar SEMPRE estas faixas
// nos ganchos de case — nunca um valor fixo tipo "R$1 milhão".
// Revisar trimestralmente puxando nova amostra da SIA (gold_roi_analitico).
const FAIXA_ECONOMIA_SNE = {
  pequena: { max: 40,       mensal: 'R$78 a R$1.226',      anual: 'R$900 a R$14,7 mil' },
  media:   { max: 100,      mensal: 'R$300 a R$2.000',     anual: 'R$3,6 a R$24 mil' },
  grande:  { max: 300,      mensal: 'R$1.200 a R$3.350',   anual: 'R$14,7 a R$40 mil' },
  key:     { max: Infinity, mensal: 'R$3.350 a R$54.672+', anual: 'R$40 a R$650 mil+' },
};
function faixaPorPlacas(placas) {
  const p = Number(placas) || 0;
  if (p <= 40) return FAIXA_ECONOMIA_SNE.pequena;
  if (p <= 100) return FAIXA_ECONOMIA_SNE.media;
  if (p <= 300) return FAIXA_ECONOMIA_SNE.grande;
  return FAIXA_ECONOMIA_SNE.key;
}

// ─── Segmentação de cadência (porte x Origem N1) ─────────────────────────
// Racional: mediana de ciclo real por segmento (HubSpot, deals Ganhos
// jan/2025–set/2026) + benchmark de mercado B2B (8–12 toques / 2–4 semanas;
// inbound rápido 4–6 / 5–7 dias; outbound/enterprise 8–12+ / 3–4 semanas).
function segmentoCadencia(placas, origem) {
  const p = Number(placas) || 0;
  const isOutbound = String(origem || '').toLowerCase() === 'outbound';
  if (isOutbound) return { toques: '8 a 12', janela: '21 a 28 dias', multithread: true };
  if (p <= 40)  return { toques: '4 a 5', janela: '7 a 10 dias', multithread: false };
  if (p <= 100) return { toques: '5 a 6', janela: '10 a 12 dias', multithread: false };
  if (p <= 300) return { toques: '6 a 8', janela: '14 a 18 dias', multithread: false };
  return { toques: '8 a 10', janela: '21 a 25 dias', multithread: true };
}

// ─── Mapa executivo → webhook Slack do canal de follow-up ────────────────
// Um Incoming Webhook por canal (mesmo padrão já usado em SLACK_WEBHOOK_URL),
// nada de app/bot novo. Cada webhook precisa ser criado no canal
// correspondente e o valor colado na env var do Render.
const SLACK_FOLLOWUP_WEBHOOKS = {
  'palloma':        process.env.SLACK_FOLLOWUP_WEBHOOK_PALLOMA,
  'julio':          process.env.SLACK_FOLLOWUP_WEBHOOK_JULIO,
  'júlio':          process.env.SLACK_FOLLOWUP_WEBHOOK_JULIO,
  'ravila':         process.env.SLACK_FOLLOWUP_WEBHOOK_RAVILA,
  'rávila':         process.env.SLACK_FOLLOWUP_WEBHOOK_RAVILA,
  'thais':          process.env.SLACK_FOLLOWUP_WEBHOOK_THAIS,
  'william':        process.env.SLACK_FOLLOWUP_WEBHOOK_WILLIAM,
  'willîam':        process.env.SLACK_FOLLOWUP_WEBHOOK_WILLIAM,
  'bruno pereira':  process.env.SLACK_FOLLOWUP_WEBHOOK_BRUNOP,
  'bruno.pereira':  process.env.SLACK_FOLLOWUP_WEBHOOK_BRUNOP,
};
function getFollowupWebhook(nomeOuEmail) {
  if (!nomeOuEmail) return null;
  const lower = nomeOuEmail.toLowerCase();
  for (const [key, url] of Object.entries(SLACK_FOLLOWUP_WEBHOOKS)) {
    if (lower.includes(key) && url) return url;
  }
  return null;
}

// ─── Busca o Deal no HubSpot pelo nome da empresa (casamento por token) ──
// Nunca lança erro — qualquer falha resolve null e a régua segue só com a
// transcrição. Assume HUBSPOT_SERVICE_TOKEN = mesma Chave de Serviço usada
// no super-briefing.js (escopo crm.objects.deals.read). CONFIRMAR o nome
// exato da env var ao integrar, caso o super-briefing.js use outro nome.
function buscarDealHubSpot(nomeEmpresa) {
  return new Promise((resolve) => {
    if (!process.env.HUBSPOT_SERVICE_TOKEN || !nomeEmpresa) { resolve(null); return; }
    const body = JSON.stringify({
      filterGroups: [{
        filters: [
          { propertyName: 'dealname', operator: 'CONTAINS_TOKEN', value: nomeEmpresa },
          { propertyName: 'pipeline', operator: 'EQ', value: 'default' },
        ],
      }],
      properties: [
        'dealname', 'placas_totais_do_contrato', 'motivo_da_dor', 'aderente_sne',
        'frota_pj_ou_pf', 'frota_propria', 'origem', 'uf_de_atuacao', 'super_briefing',
      ],
      limit: 1,
    });
    const req = https.request({
      hostname: 'api.hubapi.com', path: '/crm/v3/objects/deals/search', method: 'POST',
      headers: {
        'authorization': `Bearer ${process.env.HUBSPOT_SERVICE_TOKEN}`,
        'content-type': 'application/json',
        'content-length': Buffer.byteLength(body),
      },
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const p = JSON.parse(data);
          resolve((p.results && p.results[0] && p.results[0].properties) || null);
        } catch (e) {
          console.log('[Followup] Falha ao parsear resposta HubSpot (não bloqueante):', e.message);
          resolve(null);
        }
      });
    });
    req.on('error', (e) => {
      console.log('[Followup] Erro ao buscar deal no HubSpot (não bloqueante):', e.message);
      resolve(null);
    });
    req.setTimeout(10000, () => { req.destroy(); resolve(null); });
    req.write(body); req.end();
  });
}

const SYSTEM_FOLLOWUP = `Você é especialista em follow-up de vendas B2B da Frota162 (gestão de frotas — multas, CNH, IPVA, SNE).

FONTE PRIMÁRIA: a transcrição da call. Extraia PRIMEIRO, literalmente:
- a objeção específica que travou (nas palavras do cliente, não uma categoria genérica)
- o que o Executivo prometeu enviar ou fazer
- o próximo passo combinado (com ou sem data)
- qualquer sinal de urgência, frieza ou entusiasmo dito explicitamente

Os campos de HubSpot fornecidos são CONTEXTO DE APOIO — preenchem o que a call não cobriu (motivo_da_dor formal, UF para case regional, se já é aderente ao SNE). Nunca inverta essa prioridade: a call manda, o HubSpot orienta.

NÃO USE, sob nenhuma hipótese: temperatura do deal, status de POC. Foram removidos e não devem influenciar toques, tom ou conteúdo.

REGRA DE CASE-STUDY: nunca invente nome de empresa. Se o Super Briefing fornecido já cita um case regional real, reaproveite o nome de lá. Se não houver Super Briefing e você tiver acesso a busca, procure 1 cliente real da Frota162 na mesma UF (site/notícia pública) — se não achar nada confiável, deixe case_nome vazio e use o placeholder "[Cliente na mesma região]" dentro das mensagens.

REGRA DE NÚMEROS: use SOMENTE a faixa de economia SNE fornecida no contexto (dado real da base). Nunca fabrique outro número.

MECANISMO DO TOQUE DE FECHAMENTO (breakup, sempre o último da régua): funciona por aversão à perda — precisa (a) fechar de fato ("vou encerrar/arquivar esse contato", nunca deixar em aberto pra sempre) e (b) quantificar o custo da inação com a faixa fornecida. Nunca terminar em "sem problema, quando quiser" sem esses dois elementos.

E-MAILS: a proposta já foi enviada por outro canal (champion letter) — NUNCA reproponha do zero nem repita a proposta inteira. Um e-mail de follow-up reforça UM ponto ligado à objeção da call; o e-mail de fechamento (breakup) só aparece no último toque, se a régua terminar em e-mail.

Retorne SOMENTE JSON válido, sem markdown, sem backticks, neste formato:
{"objecao_literal":"","promessa_executivo":"","proximo_passo_combinado":"","multithread_necessario":true,"case_nome":"","toques":[{"dia":"D+0","canal":"WhatsApp","objetivo":"1 frase","mensagem":"texto literal, pronto para copiar e enviar"}]}
O array "toques" deve ter a quantidade e cobrir a janela informadas em "Cadência a seguir". O ÚLTIMO item do array é sempre o toque de fechamento (breakup), seguindo o mecanismo acima. "multithread_necessario" só é true quando a cadência pedir multithread E o perfil do contato parecer influenciador (não decisor) pela transcrição.`;

function callClaudeFollowup({ transcricao, empresa, executivo, hubspotProps, cadencia, faixaEconomia }) {
  const superBriefing = (hubspotProps && hubspotProps.super_briefing) || '';
  const temSuperBriefing = superBriefing && superBriefing.trim().length > 20;

  const contextoHubspot = hubspotProps
    ? `\nCampos HubSpot (contexto de apoio, nunca prioridade sobre a transcrição):
- Motivo da dor: ${hubspotProps.motivo_da_dor || 'não informado'}
- Aderente SNE: ${hubspotProps.aderente_sne || 'não informado'}
- Frota PJ/PF: ${hubspotProps.frota_pj_ou_pf || 'não informado'}
- Tipo de frota: ${hubspotProps.frota_propria || 'não informado'}
- Origem N1: ${hubspotProps.origem || 'não informado'}
- UF de atuação: ${hubspotProps.uf_de_atuacao || 'não informado'}\n`
    : '\n(Deal não encontrado no HubSpot — siga só pela transcrição.)\n';

  const contextoSuperBriefing = temSuperBriefing
    ? `\nSuper Briefing já existente para este deal (reaproveite o case regional se houver um citado aqui):\n${superBriefing}\n`
    : '\n(Sem Super Briefing prévio para este deal — se usar busca para achar um case regional, deixe case_nome vazio e use o placeholder "[Cliente na mesma região]" caso não confirme um nome real.)\n';

  const conteudo = `Empresa: ${empresa}\nExecutivo: ${executivo}\n${contextoHubspot}${contextoSuperBriefing}
Cadência a seguir: ${cadencia.toques} toques em ${cadencia.janela}${cadencia.multithread ? ' — incluir pelo menos 1 toque de multithread (e-mail/LinkedIn para um contato de nível decisor, diferente do contato principal, se a transcrição sugerir que o contato principal é influenciador).' : '.'}
Faixa de economia SNE a usar nos ganchos de case (dado real — não fabricar outro número): mensal ${faixaEconomia.mensal} | anual ${faixaEconomia.anual}.

Transcrição da call:
${transcricao}`;

  // Sem Super Briefing prévio: habilita web_search para tentar achar 1 case
  // regional real (mesma UF) — mesmo princípio de fallback do super-briefing.js.
  const tools = temSuperBriefing ? undefined : [{ type: 'web_search_20250305', name: 'web_search' }];

  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 3000,
      system: SYSTEM_FOLLOWUP,
      messages: [{ role: 'user', content: conteudo }],
      ...(tools ? { tools } : {}),
    });
    const req = https.request({
      hostname: 'api.anthropic.com', path: '/v1/messages', method: 'POST',
      headers: {
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
        'content-length': Buffer.byteLength(body),
      },
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const p = JSON.parse(data);
          if (p.type === 'error' || !p.content || !Array.isArray(p.content)) {
            reject(new Error('Claude API error (followup): ' + (p.error?.message || p.error?.type || JSON.stringify(p).slice(0, 200))));
            return;
          }
          // Com web_search habilitado, a resposta pode ter blocos de
          // tool_use/tool_result intercalados — o JSON final é sempre o
          // ÚLTIMO bloco de texto retornado.
          const blocosTexto = p.content.filter(b => b.type === 'text');
          const ultimoTexto = blocosTexto[blocosTexto.length - 1];
          if (!ultimoTexto) { reject(new Error('Resposta da Claude sem bloco de texto final (followup)')); return; }
          const t = ultimoTexto.text.replace(/```json/gi, '').replace(/```/g, '').trim();
          resolve(JSON.parse(t));
        } catch (e) {
          reject(new Error('Falha ao parsear resposta da Claude (followup): ' + e.message));
        }
      });
    });
    req.on('error', reject);
    req.setTimeout(60000, () => { req.destroy(); reject(new Error('Timeout Claude (followup)')); });
    req.write(body); req.end();
  });
}

function formatarMensagemSlack(empresa, plano, cadencia) {
  const linhasToques = plano.toques.map((t, i) => {
    const marcaFinal = i === plano.toques.length - 1 ? ' _(fechamento)_' : '';
    const mensagemCitada = String(t.mensagem || '').split('\n').join('\n> ');
    return `*${t.dia} — ${t.canal}${marcaFinal}:* ${t.objetivo}\n> ${mensagemCitada}`;
  }).join('\n\n');

  const linhaMultithread = plano.multithread_necessario
    ? '\n\n:busts_in_silhouette: *Multithread necessário* — contato principal parece influenciador; pelo menos 1 toque acima deve ir para um decisor diferente.'
    : '';

  const linhaCase = plano.case_nome
    ? `\n- *Case a usar:* ${plano.case_nome}`
    : '\n- *Case a usar:* nenhum confirmado — usar o placeholder da mensagem ou buscar manualmente antes de enviar';

  return `:calling: *Sugestão de Follow-up* — ${empresa}\n\n- *Objeção da call:* ${plano.objecao_literal || 'não identificada'}\n- *Prometido na call:* ${plano.promessa_executivo || '—'}\n- *Próximo passo combinado:* ${plano.proximo_passo_combinado || '—'}\n- *Cadência:* ${cadencia.toques} toques em ${cadencia.janela}${linhaCase}${linhaMultithread}\n\n${linhasToques}`;
}

// ─── Ponto de entrada — chamar depois do markProcessed no server.js ──────
// Parâmetros:
//   empresa, executivo, transcricao — já disponíveis no handler do webhook
//   dCall — objeto retornado por callClaude() (análise do slide estratégico)
//   postSlack — a função postSlack(msg, webhookUrl) já existente no server.js
// Nunca lança erro — qualquer falha é logada e o pipeline principal segue.
async function gerarEEnviarFollowup({ empresa, executivo, transcricao, dCall, postSlack }) {
  try {
    const webhookUrl = getFollowupWebhook(executivo);
    if (!webhookUrl) {
      console.log(`[Followup] Sem canal Slack configurado para "${executivo}" — pulando (configure SLACK_FOLLOWUP_WEBHOOK_* no Render).`);
      return;
    }

    const hubspotProps = await buscarDealHubSpot(empresa);
    const placas = (hubspotProps && Number(hubspotProps.placas_totais_do_contrato)) || (dCall && dCall.placas) || 0;
    const origem = (hubspotProps && hubspotProps.origem) || '';
    const cadencia = segmentoCadencia(placas, origem);
    const faixaEconomia = faixaPorPlacas(placas);

    const plano = await callClaudeFollowup({ transcricao, empresa, executivo, hubspotProps, cadencia, faixaEconomia });
    const msg = formatarMensagemSlack(empresa, plano, cadencia);

    await postSlack(msg, webhookUrl);
    console.log(`[Followup] Enviado com sucesso — empresa:"${empresa}" executivo:"${executivo}"`);
  } catch (e) {
    console.error('[Followup] Falha ao gerar/enviar sugestão de follow-up (não bloqueante):', e.message);
  }
}

module.exports = { gerarEEnviarFollowup, faixaPorPlacas, segmentoCadencia, getFollowupWebhook };
