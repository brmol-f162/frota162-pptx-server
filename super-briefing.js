// ============================================================================
// SUPER BRIEFING — Frota162
// Gatilho: Workflow do HubSpot (Deal entra em "Reunião Agendada (AE)")
//          → ação "Enviar um webhook" → POST aqui.
//
// Este arquivo é pra ser INTEGRADO ao server.js já existente do
// frota162-pptx-server, não rodar sozinho: ele espera que `app`, `drive`,
// `claimCall` e `markProcessed` já existam nesse escopo (mesmo padrão usado
// no endpoint /webhook/salesbud).
//
// Variável de ambiente NOVA que precisa ser criada no Render:
//   HUBSPOT_TOKEN  → token de Chave de Serviço com escopos:
//                     crm.objects.deals.read, crm.objects.deals.write,
//                     crm.objects.contacts.read, crm.objects.contacts.write
//   (a API de Notes não tem escopo próprio — o acesso de leitura às Notes
//   do Deal é liberado pelos escopos de contacts.read/write, conforme a
//   doc oficial da Notes API. Não precisa de Private App legado.)
//
// ANTHROPIC_API_KEY já existe no ambiente (reaproveitado do pipeline Salesbud).
// ============================================================================

const HUBSPOT_TOKEN = process.env.HUBSPOT_TOKEN;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const HUBSPOT_BASE = 'https://api.hubapi.com';

// [Suposição] confirmar este model string no console da Anthropic antes de
// subir pra produção — nomes de modelo mudam com o tempo.
const CLAUDE_MODEL = 'claude-sonnet-5';

const MAX_BUSCAS = 4; // trava de custo combinada com o Bruno

// Property nova a ser criada no Deal (tipo: área de texto multilinha) —
// AJUSTAR o nome aqui se você criar com um nome interno diferente
const PROPERTY_SUPER_BRIEFING = 'super_briefing';

// Propriedades do Deal que entram no contexto do briefing
const DEAL_PROPERTIES = [
  'dealname',
  'autoridade_do_lead',
  'motivo_da_dor',
  'placas_totais_do_contrato',
  'aderente_sne',
  'origem',
  'sub_origem',
  'quantidade_de_multas_mensais',
  'frota_propria',       // própria / terceirizada / mista
  'frota_pj_ou_pf',      // PJ / PF / mista
  'site_da_empresa',
  'instagram_da_empresa'
];

// Domínios de e-mail pessoais — se o e-mail do contato NÃO for um desses,
// tratamos o domínio como candidato a site da empresa (fallback do ponto 7)
const DOMINIOS_EMAIL_PESSOAIS = [
  'gmail.com', 'hotmail.com', 'outlook.com', 'yahoo.com', 'yahoo.com.br',
  'icloud.com', 'live.com', 'msn.com', 'bol.com.br', 'uol.com.br',
  'terra.com.br', 'ig.com.br', 'r7.com', 'globo.com', 'protonmail.com'
];

// ----------------------------------------------------------------------------
// Helper genérico de chamada à API do HubSpot
// ----------------------------------------------------------------------------
async function hubspotFetch(path, opts = {}) {
  const res = await fetch(`${HUBSPOT_BASE}${path}`, {
    ...opts,
    headers: {
      'Authorization': `Bearer ${HUBSPOT_TOKEN}`,
      'Content-Type': 'application/json',
      ...(opts.headers || {})
    }
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`HubSpot ${res.status} em ${path}: ${body}`);
  }
  return res.json();
}

// ----------------------------------------------------------------------------
// 1. Buscar TUDO do Deal: properties + 100% das Notes associadas + e-mail do
//    contato (pro fallback de domínio do ponto 7)
// ----------------------------------------------------------------------------
async function buscarDealCompleto(dealId) {
  const deal = await hubspotFetch(
    `/crm/v3/objects/deals/${dealId}` +
    `?properties=${DEAL_PROPERTIES.join(',')}&associations=contacts,notes`
  );

  // Todas as Notes associadas ao Deal — não só uma property isolada
  const noteIds = (deal.associations?.notes?.results || []).map(r => r.id);
  let observacoes = [];
  if (noteIds.length > 0) {
    const batch = await hubspotFetch('/crm/v3/objects/notes/batch/read', {
      method: 'POST',
      body: JSON.stringify({
        inputs: noteIds.map(id => ({ id })),
        properties: ['hs_note_body', 'hs_timestamp']
      })
    });
    observacoes = (batch.results || [])
      .sort((a, b) => new Date(a.properties.hs_timestamp) - new Date(b.properties.hs_timestamp))
      .map(n => n.properties.hs_note_body)
      .filter(Boolean);
  }

  // E-mail do contato principal, pro fallback de domínio
  const contactId = deal.associations?.contacts?.results?.[0]?.id;
  let emailContato = null;
  if (contactId) {
    const contact = await hubspotFetch(`/crm/v3/objects/contacts/${contactId}?properties=email`);
    emailContato = contact.properties?.email || null;
  }

  return { dealId, props: deal.properties, observacoes, emailContato };
}

// ----------------------------------------------------------------------------
// 2. Resolver a fonte de empresa: site > Instagram > domínio do e-mail
//    corporativo > nada (registra a ausência, não inventa)
// ----------------------------------------------------------------------------
function resolverFonteEmpresa(props, emailContato) {
  if (props.site_da_empresa) return { tipo: 'site', valor: props.site_da_empresa };
  if (props.instagram_da_empresa) return { tipo: 'instagram', valor: props.instagram_da_empresa };

  if (emailContato) {
    const dominio = emailContato.split('@')[1]?.toLowerCase();
    if (dominio && !DOMINIOS_EMAIL_PESSOAIS.includes(dominio)) {
      return { tipo: 'dominio_email', valor: `https://${dominio}` };
    }
  }
  return { tipo: 'nenhuma', valor: null };
}

// ----------------------------------------------------------------------------
// 3. System prompt — regras de negócio + guardrails
// ----------------------------------------------------------------------------
const SYSTEM_PROMPT = `
Você monta o "Super Briefing" que o Executivo de Vendas da Frota162 lê ANTES
da call com o lead. O resultado vai direto num campo do Deal no HubSpot.

REGRAS INEGOCIÁVEIS:
- NUNCA cite nome de plano (Enterprise 1/2/3), preço ou percentual de
  desconto. O objetivo é dar direcionamento estratégico, não cotação.
- Toda notícia ou dado de mercado vem com o link da fonte ao lado, entre
  parênteses. Nenhuma afirmação de mercado sem link.
- Se não achar site, Instagram, domínio corporativo ou notícia relevante,
  escreva explicitamente "Não foi possível localizar [X]". Nunca preencha
  com suposição genérica ou dado inventado.
- Máximo de ${MAX_BUSCAS} buscas na web por briefing. Se já existe uma URL de
  site/Instagram/domínio de e-mail no contexto, use fetch direto nela em vez
  de gastar busca com isso.
- A maioria dos leads (frotas pequenas/médias) não tem imprensa própria —
  está tudo bem trazer só contexto de segmento (regulação de multas/CNH/
  ANTT/frotas) quando não existir notícia específica da empresa. Não force.

ESTRUTURA DO TEXTO, NESSA ORDEM:

1. "Observações do Pré-Vendas"
   Consolide o texto bruto recebido (pode vir de mais de uma Nota, inclusive
   de preenchimento de formulário) de forma organizada, sem reescrever o que
   já foi dito.

2. "Contexto de Mercado"
   Notícia de segmento e, quando existir, notícia específica da empresa —
   sempre com link. Registre explicitamente quando não achar nada.

3. "Estratégia para a Call"
   Pontos táticos para o Executivo, SEM citar plano ou preço:
   - aderente_sne = não/vazio → reforçar que aderir ao SNE aumenta a
     eficiência operacional da automação de robôs e reduz o valor pago por
     multa (sem citar percentual).
   - frota_pj_ou_pf = "Mista" → alertar que o mix de placas PF e PJ limita a
     varredura automática dos robôs e trava a integração com o SNE —
     antecipar isso como possível ponto de atrito, não só reagir se o
     cliente trouxer.
   - placas_totais_do_contrato baixo + indício de pouca estrutura de gestão
     própria → reforçar o diferencial de a Frota162 operar a rotina
     documental pelo cliente (sem citar nome de produto).
   - autoridade_do_lead indicando que o contato NÃO é decisor final E volume
     de placas alto → instruir o Executivo a mapear e trazer o decisor real
     antes de avançar qualquer proposta. Isso vem antes de qualquer
     argumento de valor.
   - Sempre ancorar o argumento em motivo_da_dor e, quando existir, na
     notícia de mercado encontrada no passo 2.
   - quantidade_de_multas_mensais disponível → pode citar como estimativa
     preliminar de economia, rotulada explicitamente "estimativa
     preliminar, a validar na call" — nunca como número fechado.

Escreva em português, direto, sem enrolação. O Executivo vai ler isso em
menos de 2 minutos antes de entrar na call.
`.trim();

// ----------------------------------------------------------------------------
// 4. Chamada à API da Anthropic (web_search + web_fetch na mesma request)
// ----------------------------------------------------------------------------
async function chamarClaudeSuperBriefing(dealData, fonteEmpresa) {
  const { props, observacoes } = dealData;

  const contextoEmpresa = fonteEmpresa.tipo !== 'nenhuma'
    ? `Fonte de empresa disponível (${fonteEmpresa.tipo}): ${fonteEmpresa.valor}. Use web_fetch nisso ANTES de gastar busca.`
    : 'Nenhum site, Instagram ou domínio de e-mail corporativo disponível para esta empresa — registre isso na Nota.';

  const userMsg = `
DEAL: ${props.dealname || '(sem nome)'}
Origem / Sub-origem: ${props.origem || '-'} / ${props.sub_origem || '-'}
Autoridade do lead: ${props.autoridade_do_lead || '-'}
Motivo da dor: ${props.motivo_da_dor || '-'}
Placas totais do contrato: ${props.placas_totais_do_contrato || '-'}
Frota própria/terceirizada/mista: ${props.frota_propria || '-'}
Frota PJ/PF/Mista: ${props.frota_pj_ou_pf || '-'}
Aderente ao SNE: ${props.aderente_sne || '-'}
Multas mensais: ${props.quantidade_de_multas_mensais || '-'}

${contextoEmpresa}

OBSERVAÇÕES ORIGINAIS DO PRÉ-VENDAS (ordem cronológica):
${observacoes.length
    ? observacoes.map((o, i) => `[Nota ${i + 1}]\n${o}`).join('\n\n')
    : '(nenhuma Nota encontrada associada a este Deal)'}

Monte o Super Briefing seguindo a estrutura definida no system prompt.
`.trim();

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
      'anthropic-beta': 'web-fetch-2025-09-10',
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 2000,
      system: SYSTEM_PROMPT,
      tools: [
        { type: 'web_search_20250305', name: 'web_search', max_uses: MAX_BUSCAS },
        { type: 'web_fetch_20250910', name: 'web_fetch', max_uses: 3 }
      ],
      messages: [{ role: 'user', content: userMsg }]
    })
  });

  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Anthropic API ${res.status}: ${body}`);
  }

  const data = await res.json();
  return (data.content || [])
    .filter(b => b.type === 'text')
    .map(b => b.text)
    .join('\n')
    .trim();
}

// ----------------------------------------------------------------------------
// 5. Gravar o briefing final na property do próprio Deal (não numa Nota —
//    Chaves de Serviço não têm escopo de Notes/Engagements ainda)
// ----------------------------------------------------------------------------
async function salvarBriefingNoDeal(dealId, corpoTexto) {
  return hubspotFetch(`/crm/v3/objects/deals/${dealId}`, {
    method: 'PATCH',
    body: JSON.stringify({
      properties: {
        [PROPERTY_SUPER_BRIEFING]: corpoTexto
      }
    })
  });
}

// ----------------------------------------------------------------------------
// 6. Rota — cola isto no server.js, dentro do escopo onde `app`, `drive`,
//    `claimCall` e `markProcessed` já existem
// ----------------------------------------------------------------------------
function registrarRotaSuperBriefing(app, getDriveClient, claimCall, markProcessed) {
  app.post('/webhook/hubspot-novo-deal', (req, res) => {
    res.sendStatus(200); // responde na hora, processa em background (mesmo padrão do /webhook/salesbud)

    (async () => {
      try {
        // Ajustar o nome do campo conforme o payload configurado no
        // "Enviar um webhook" do Workflow do HubSpot
        const dealId = req.body?.objectId || req.body?.dealId;
        if (!dealId) {
          console.log('HUBSPOT_DEAL: payload sem dealId', req.body);
          return;
        }

        console.log('HUBSPOT_DEAL RECEBIDO', dealId);

        const drive = getDriveClient(); // mesmo padrão do /webhook/salesbud: instancia por request

        const devoProcessar = await claimCall(drive, `hs_${dealId}`);
        if (!devoProcessar) {
          console.log('HUBSPOT_DEAL já processado', dealId);
          return;
        }

        const dealData = await buscarDealCompleto(dealId);
        const fonteEmpresa = resolverFonteEmpresa(dealData.props, dealData.emailContato);
        const briefing = await chamarClaudeSuperBriefing(dealData, fonteEmpresa);
        await salvarBriefingNoDeal(dealId, briefing);

        await markProcessed(drive, `hs_${dealId}`);
        console.log('HUBSPOT_DEAL SUCESSO', dealId);
      } catch (err) {
        console.error('HUBSPOT_DEAL Background error', err);
      }
    })();
  });
}

module.exports = { registrarRotaSuperBriefing };
