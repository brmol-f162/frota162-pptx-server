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
  'uf_de_atuacao',       // já existe no HubSpot — usado pros case studies regionais
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
//    corporativo > nada (o system prompt trata a ausência silenciosamente,
//    sem explicar o motivo — só não inventa)
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
// 2.5. Benchmark por faixa de placas + case studies regionais — via API de
// BUSCA do próprio HubSpot (não via SIA/Athena). Os campos que precisamos
// (uf_de_atuacao, hs_is_closed_won, data_de_churn, company_name,
// placas_totais_do_contrato, quantidade_de_multas_mensais) já são properties
// do próprio Deal — não precisa de infraestrutura nova nem de outro token,
// reusa o mesmo HUBSPOT_TOKEN e fica mais em tempo real que a SIA (que só
// espelha uma vez por dia).
// ----------------------------------------------------------------------------

// Filtro base reaproveitado nas duas buscas: cliente fechado ganho e sem
// data de churn preenchida (ativo).
const FILTRO_CLIENTE_ATIVO = [
  { propertyName: 'hs_is_closed_won', operator: 'EQ', value: 'true' },
  { propertyName: 'data_de_churn', operator: 'NOT_HAS_PROPERTY' },
];

async function hubspotSearchDeals(filters, properties, limit) {
  const resultado = await hubspotFetch('/crm/v3/objects/deals/search', {
    method: 'POST',
    body: JSON.stringify({
      filterGroups: [{ filters }],
      properties,
      limit: limit || 100,
    }),
  });
  return resultado.results || [];
}

// Faixas alinhadas com a tabela de preços já usada no resto da Frota162
function faixaDePlacas(placas) {
  const n = Number(placas);
  if (!n || n <= 0) return null;
  if (n <= 40) return [1, 40];
  if (n <= 99) return [41, 99];
  if (n <= 199) return [100, 199];
  if (n <= 299) return [200, 299];
  if (n <= 399) return [300, 399];
  if (n <= 499) return [400, 499];
  if (n <= 999) return [500, 999];
  return [1000, 999999];
}

async function buscarBenchmarkPlacas(placasTotais) {
  const faixa = faixaDePlacas(placasTotais);
  if (!faixa) return null;
  const [min, max] = faixa;

  const filters = [
    ...FILTRO_CLIENTE_ATIVO,
    { propertyName: 'placas_totais_do_contrato', operator: 'HAS_PROPERTY' },
    { propertyName: 'quantidade_de_multas_mensais', operator: 'HAS_PROPERTY' },
  ];
  const deals = await hubspotSearchDeals(
    filters,
    ['placas_totais_do_contrato', 'quantidade_de_multas_mensais', 'valor_infracoes_adm'],
    200
  );

  // Filtra pela faixa e calcula as médias no próprio código — mais simples e
  // confiável do que tentar acertar o operador BETWEEN da Search API.
  const naFaixa = deals
    .map(d => ({
      placas: Number(d.properties.placas_totais_do_contrato),
      multas: Number(d.properties.quantidade_de_multas_mensais),
      valor: Number(d.properties.valor_infracoes_adm),
    }))
    .filter(d => d.placas >= min && d.placas <= max && !isNaN(d.multas));

  if (naFaixa.length < 5) return null; // amostra pequena demais pra ser útil

  const mediaMultas = naFaixa.reduce((soma, d) => soma + d.multas, 0) / naFaixa.length;

  // Valor em R$ nem sempre está preenchido — calcula a média só com quem tem,
  // e só reporta se sobrar amostra suficiente pra não virar número solto.
  const comValor = naFaixa.filter(d => !isNaN(d.valor) && d.valor > 0);
  const mediaValor = comValor.length >= 5
    ? comValor.reduce((soma, d) => soma + d.valor, 0) / comValor.length
    : null;

  return {
    faixa: `${min}-${max === 999999 ? '1000+' : max} placas`,
    amostra: naFaixa.length,
    mediaMultasMes: mediaMultas.toFixed(1),
    mediaValorMultasMes: mediaValor ? mediaValor.toFixed(2) : null,
  };
}

async function buscarCaseStudiesRegionais(uf, nomeExcluir) {
  if (!uf) return []; // só busca se a UF do novo deal estiver preenchida

  const filters = [
    ...FILTRO_CLIENTE_ATIVO,
    { propertyName: 'uf_de_atuacao', operator: 'EQ', value: uf },
  ];
  const deals = await hubspotSearchDeals(filters, ['company_name'], 20);

  const nomes = [...new Set(deals.map(d => d.properties.company_name).filter(Boolean))];
  return nomes
    .filter(nome => nome.toLowerCase() !== String(nomeExcluir || '').toLowerCase())
    .slice(0, 5);
}

// ----------------------------------------------------------------------------
// 3. System prompt — regras de negócio + guardrails
// ----------------------------------------------------------------------------
const SYSTEM_PROMPT = `
Você monta o "Super Briefing" que o Executivo de Vendas da Frota162 lê ANTES
da call com o lead. O resultado vai direto num campo do Deal no HubSpot.

CONHECIMENTO DE BASE (use quando relevante, NUNCA explique o óbvio):
O Executivo já sabe o que é NIC, ANTT, CONTRAN e como funciona notificação de
multa. NÃO explique conceitos básicos. Use estes fatos regulatórios só quando
a estratégia realmente pedir, de forma direta, sem aula:
- Locadoras: a indicação de condutor é obrigação legal: atraso ou omissão
  transfere a responsabilidade financeira da multa e da NIC para a própria
  locadora (não para o condutor/cliente final).
- Regra ANTT/CONTRAN de reincidência: o valor da NIC por não indicar condutor
  é multiplicado pelo número de infrações iguais registradas em nome do
  mesmo CNPJ nos últimos 12 meses — quanto mais reincidência, maior o
  multiplicador. Use isso como argumento de urgência quando o perfil do
  cliente sugerir alto volume ou reincidência, não como explicação genérica.

COMPETIDORES CONHECIDOS (Beemon, Bluefleet, Broobot, Caça Multa, CertaDoc,
Click Multas, DR Multa, EasyGo, Infleet, LW, Monaco, NSTech, Sem Parar,
Smartec, Soluxlog, Ticket Log, Touc, Movic — "Solução própria" não conta como
concorrente real): se qualquer um desses nomes aparecer nas observações do
Pré-Vendas, monte um bloco DEDICADO "Concorrência" na Estratégia, com 2-3
bullets de como contornar objeções esperadas desse concorrente específico.
Base o contra-argumento SOMENTE nos diferenciais reais e verificáveis da
Frota162 (SNE, Enterprise 3 completo, Service/BPO, especialista dedicado,
preço transparente) — NUNCA invente ou afirme algo negativo não confirmado
sobre o concorrente. Se nenhum concorrente for mencionado, não crie esse
bloco (não force).

REGRAS INEGOCIÁVEIS:
- TAMANHO MÁXIMO: o texto final inteiro (as 3 seções somadas) não pode passar
  de 2.500 caracteres. O Executivo lê isso em pé, antes de entrar na call —
  não é um relatório, é um resumo tático. Se sobrar informação, corte a menos
  relevante; não afine a fonte nem espreme linhas — reduza o conteúdo mesmo.
- Bullets curtos (uma linha cada), nunca parágrafos corridos. Sem introdução,
  sem "vamos analisar", sem repetir o que já está óbvio no nome do campo.
- Cada seção tem um teto: "Observações do Pré-Vendas" resume em até 4-5
  bullets (não copia a observação inteira palavra por palavra se for longa —
  resume mantendo os fatos). "Contexto de Mercado" no máximo 2-3 pontos.
  "Estratégia para a Call" no máximo 4-5 bullets, só os mais relevantes pro
  perfil deste deal específico — não lista todas as regras do system prompt,
  só as que se aplicam.
- NUNCA cite nome de plano (Enterprise 1/2/3), preço ou percentual de
  desconto. O objetivo é dar direcionamento estratégico, não cotação.
- Toda notícia ou dado de mercado vem com o link da fonte ao lado, entre
  parênteses. Nenhuma afirmação de mercado sem link.
- Se não achar site, Instagram, domínio corporativo ou notícia relevante,
  simplesmente OMITA esse ponto — não escreva nada explicando que não achou.
  A ausência de uma seção ou bullet já fala por si. Nunca preencha com
  suposição genérica, dado inventado, ou frase tipo "não foi possível
  localizar" — isso só ocupa espaço sem ajudar o Executivo.
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
   - Se vier "Benchmark real Frota162" no contexto: use como simulação pro
     Executivo apresentar quando o CLIENTE não souber o próprio número de
     multas — sempre deixando explícito que é média de empresas atendidas
     na mesma faixa de placas, nunca apresentado como o dado exato dele.
   - Se vier "Clientes ativos da Frota162 na mesma UF": cite os nomes como
     prova social regional (ex: "já atendemos [empresas] na sua região"),
     só se fizer sentido no fluxo da estratégia — não force.

Escreva em português, direto, sem enrolação. O Executivo vai ler isso em
menos de 2 minutos antes de entrar na call. Prefira cortar informação a
estourar o limite de 2.500 caracteres.
`.trim();

// ----------------------------------------------------------------------------
// 4. Chamada à API da Anthropic (web_search + web_fetch na mesma request)
// ----------------------------------------------------------------------------
async function chamarClaudeSuperBriefing(dealData, fonteEmpresa) {
  const { props, observacoes } = dealData;

  const contextoEmpresa = fonteEmpresa.tipo !== 'nenhuma'
    ? `Fonte de empresa disponível (${fonteEmpresa.tipo}): ${fonteEmpresa.valor}. Use web_fetch nisso ANTES de gastar busca.`
    : 'Nenhum site, Instagram ou domínio de e-mail corporativo disponível para esta empresa — não gaste espaço mencionando essa ausência, apenas siga com o que houver de outras fontes.';

  // Benchmark por faixa de placas + case studies regionais — busca direto no
  // HubSpot (não SIA/Athena). Falha aqui é engolida — não derruba o resto do
  // briefing por causa de uma seção opcional.
  let contextoBenchmark = '';
  try {
    const benchmark = await buscarBenchmarkPlacas(props.placas_totais_do_contrato);
    if (benchmark) {
      const parteValor = benchmark.mediaValorMultasMes
        ? `, valor médio de multas R$${benchmark.mediaValorMultasMes}/mês`
        : '';
      contextoBenchmark += `\nBenchmark real Frota162 (clientes ativos, faixa ${benchmark.faixa}, amostra de ${benchmark.amostra} contas): média de ${benchmark.mediaMultasMes} multas/mês${parteValor}. Use isso SÓ como estimativa preliminar de simulação pro Executivo levar à call, deixando claro ao cliente que é média geral de empresas atendidas, não o número exato dele. NÃO mencione NIC aqui — não há dado histórico confiável para isso ainda.\n`;
    }
  } catch (e) {
    console.error('HUBSPOT_DEAL: falha ao buscar benchmark de placas (não bloqueante):', e.message);
  }

  try {
    const cases = await buscarCaseStudiesRegionais(props.uf_de_atuacao, props.dealname);
    if (cases.length > 0) {
      contextoBenchmark += `\nClientes ativos da Frota162 na mesma UF (${props.uf_de_atuacao}): ${cases.join(', ')}. Pode citar como prova social regional se fizer sentido na estratégia.\n`;
    }
  } catch (e) {
    console.error('HUBSPOT_DEAL: falha ao buscar case studies regionais (não bloqueante):', e.message);
  }

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
${contextoBenchmark}
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
      max_tokens: 4000,
      thinking: { type: 'disabled' }, // não precisamos de raciocínio estendido aqui —
      // sem isso, o Sonnet 5 gasta parte do max_tokens em "thinking" (modo adaptativo
      // é o padrão do modelo) antes mesmo de chegar no texto final, e o orçamento
      // pode acabar no meio do uso das ferramentas (foi o que causou o stop_reason
      // 'max_tokens' sem nenhum texto de resposta).
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
  const texto = (data.content || [])
    .filter(b => b.type === 'text')
    .map(b => b.text)
    .join('\n')
    .trim();

  // Não deixar gravar vazio em silêncio — se não veio texto final, o motivo
  // mais comum é max_tokens estourado no meio do uso de ferramentas
  // (stop_reason 'max_tokens'). Loga o suficiente pra diagnosticar sem
  // precisar adivinhar da próxima vez.
  if (!texto) {
    console.error(
      'HUBSPOT_DEAL: resposta da Anthropic sem texto final. stop_reason:',
      data.stop_reason,
      '| tipos de bloco recebidos:',
      (data.content || []).map(b => b.type).join(', ') || '(nenhum)'
    );
    throw new Error(`Resposta vazia da Anthropic (stop_reason: ${data.stop_reason})`);
  }

  // Rede de segurança: HubSpot aceita até 65.536 caracteres em multi-line
  // text. O prompt já pede até 2.500, isso aqui é só para o caso raro do
  // modelo estourar — corta com aviso em vez de deixar o PATCH falhar.
  const LIMITE_SEGURANCA = 60000;
  if (texto.length > LIMITE_SEGURANCA) {
    console.error(`HUBSPOT_DEAL: texto com ${texto.length} caracteres, truncando para ${LIMITE_SEGURANCA}`);
    return texto.slice(0, LIMITE_SEGURANCA) + '\n\n[...texto truncado — passou do limite de segurança]';
  }

  return texto;
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
        // Igual ao resto do server.js: por causa do express.text({type:'*/*'})
        // no topo do arquivo, req.body SEMPRE chega como string, nunca como
        // objeto já parseado — precisa fazer o JSON.parse manual aqui também.
        const rawBody = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
        let payload;
        try {
          payload = JSON.parse(rawBody);
        } catch (e) {
          console.log('HUBSPOT_DEAL: payload não é JSON válido', rawBody);
          return;
        }

        const dealId = payload?.objectId || payload?.dealId;
        if (!dealId) {
          console.log('HUBSPOT_DEAL: payload sem dealId', payload);
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
