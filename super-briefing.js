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

// Estimativa de multas/NIC por fórmula fixa (validada pelo Bruno) — NÃO tira
// mais média do histórico de deals fechados: o campo quantidade_de_multas_mensais
// no HubSpot tem lixo de digitação (gente lançando total anual como se fosse
// mensal, etc.) e produzia número absurdo (ex: 111 multas/mês pra 8 placas).
// Fórmula: 0,8 multas/placa/mês, R$220 valor médio de multa, 30% delas viram
// NIC (valor dobrado).
const TAXA_MULTAS_POR_PLACA = 0.8;
const VALOR_MEDIO_MULTA = 220;
const TAXA_NIC = 0.30;

function calcularEstimativaMultas(placasTotais) {
  const n = Number(placasTotais);
  if (!n || n <= 0) return null;

  const multasMes = n * TAXA_MULTAS_POR_PLACA;
  const nicMes = multasMes * TAXA_NIC;
  const multasNormaisMes = multasMes - nicMes;
  const valorMes = (multasNormaisMes * VALOR_MEDIO_MULTA) + (nicMes * VALOR_MEDIO_MULTA * 2);

  return {
    multasMes: multasMes.toFixed(1),
    nicMes: nicMes.toFixed(1),
    valorMes: valorMes.toFixed(2),
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
- FORMATO DE SAÍDA: a resposta é SOMENTE o conteúdo final das 3 seções.
  NUNCA narre o processo — sem frases tipo "vou buscar", "encontrei",
  "agora vou montar o briefing", "nenhuma notícia foi encontrada", "sigo
  com contexto de segmento", ou qualquer variação que descreva o que você
  fez, achou ou não achou durante a pesquisa. Isso vale pra QUALQUER lugar
  do texto, não só a primeira linha — inclusive dentro da seção de mercado.
  ERRADO: "Nenhuma notícia específica sobre a [Empresa] foi encontrada —
  sigo com contexto de segmento." CERTO: se não achou notícia específica,
  a seção de mercado simplesmente começa direto com o contexto de segmento
  que você tiver, sem qualquer frase de transição sobre a ausência. A
  primeira linha da resposta é sempre o título da primeira seção ou o
  primeiro bullet — nunca uma frase sobre o processo de pesquisa.
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
   Reescreva em texto limpo e bem formatado — NÃO copie quebras de linha
   ruins ou frases cortadas no meio vindas da Nota original (fonte pode ter
   sido colada de outro lugar com formatação quebrada). Frases completas,
   bullets quando fizer sentido, nunca uma sentença partida em várias linhas
   soltas. Preserve os FATOS exatamente como estão, só conserte a forma.
   Corte linhas de checklist administrativo sem conteúdo de venda (ex:
   "Tripé Contato/Empresa/Negócio confirmado" ou variações de confirmação
   de checklist interno) — isso não ajuda o Executivo, é ruído de processo.

2. "Contexto de Mercado"
   Notícia de segmento e, quando existir, notícia específica da empresa —
   sempre com link. Se não achar nada relevante, omita a seção ou o ponto
   em silêncio (mesma regra do topo do prompt) — NUNCA registre a ausência.

3. "Estratégia para a Call"
   Pontos táticos para o Executivo, SEM citar plano ou preço:
   - aderente_sne = não/vazio → reforçar que aderir ao SNE aumenta a
     eficiência operacional da automação de robôs e reduz o valor pago por
     multa (sem citar percentual).
   - frota_pj_ou_pf = "Mista" → alertar que o mix de placas PF e PJ limita a
     varredura automática dos robôs e trava a integração com o SNE —
     antecipar isso como possível ponto de atrito, não só reagir se o
     cliente trouxer.
   - placas_totais_do_contrato MENOR OU IGUAL A 40 (regra dura, sem exceção
     e sem generalizar pra faixas maiores) → reforçar o diferencial de a
     Frota162 operar a rotina documental pelo cliente (sem citar nome de
     produto). ACIMA de 40 placas, NUNCA mencione isso como opção — não é
     elegível, ponto final, independente de outros sinais no deal.
   - autoridade_do_lead indicando que o contato NÃO é decisor final E volume
     de placas alto → instruir o Executivo a mapear e trazer o decisor real
     antes de avançar qualquer proposta. Isso vem antes de qualquer
     argumento de valor.
   - Sempre ancorar o argumento em motivo_da_dor e, quando existir, na
     notícia de mercado encontrada no passo 2.
   - quantidade_de_multas_mensais disponível → pode citar como estimativa
     preliminar de economia, rotulada explicitamente "estimativa
     preliminar, a validar na call" — nunca como número fechado.
   - Se vier "Estimativa Frota162" no contexto: use como simulação pro
     Executivo apresentar quando o CLIENTE não souber o próprio número de
     multas — sempre deixando explícito que é estimativa de mercado (não o
     dado exato dele), incluindo a parte de NIC quando fizer sentido.
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
// Rede de segurança contra narração de processo que passa pelo prompt.
// Remove linhas/frases que descrevem o que o modelo fez/achou/não achou
// durante a pesquisa, em vez de conteúdo do briefing em si. Case-insensitive,
// aplicado por sentença (separadas por ponto final ou quebra de linha) pra
// não jogar fora o resto de um parágrafo só por causa de uma frase solta.
const PADROES_NARRACAO = [
  /\b(vou|agora vou|preciso|deixe-?me)\s+(buscar|pesquisar|montar|verificar|checar|analisar)\b/i,
  /\bencontrei\b.{0,40}\b(dado|not[íi]cia|informa[çc][ãa]o|contexto)\b/i,
  /\bnenhuma\s+not[íi]cia\b/i,
  /\bn[ãa]o\s+(foi\s+poss[íi]vel|h[aá])\s+(encontrar|encontrada?s?|not[íi]cia)/i,
  /\bsigo\s+com\b/i,
  /\bvou\s+(compor|montar)\s+o\s+briefing\b/i,
  /\btrip[eé]\s+contato\/?empresa\/?neg[oó]cio\s+confirmad/i,
];

function removerNarracao(texto) {
  return texto
    .split('\n')
    .filter(linha => !PADROES_NARRACAO.some(re => re.test(linha)))
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// Quando uf_de_atuacao vem vazia do HubSpot, tenta inferir o estado via
// pesquisa rápida (site/Instagram/busca pelo nome), pra ainda conseguir
// rodar os case studies regionais. Chamada separada e barata (max_tokens
// baixo, poucas buscas) — só roda quando realmente falta a UF.
async function inferirUFViaPesquisa(props, fonteEmpresa) {
  const contextoEmpresa = fonteEmpresa.tipo !== 'nenhuma'
    ? `Fonte disponível (${fonteEmpresa.tipo}): ${fonteEmpresa.valor}.`
    : '';
  const userMsg = `
Empresa: ${props.dealname || '(sem nome)'}
${contextoEmpresa}
Segmento/observação: ${props.motivo_da_dor || '-'}

Pesquise rapidamente (site, Instagram, ou busca pelo nome da empresa) em qual
estado brasileiro (UF) essa empresa atua ou está sediada. Responda SOMENTE
com a sigla de 2 letras da UF (ex: SP, MG, RJ). Se não conseguir determinar
com razoável confiança, responda exatamente DESCONHECIDO. Não escreva mais
nada além disso — nem explicação, nem pontuação extra.
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
      max_tokens: 300,
      thinking: { type: 'disabled' },
      tools: [
        { type: 'web_search_20250305', name: 'web_search', max_uses: 2 },
        { type: 'web_fetch_20250910', name: 'web_fetch', max_uses: 1 }
      ],
      messages: [{ role: 'user', content: userMsg }]
    })
  });

  if (!res.ok) return null;
  const data = await res.json();
  const content = data.content || [];
  let textoFinal = [];
  for (let i = content.length - 1; i >= 0; i--) {
    if (content[i].type === 'text') textoFinal.unshift(content[i].text);
    else break;
  }
  const resposta = textoFinal.join('').trim().toUpperCase();
  const match = resposta.match(/\b([A-Z]{2})\b/);
  if (!match) return null;

  const UFS_VALIDAS = ['AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'];
  return UFS_VALIDAS.includes(match[1]) ? match[1] : null;
}

async function chamarClaudeSuperBriefing(dealData, fonteEmpresa) {
  const { props, observacoes } = dealData;

  const COMPETIDORES_CONHECIDOS = [
    'Beemon', 'Bluefleet', 'Broobot', 'Caça Multa', 'CertaDoc', 'Click Multas',
    'DR Multa', 'EasyGo', 'Infleet', 'LW', 'Monaco', 'NSTech', 'Sem Parar',
    'Smartec', 'Soluxlog', 'Ticket Log', 'Touc', 'Movic',
  ];
  const textoObservacoes = observacoes.join(' ');
  const competidoresMencionados = COMPETIDORES_CONHECIDOS.filter(c =>
    new RegExp(`\\b${c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i').test(textoObservacoes)
  );
  if (competidoresMencionados.length > 0) {
    console.log('HUBSPOT_DEAL: concorrente(s) detectado(s) nas observações:', competidoresMencionados.join(', '));
  } else {
    console.log('HUBSPOT_DEAL: nenhum concorrente conhecido mencionado nas observações');
  }

  const contextoEmpresa = fonteEmpresa.tipo !== 'nenhuma'
    ? `Fonte de empresa disponível (${fonteEmpresa.tipo}): ${fonteEmpresa.valor}. Use web_fetch nisso ANTES de gastar busca.`
    : 'Nenhum site, Instagram ou domínio de e-mail corporativo disponível para esta empresa — não gaste espaço mencionando essa ausência, apenas siga com o que houver de outras fontes.';

  // Benchmark por faixa de placas + case studies regionais — busca direto no
  // HubSpot (não SIA/Athena). Falha aqui é engolida — não derruba o resto do
  // briefing por causa de uma seção opcional.
  let contextoBenchmark = '';
  const estimativa = calcularEstimativaMultas(props.placas_totais_do_contrato);
  if (estimativa) {
    contextoBenchmark += `\nEstimativa Frota162 (fórmula: 0,8 multas/placa/mês, R$${VALOR_MEDIO_MULTA} valor médio, ${Math.round(TAXA_NIC * 100)}% viram NIC com valor dobrado) para ${props.placas_totais_do_contrato} placas: ~${estimativa.multasMes} multas/mês (das quais ~${estimativa.nicMes} NIC), ~R$${estimativa.valorMes}/mês em multas. Use isso SÓ quando o cliente não souber seus próprios números, como estimativa preliminar de simulação — deixando claro que é uma média de mercado, não o número exato dele.\n`;
  }

  try {
    let ufParaBusca = props.uf_de_atuacao;
    let ufInferida = false;

    if (!ufParaBusca) {
      console.log('HUBSPOT_DEAL: uf_de_atuacao vazia — tentando inferir via pesquisa');
      ufParaBusca = await inferirUFViaPesquisa(props, fonteEmpresa);
      if (ufParaBusca) {
        ufInferida = true;
        console.log('HUBSPOT_DEAL: UF inferida via pesquisa:', ufParaBusca);
      } else {
        console.log('HUBSPOT_DEAL: não foi possível inferir UF via pesquisa — case study regional não roda');
      }
    }

    if (ufParaBusca) {
      const cases = await buscarCaseStudiesRegionais(ufParaBusca, props.dealname);
      if (cases.length > 0) {
        console.log(`HUBSPOT_DEAL: ${cases.length} case study(ies) encontrado(s) na UF ${ufParaBusca}${ufInferida ? ' (inferida)' : ''}:`, cases.join(', '));
        const observacaoOrigem = ufInferida ? ', UF inferida via pesquisa — não confirmada no CRM' : '';
        contextoBenchmark += `\nClientes ativos da Frota162 na mesma UF (${ufParaBusca}${observacaoOrigem}): ${cases.join(', ')}. Pode citar como prova social regional se fizer sentido na estratégia.\n`;
      } else {
        console.log(`HUBSPOT_DEAL: UF ${ufParaBusca} (${ufInferida ? 'inferida' : 'do CRM'}), mas nenhum cliente ativo encontrado nela`);
      }
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
  const content = data.content || [];

  // FIX: não juntar todo bloco de texto da resposta. O modelo escreve texto
  // NORMALMENTE entre usos de ferramenta (ex: "vou buscar X", "encontrei Y")
  // — isso é comportamento padrão de tool-use, não desobediência ao prompt.
  // Pega só a sequência de blocos de texto do FINAL (depois do último uso de
  // ferramenta), que é a resposta de verdade; descarta qualquer texto que
  // veio antes disso.
  let textoFinal = [];
  for (let i = content.length - 1; i >= 0; i--) {
    if (content[i].type === 'text') {
      textoFinal.unshift(content[i].text);
    } else {
      break;
    }
  }
  const texto = textoFinal.join('\n').trim();
  const textoLimpo = removerNarracao(texto);

  // Não deixar gravar vazio em silêncio — se não veio texto final, o motivo
  // mais comum é max_tokens estourado no meio do uso de ferramentas
  // (stop_reason 'max_tokens'). Loga o suficiente pra diagnosticar sem
  // precisar adivinhar da próxima vez.
  if (!textoLimpo) {
    console.error(
      'HUBSPOT_DEAL: resposta da Anthropic sem texto final. stop_reason:',
      data.stop_reason,
      '| tipos de bloco recebidos:',
      (content || []).map(b => b.type).join(', ') || '(nenhum)'
    );
    throw new Error(`Resposta vazia da Anthropic (stop_reason: ${data.stop_reason})`);
  }

  // Rede de segurança: HubSpot aceita até 65.536 caracteres em multi-line
  // text. O prompt já pede até 2.500, isso aqui é só para o caso raro do
  // modelo estourar — corta com aviso em vez de deixar o PATCH falhar.
  const LIMITE_SEGURANCA = 60000;
  if (textoLimpo.length > LIMITE_SEGURANCA) {
    console.error(`HUBSPOT_DEAL: texto com ${textoLimpo.length} caracteres, truncando para ${LIMITE_SEGURANCA}`);
    return textoLimpo.slice(0, LIMITE_SEGURANCA) + '\n\n[...texto truncado — passou do limite de segurança]';
  }

  return textoLimpo;
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

// ----------------------------------------------------------------------------
// 6.5. Polling — alternativa ao Workflow do HubSpot (que exige Operations Hub
// Pro/Ent pra ação de webhook, fora do plano da Frota162). Roda via Render
// Cron Job, mesmo padrão do /cron/checklist-diario. Critério de "já
// processei" é o próprio campo super_briefing estar vazio — se uma tentativa
// falhar, o próximo ciclo tenta de novo sozinho, sem precisar de marcador
// manual como no fluxo por webhook.
// ----------------------------------------------------------------------------
function registrarRotaPolling(app, getDriveClient, claimCall, markProcessed) {
  const handler = (req, res) => {
    res.json({ ok: true, status: 'processing' });

    (async () => {
      try {
        const etapaId = process.env.HUBSPOT_ETAPA_REUNIAO_AGENDADA_ID;
        if (!etapaId) {
          console.error('HUBSPOT_DEAL_POLL: variável HUBSPOT_ETAPA_REUNIAO_AGENDADA_ID não configurada');
          return;
        }

        const filters = [
          { propertyName: 'dealstage', operator: 'EQ', value: etapaId },
          { propertyName: PROPERTY_SUPER_BRIEFING, operator: 'NOT_HAS_PROPERTY' },
        ];
        const deals = await hubspotSearchDeals(filters, ['dealname'], 20);

        if (deals.length === 0) {
          console.log('HUBSPOT_DEAL_POLL: nenhum deal novo pendente');
          return;
        }
        console.log(`HUBSPOT_DEAL_POLL: ${deals.length} deal(s) pendente(s) —`, deals.map(d => d.id).join(', '));

        const drive = getDriveClient();

        for (const deal of deals) {
          try {
            // Trava atômica — evita processar o mesmo deal duas vezes quando
            // dois ciclos do polling caem muito próximos (ex: dois monitores
            // de uptime, ou um ciclo que ainda não terminou de gravar quando
            // o próximo já rodou).
            const devoProcessar = await claimCall(drive, `hs_poll_${deal.id}`);
            if (!devoProcessar) {
              console.log('HUBSPOT_DEAL_POLL já sendo processado por outro ciclo, pulando', deal.id);
              continue;
            }

            const dealData = await buscarDealCompleto(deal.id);
            const fonteEmpresa = resolverFonteEmpresa(dealData.props, dealData.emailContato);
            const briefing = await chamarClaudeSuperBriefing(dealData, fonteEmpresa);
            await salvarBriefingNoDeal(deal.id, briefing);
            await markProcessed(drive, `hs_poll_${deal.id}`);
            console.log('HUBSPOT_DEAL_POLL SUCESSO', deal.id);
          } catch (err) {
            // Não quebra o lote inteiro por causa de um deal problemático —
            // como não chamamos markProcessed em caso de erro, esse deal
            // continua elegível e será tentado de novo no próximo ciclo.
            console.error('HUBSPOT_DEAL_POLL erro no deal', deal.id, err.message);
          }
        }
      } catch (err) {
        console.error('HUBSPOT_DEAL_POLL Background error', err);
      }
    })();
  };

  // GET e POST no mesmo handler — GET permite disparar isso de graça via
  // UptimeRobot (que só manda GET no plano free), sem precisar do Cron Job
  // pago do Render (US$1/mês, sem tier gratuito).
  app.get('/cron/verificar-deals-novos', handler);
  app.post('/cron/verificar-deals-novos', handler);
}

module.exports = { registrarRotaSuperBriefing, registrarRotaPolling };
