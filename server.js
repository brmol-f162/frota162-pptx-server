const crypto = require('crypto');
const express = require('express');
const PptxGenJS = require('pptxgenjs');
const { google } = require('googleapis');
const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { registrarRotaSuperBriefing, registrarRotaPolling } = require('./super-briefing');
const { gerarEEnviarFollowup } = require('./followup-plan');
const { registrarRotaSalesbudSync } = require('./salesbud-sync');

const app = express();
app.use(express.text({ type: '*/*', limit: '50mb' }));
app.use(express.json({ limit: '50mb' }));

// ─── Controle de duplicatas — MARCADOR ATÔMICO POR CALL ────────────── 
// Em vez de um único JSON (que sofre corrida de leitura/escrita quando várias
// calls chegam juntas), usamos UM arquivo marcador por call_id dentro de uma
// pasta de controle. Criar/checar um arquivo com nome único é atômico e à prova
// de paralelismo. Pasta de controle: env PROCESSED_FOLDER_ID (ou PASTA_RAIZ_ID).
function getDriveClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
    scopes: ['https://www.googleapis.com/auth/drive'],
  });
  return google.drive({ version: 'v3', auth });
}

function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

// Salva a transcrição completa em .txt no Drive e retorna o link — usado tanto
// como backup individual quanto como referência na planilha de histórico.
async function salvarTranscricaoDrive(drive, titulo, dataCallFormatada, executivo, callId, transcricao) {
  const dataPrefixo = (dataCallFormatada||'').slice(0,10) || 'sem-data';
  const nomeArq = `${dataPrefixo} - ${titulo}.txt`;
  const conteudo = `TITULO: ${titulo}\nDATA: ${dataCallFormatada}\nEXECUTIVO: ${executivo}\nCALL_ID: ${callId}\n\n${transcricao}`;
  const pastaId = process.env.PASTA_RAIZ_ID;

  const uploaded = await drive.files.create({
    supportsAllDrives: true,
    requestBody: { name: nomeArq, parents: [pastaId], mimeType: 'text/plain' },
    media: { mimeType: 'text/plain', body: conteudo },
    fields: 'id,webViewLink',
  });
  await drive.permissions.create({
    fileId: uploaded.data.id, supportsAllDrives: true,
    requestBody: { role: 'writer', type: 'anyone' },
  });
  return uploaded.data.webViewLink;
}

// Grava uma linha no histórico consultável (Google Sheets). Não bloqueia o
// pipeline se falhar (planilha não configurada, sem permissão, API desativada
// etc.) — só loga o erro e segue. Colunas fixas, nesta ordem:
// Data | Executivo | Empresa | Placas | Temperatura | ROI anual | Score Salesbud
// | Concorrentes | Tags | Perfil do lead | Próximo passo | Link PPTX | Link Transcrição | Call ID
// | MRR citado (coluna O, adicionada na v27)
async function salvarHistoricoPlanilha(linha) {
  const spreadsheetId = process.env.SPREADSHEET_ID;
  if (!spreadsheetId) {
    console.log('[Salesbud] SPREADSHEET_ID não configurado — pulando gravação no histórico');
    return;
  }
  const sheets = getSheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: 'A1',
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [linha] },
  });
}

// Dois tipos de marcador:
// - "claiming_" — temporário, só resolve corrida entre calls do MESMO lote (paralelismo).
//   Não significa sucesso; várias podem existir e sumir sem problema.
// - "processed_" — DEFINITIVO, só é criado depois que o Slack confirma o envio.
//   É o único que o isProcessed() consulta — se a call falhar antes do Slack,
//   nenhum "processed_" existe e ela pode ser tentada de novo no próximo lote.
function claimingName(callId) { return `claiming_${callId}.marker`; }
function processedName(callId) { return `processed_${callId}.marker`; }

async function listMarkersByName(drive, name) {
  try {
    const folderId = process.env.PROCESSED_FOLDER_ID || process.env.PASTA_RAIZ_ID;
    const res = await drive.files.list({
      q: `name='${name}' and '${folderId}' in parents and trashed=false`,
      corpora: 'allDrives',
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
      fields: 'files(id,createdTime)',
      spaces: 'drive',
      orderBy: 'createdTime',
    });
    return res.data.files || [];
  } catch(e) { console.log('listMarkersByName error:', e.message); return []; }
}

async function createMarkerFile(drive, name) {
  const folderId = process.env.PROCESSED_FOLDER_ID || process.env.PASTA_RAIZ_ID;
  const created = await drive.files.create({
    requestBody: { name, parents: [folderId], mimeType: 'text/plain' },
    media: { mimeType: 'text/plain', body: String(Date.now()) },
    supportsAllDrives: true,
    fields: 'id,createdTime',
  });
  return created.data.id;
}

// Já foi processada com SUCESSO (entregue no Slack)? Só isso pula a call.
async function isProcessed(drive, callId) {
  const marcadores = await listMarkersByName(drive, processedName(callId));
  return marcadores.length > 0;
}

// Marca sucesso DEFINITIVO — chamar SÓ depois do postSlack confirmar o envio.
async function markProcessed(drive, callId) {
  await createMarkerFile(drive, processedName(callId));
}

// Marcadores genéricos de "já fiz X uma vez" — usado para o aviso de transcrição
// vazia/curta não repetir a cada execução do Make para a mesma call.
async function isMarked(drive, chave) {
  const marcadores = await listMarkersByName(drive, `${chave}.marker`);
  return marcadores.length > 0;
}
async function markGeneric(drive, chave) {
  await createMarkerFile(drive, `${chave}.marker`);
}

// Resolve corrida de paralelismo com marcador TEMPORÁRIO (claiming).
// Retorna true se ESTA instância deve seguir processando, false se deve pular
// (ou porque já tem sucesso definitivo, ou porque perdeu a corrida do lote).
// IMPORTANTE: marcadores "claiming_" com mais de 5 minutos são considerados
// ÓRFÃOS (de uma tentativa anterior que travou/crashou antes do catch final)
// e são ignorados — senão um claiming órfão bloquearia a call PARA SEMPRE.
const CLAIMING_TTL_MS = 5 * 60 * 1000;

async function claimCall(drive, callId) {
  if (await isProcessed(drive, callId)) return false; // já teve sucesso antes

  const nome = claimingName(callId);
  const agora = Date.now();

  const existentesAntes = await listMarkersByName(drive, nome);
  const vivosAntes = existentesAntes.filter(m => (agora - new Date(m.createdTime).getTime()) < CLAIMING_TTL_MS);

  const meuId = await createMarkerFile(drive, nome);
  if (!meuId) return false;

  // Pequena espera para que criações concorrentes fiquem visíveis
  await new Promise(r => setTimeout(r, 1500));

  const todos = await listMarkersByName(drive, nome);
  const vivos = todos.filter(m => (agora - new Date(m.createdTime).getTime()) < CLAIMING_TTL_MS || m.id === meuId);

  if (vivos.length <= 1) return true; // só o meu (os órfãos foram ignorados)

  vivos.sort((a,b) => {
    if (a.createdTime !== b.createdTime) return a.createdTime < b.createdTime ? -1 : 1;
    return a.id < b.id ? -1 : 1;
  });
  const vencedor = vivos[0].id;
  return vencedor === meuId;
}
// Libera a reserva ("claiming_") de uma tentativa que FALHOU. Sem isso, o marcador temporário
// (TTL 5 min) faz a tentativa seguinte concluir que "outro processo está cuidando" e a
// reunião nunca é refeita. Best-effort: nunca lança erro.
async function liberarClaim(drive, callId) {
  try {
    const marcadores = await listMarkersByName(drive, claimingName(callId));
    for (const m of marcadores) await drive.files.delete({ fileId: m.id, supportsAllDrives: true });
  } catch (e) { console.log('liberarClaim (não bloqueante):', e.message); }
}

const COR = {
  laranja:'E8401C', dark:'1A1A1A', branco:'FFFFFF', fundo:'F7F6F4',
  divisor:'E0DFDD', verde:'1E6B1E', vermelho:'CC2200', azul:'1565C0', cinza:'888888',
};

// Mapa email/nome → Slack User ID
const SLACK_IDS = {
  'palloma': 'U09G3SJJXDX',
  'julio': 'U04MXBB585P',
  'júlio': 'U04MXBB585P',
  'ravila': 'U0951EZEQ69',
  'rávila': 'U0951EZEQ69',
  'thais': 'U0A3B5XV24S',
  'william': 'U05QVS86Z2N',
  'willîam': 'U05QVS86Z2N',
  'bruno pereira': 'U08FJSBCWAZ',
  'bruno.pereira': 'U08FJSBCWAZ',
};

function getSlackId(nomeOuEmail) {
  if (!nomeOuEmail) return null;
  const lower = nomeOuEmail.toLowerCase();
  for (const [key, id] of Object.entries(SLACK_IDS)) {
    if (lower.includes(key)) return id;
  }
  return null;
}

// ─── Matriz de oferta por volume de placas (regra vigente desde out/2026) ───
// 1-10   -> SOMENTE Service (o lead nem deve saber que existem planos)
// 11-40  -> Service preferencial; Enterprise avulso permitido
// 41-60  -> Service preferencial; Basic/Professional/Enterprise avulsos permitidos
// 61+    -> sem Service; planos conforme necessidade
// A decisão é feita aqui, em código, a partir de d.placas — não depende de o
// Claude "lembrar" a regra na hora de montar o JSON.
function bandaOferta(placas) {
  const p = Number(placas) || 0;
  if (p <= 0) return 'desconhecida';
  if (p <= 10) return 'service_somente';
  if (p <= 60) return 'service_preferencial';
  return 'planos';
}

// Texto do MRR para o Slack: só o que foi CITADO na call. Sem valor citado,
// diz explicitamente que não conseguiu mapear — nunca preenche com tabela/piso.
function textoMrr(d) {
  const v = d && d.mrr_citado ? String(d.mrr_citado).trim() : '';
  return v ? v : 'MRR: não consegui mapear (valor não citado na call)';
}

// Texto de placas para o Slack: parceria não tem frota a atender; placas 0 = não mapeado.
function textoPlacas(d) {
  if (d && d.tipo_reuniao === 'parceria') return 'não se aplica (reunião de parceria)';
  const p = Number(d && d.placas) || 0;
  return p > 0 ? `${p} placas` : 'placas não mapeadas';
}

const SYSTEM = `Você é especialista em vendas B2B da Frota162. Analise a transcrição e retorne SOMENTE JSON válido sem markdown sem backticks.

FIDELIDADE AOS DADOS DA CALL (vale acima de qualquer outra regra): placas e valores vêm do que foi dito na transcrição. O material é SEMPRE gerado, com tudo que a call trouxe; só o MRR exige honestidade total.
(1) Volume de placas: use o número dito na call, aceitando aproximações ditas ("uns 12 veículos"). Havendo vários números, use o da frota que a Frota162 atenderia — não confunda com a base de clientes de um parceiro, nem com a frota de um terceiro indicado. Se a call realmente não permitir determinar, placas=0 e gere o material mesmo assim (placas=0 segue a regra de 1 a 10 placas).
(2) Valores mensais (MRR): use somente o que o executivo CITOU na call. Se citou mais de um (ex.: plataforma e Service), registre todos em mrr_citado, dizendo a qual oferta cada um se refere. Só aritmética direta sobre números citados (ex.: placas x valor por placa citado). Se nenhum valor foi citado, mrr_citado="" e z3_investimento="A confirmar" — NUNCA preencher com valor de tabela, de exemplo ou de memória; a ausência deve ficar explícita.
(3) R$649 é o piso mensal do plano Enterprise até 40 placas. NÃO é o valor do Service nem de plano acima de 40 placas, e só pode aparecer se o executivo o citou.
(4) Valores que o próprio executivo disse de memória ("acho que", "pelo que me lembro") entram no material marcados como "a confirmar".

TIPO DE REUNIÃO (campo tipo_reuniao): "venda_direta" (padrão: empresa com frota própria avaliando contratar a Frota162) ou "parceria" (a outra empresa quer indicar clientes, revender, integrar via API ou trocar clientes, sem contratar a plataforma para uma frota própria). Em parceria: placas=0 e o material descreve a OPORTUNIDADE DE PARCERIA, não dor de frota — s1 cards = quem é o parceiro, tamanho das bases e modelo comercial citado; passos = como a parceria funciona; z3_stat = modelo da parceria (ex.: "API + indicação"); z3_investimento = a condição comercial citada (ex.: "10% por 12 meses"), sem inventar; campos service_* e z3_alt_* vazios; tem_roi=false e custo_mensal=0. mrr_citado começa com "MRR não mapeado: reunião de parceria, sem plano cotado" e depois lista os termos comerciais citados. Em parceria não se aplicam matriz de placas, planos nem Service. FIDELIDADE EM PARCERIA: (a) cada número pertence a quem o disse — nunca atribua ao parceiro uma métrica que o executivo citou sobre a própria Frota162 (ex.: o perfil de cliente de 40 a 100 placas é da Frota162, não da base do parceiro); (b) NÃO afirme o que o parceiro não faz, não oferece ou não tem hoje, a menos que ele tenha dito isso; na dúvida, escreva "a confirmar"; (c) eventos só como "convite" ou "presença conjunta" se ambos confirmaram; (d) quem fará cada próximo passo deve seguir o que a transcrição indica (ex.: o executivo falar com o próprio diretor não é ação do parceiro).

PLANOS (nomes atuais): Basic (antigo Enterprise 1) = notificações + multas + SNE + 1 CNPJ NTT. Professional (antigo Enterprise 2) = Basic + IPVA/licenciamento + indicação de condutor + 3 CNPJs. Enterprise (antigo Enterprise 3) = Professional + consulta de CNH + toxicológico + 5 CNPJs. Se a transcrição usar os nomes antigos, converta para os atuais. Recomendação por necessidade (para planos avulsos acima de 40 placas): consulta de CNH = sim -> Enterprise; indicação de condutor e/ou IPVA/licenciamento = sim (sem CNH) -> Professional; nenhum dos três -> Basic. Preços por placa NÃO ficam neste prompt: use apenas os citados na call.

MATRIZ DE OFERTA POR VOLUME DE PLACAS (somente venda_direta; regra dura, vigente desde out/2026):
- 1 a 10 placas: SOMENTE Service. O lead nem deve saber que existem planos: o material NÃO menciona Basic, Professional, Enterprise, a palavra "plano", a opção de operar a plataforma sozinho nem valor de plataforma avulsa. Oferta única: "Service Frota162" com UM valor mensal total em z3_investimento (mesmo que o executivo tenha citado plataforma e adicional em separado, exiba só o total). Deixe vazios service_sem_titulo, service_sem_itens, z3_alt_label, z3_alt_investimento e z3_alt_tagline. service_header_sub e service_com_titulo não citam plano.
- 11 a 40 placas: preferencialmente Service; o plano Enterprise avulso pode ser vendido (único plano permitido nesta faixa — nunca Basic ou Professional).
- 41 a 60 placas: preferencialmente Service; Basic, Professional e Enterprise avulsos podem ser vendidos.
- Acima de 60 placas: sem Service. Deixe vazios todos os campos service_* e z3_alt_*; z3_stat é o plano recomendado.
Para 11 a 60 placas: z3_stat="Service Frota162" e z3_investimento = valor total do Service citado na call; a opção avulsa vai na linha secundária: z3_alt_label="Ou só a plataforma - você opera (NomeDoPlano)" e z3_alt_investimento = valor avulso citado (ou "A confirmar" se não foi citado). z3_alt_tagline SÓ se os DOIS valores foram citados: "Menos gente, tempo e risco - por só R$X a mais", com X = valor do Service menos o da plataforma.
Se o executivo ofereceu na call algo FORA desta matriz (ex.: plano avulso para até 10 placas; Basic ou Professional para até 40; Service acima de 60), o material segue a matriz e alerta_regra_placas descreve o desvio em uma frase curta (vazio se não houve desvio).

SERVICE — posicionamento (do treinamento oficial): NÃO é terceirizar para despachante — é BPO documental, a Frota162 assume o processo com especialistas dedicados. Nunca revelar limites internos de quantidade ao cliente (ex: "144 indicações/ano") — dizer apenas "está incluso". Sempre mencionar SNE (desconto por adesão) como pilar central do Service. Benefícios reais a explorar: zero novas contratações, zero curva de aprendizado da equipe do cliente, redução de risco (prazo/indicação/CNH não dependem mais da memória do cliente), tempo da equipe liberado para a operação. Os campos service_* ficam SEM valores em R$ (a revelação numérica fica no slide de solução). Preencha os campos service_* para 1 a 60 placas.

ROI: economia_multas=multas_mes x valor x 0.4(SNE) ou 0.2. economia_NIC=NIC_tratadas x valor x 0.6. economia_pessoas=(func-1) x 2500. ROI_anual=economia_total_anual - investimento_anual. payback = investimento_mensal / (ROI_anual / 12). investimento_mensal_num = valor mensal da oferta principal CITADO na call (0 se não citado). SOMENTE calcular dias_payback quando tem_roi=true e ROI_anual > 0; senão dias_payback=0.
tem_roi e custo_mensal: SÓ preencher quando o volume de multas por mês foi informado ou confirmado na call (valor da multa: o citado, ou R$130 como mínimo quando não citado). Caso contrário tem_roi=false, custo_mensal=0, roi_anual=0, dias_payback=0 — nesse caso o slide "custo de esperar" NÃO é gerado. NUNCA fabricar números de ROI/economia que a call não confirmou.

LINGUAGEM: material apresentado pelo executivo Frota162 à diretoria do cliente. Use linguagem voltada ao cliente: "sua frota", "seu time", "sua operação". NÃO linguagem interna da Frota162.

REGRAS: valor mínimo de multa R$130 (dizer "mínimo", nunca "médio"). sinal menos SÓ nas barras do custo de esperar. títulos máx 40 chars. z3_stat é "Service Frota162" (1 a 60 placas), ou o plano recomendado (Basic, Professional ou Enterprise) acima de 60 placas, ou "ROI anual R$X" quando tem_roi=true — NUNCA nome inventado como "Plano Intermediário". Em s2_header_bold, s2_header_normal, z3_sub1 e z3_badge para 1 a 10 placas: nenhuma menção a plano. headers provocativos e específicos para ESTE cliente. NUNCA usar emojis em nenhum campo. z1_stat1 e z1_stat2 devem ser curtos max 12 chars ex: 'R$90k' '30%' '5.000+' nunca frases longas.

TEMPERATURA: quente=lead engajado perguntas próximos passos decisor envolvido. morno=interesse sem comprometimento claro. frio=pouco engajamento objeções sem próximo passo.

JSON (todos obrigatórios; campos service_* e z3_alt_* só quando a matriz permitir):
{"empresa":"","perfil_lead":"decisor ou influenciador","tipo_reuniao":"venda_direta ou parceria","placas":0,"cnpjs":0,"segmento":"","tem_roi":false,"mrr_citado":"valores mensais CITADOS na call com a oferta de cada um, ex: 'R$815,00/mês (Service) | R$649,00/mês (plataforma avulsa)'; vazio se nenhum valor foi citado","alerta_regra_placas":"desvio da matriz de placas em 1 frase, ou vazio","temperatura":"quente ou morno ou frio","roi_anual":0,"s1_header_bold":"[Nome do decisor se identificado],\\nvocês têm X placas [situação específica]. Formato: Nome,\\nvocês têm 22 placas rodando SP sem visibilidade. Se sem nome: frase provocativa com dado real max 70 chars","s1_header_sub":"X placas · Y CNPJs · Região","s1_subtitulo":"contexto segmento voltado ao cliente","cards":[{"stat":"","titulo":"max 40 chars","desc":"2-3 linhas específicas voltadas ao cliente"},{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""}],"s1_footer_bold":"urgência específica com número real max 80 chars","s1_footer_normal":"complemento","service_header_bold":"frase conceitual SEM valores em R$, ex: 'A Frota162 também pode operar tudo isso para você.'","service_header_sub":"1 a 10 placas: sem citar plano (ex: 'Operação completa pela Frota162'); 11 a 60: pode citar 'Enterprise + Service'","service_sem_titulo":"SEM SERVICE - você opera (vazio para 1 a 10 placas)","service_sem_itens":["item 1 sem service","item 2","item 3","item 4"],"service_com_titulo":"COM SERVICE - a Frota162 opera (para 1 a 10 placas: 'O QUE A FROTA162 OPERA POR VOCE')","service_com_itens":["item 1 com service orientado a resultado","item 2","item 3","item 4"],"service_beneficios":[{"stat":"0","titulo":"Novas contratações","desc":"curto"},{"stat":"Menos","titulo":"Tempo da equipe","desc":"curto"},{"stat":"100%","titulo":"Do risco sai da mão","desc":"curto"}],"service_nota":"reforço qualitativo sem valores em R$","s2_header_bold":"Da dor de -R$X ao retorno de +R$Y por ano. OU X placas sem visibilidade, a recomendação é o Service Frota162 / o plano [nome].","s2_header_normal":"Como a Frota162 resolve, em 3 passos — [Empresa]","z1_stat1":"","z1_sub1":"1 linha","z1_stat2":"","z1_sub2":"1 linha","z1_bullets":["dado específico 1","dado específico 2"],"passos":[{"titulo":"max 35 chars voltado ao cliente","desc":"1 linha no contexto do cliente"},{"titulo":"","desc":""},{"titulo":"","desc":""},{"titulo":"","desc":""}],"z3_stat":"Service Frota162 OU Basic OU Professional OU Enterprise OU ROI anual R$X OU o modelo da parceria (só em parceria)","z3_sub1":"retorno por ano ou o que está incluso","z3_investimento":"valor mensal CITADO na call da oferta principal, ou 'A confirmar'","z3_badge":"diferencial específico para este cliente","z3_alt_label":"só 11 a 60 placas: 'Ou só a plataforma - você opera (NomeDoPlano)'","z3_alt_investimento":"só 11 a 60 placas: valor avulso citado, ou 'A confirmar'","z3_alt_tagline":"só 11 a 60 placas e só se os dois valores foram citados: 'Menos gente, tempo e risco - por só R$X a mais'","z3_nota":"condições comerciais e limitações reais","s2_cta_bold":"próximo passo combinado na call","s2_cta_normal":"ação concreta","custo_mensal":0,"investimento_mensal_num":0,"dias_payback":0,"s3_header_bold":"Cada mês sem a Frota162 é R$X saindo do caixa da [Empresa].","s3_header_sub":"base confirmada na call","s3_formula":"fórmula usada","s3_nota":"metodologia e limitações","slack_resumo":"2 linhas objetivas SEM EMOJIS: dor principal + alerta crucial","proximo_passo":"ação concreta para o executivo","concorrencia_detalhe":"SÓ preencher se a transcrição menciona concorrente(s): para cada um, nome + sentimento (elogiou/neutro/criticou) + o que foi dito, curto e direto. Ex: 'LW: cliente elogiou o preço mas criticou o suporte lento.' Se vários concorrentes, separar por ; . Deixar vazio se nenhum foi mencionado.","concorrencia_contra_argumento":"SÓ preencher se o cliente ELOGIOU algum concorrente (é uma objeção real a contornar - se foi neutro ou criticou, deixar vazio). Sugira uma resposta prática e curta para o executivo usar na próxima call, baseada SOMENTE nos diferenciais reais e verificáveis da Frota162 (SNE com desconto, plano completo, Service/BPO documental para operar tudo pelo cliente, especialista dedicado, tabela de preços transparente sem surpresa). NUNCA inventar ou afirmar informação negativa não confirmada sobre o concorrente específico - o foco é reforçar o valor da Frota162, não atacar o concorrente."}`;

function callClaude(text) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 4000,
      system: SYSTEM,
      messages: [{ role: 'user', content: text }]
    });
    const req = https.request({
      hostname: 'api.anthropic.com', path: '/v1/messages', method: 'POST',
      headers: {
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
        'content-length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const p = JSON.parse(data);
          // Blindagem: a Anthropic pode retornar erro (rate limit, overload, etc.)
          // em formato sem 'content' — ex: {type:'error', error:{type, message}}
          if (p.type === 'error' || !p.content || !Array.isArray(p.content) || !p.content[0]) {
            const motivo = p.error?.message || p.error?.type || JSON.stringify(p).slice(0,200);
            reject(new Error('Claude API error: ' + motivo));
            return;
          }
          let t = p.content[0].text.replace(/```json/gi,'').replace(/```/g,'').trim();
          try {
            resolve(JSON.parse(t));
          } catch(e1) {
            // Tenta fechar JSON truncado
            let opens = 0, brackets = 0;
            for (const ch of t) {
              if (ch==='{') opens++; else if (ch==='}') opens--;
              else if (ch==='[') brackets++; else if (ch===']') brackets--;
            }
            t = t.replace(/,\s*$/, '').replace(/,\s*"[^"]*"\s*:\s*[^,}\]]*$/, '');
            t += ']'.repeat(Math.max(0,brackets)) + '}'.repeat(Math.max(0,opens));
            try { resolve(JSON.parse(t)); }
            catch(e2) { reject(new Error('Claude parse: ' + e1.message)); }
          }
        } catch(e) { reject(new Error('Claude response: ' + e.message)); }
      });
    });
    req.on('error', reject);
    req.setTimeout(180000, () => { req.destroy(); reject(new Error('Claude timeout')); });
    req.write(body); req.end();
  });
}

function fetchCallData(callId) {
  return new Promise((resolve) => {
    const req = https.request({
      hostname: 'api.elephan.dev',
      path: `/v1/transcribes/${callId}`,
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${process.env.ELEPHAN_API_KEY}`,
        'Accept': 'application/json'
      }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const p = JSON.parse(data);
          const obj = p.data || p;
          const texto = obj?.transcript?.text || obj?.content || '';
          // Data real da call — tenta vários campos, prioriza ISO
          const dataRaw = obj?.created_at || obj?.createdAt || obj?.date || obj?.dateIncluded || '';
          resolve({ texto, dataRaw });
        } catch(e) { resolve({ texto:'', dataRaw:'' }); }
      });
    });
    req.on('error', () => resolve({ texto:'', dataRaw:'' }));
    req.setTimeout(30000, () => { req.destroy(); resolve({ texto:'', dataRaw:'' }); });
    req.end();
  });
}

function postSlack(msg, webhookUrl) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({ text: msg });
    const url = new URL(webhookUrl || process.env.SLACK_WEBHOOK_URL);
    const req = https.request({
      hostname: url.hostname, path: url.pathname + url.search, method: 'POST',
      headers: { 'content-type': 'application/json', 'content-length': Buffer.byteLength(body) }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        // Slack webhook responde 200 com corpo "ok" em caso de sucesso.
        // Qualquer outro status significa que a mensagem NÃO foi entregue.
        if (res.statusCode >= 200 && res.statusCode < 300) resolve();
        else reject(new Error(`Slack respondeu ${res.statusCode}: ${data.slice(0,200)}`));
      });
    });
    req.on('error', reject);
    req.setTimeout(15000, () => { req.destroy(); reject(new Error('Slack timeout')); });
    req.write(body); req.end();
  });
}

function gerarPPTX(d, outPath) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9'; // 10 x 5.625"

  // Banda de oferta decidida em código pelo volume de placas (ver bandaOferta)
  const parceria = d.tipo_reuniao === 'parceria';
  let banda = parceria ? 'parceria' : bandaOferta(d.placas);
  if (banda === 'desconhecida') banda = 'service_somente'; // placas não mapeadas: regime mais restritivo (só Service)
  const temSlideService = (banda === 'service_somente' || banda === 'service_preferencial')
    && Array.isArray(d.service_com_itens) && d.service_com_itens.length > 0;

  // ── SLIDE 1 — A dor em linguagem executiva ────────────────────────────
  const s1 = pres.addSlide();
  s1.background = { color: 'F7F6F4' };

  // Header laranja full-width
  s1.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:1.00,fill:{color:COR.laranja}});
  s1.addText(d.s1_header_bold||'',{x:0.38,y:0.05,w:9.3,h:0.60,fontFace:'Montserrat',fontSize:14,bold:true,color:COR.branco,valign:'middle',margin:0,wrap:true});
  s1.addText(d.s1_header_sub||'',{x:0.38,y:0.66,w:9.3,h:0.24,fontFace:'Montserrat',fontSize:9.5,color:'FFD0C0',valign:'middle',margin:0});
  s1.addText(d.s1_subtitulo||'',{x:0.38,y:1.03,w:9.3,h:0.20,fontFace:'Montserrat',fontSize:8,italic:true,color:'888888',margin:0});

  // 4 cards 2x2
  const CPOS=[{cx:0.22,cy:1.28},{cx:5.14,cy:1.28},{cx:0.22,cy:3.08},{cx:5.14,cy:3.08}];
  const CW=4.60, CH=1.70;
  const CPAL=[
    {fundo:'FFF5F3',strip:COR.vermelho,stat:COR.vermelho},
    {fundo:'F0EFED',strip:'999999',stat:COR.dark},
    {fundo:'F0F5FF',strip:COR.azul,stat:COR.azul},
    {fundo:'F0EFED',strip:'999999',stat:COR.dark},
  ];
  (d.cards||[]).forEach((c,i)=>{
    if(i>3) return;
    const {cx,cy}=CPOS[i]; const p=CPAL[i];
    s1.addShape(pres.ShapeType.rect,{x:cx,y:cy,w:CW,h:CH,fill:{color:p.fundo},line:{color:COR.divisor,width:0.5}});
    s1.addShape(pres.ShapeType.rect,{x:cx,y:cy+0.14,w:0.05,h:CH-0.28,fill:{color:p.strip}});
    s1.addText(c.stat||'',{x:cx+0.16,y:cy+0.08,w:CW-0.24,h:0.42,fontFace:'Montserrat',fontSize:22,bold:true,color:p.stat,margin:0});
    s1.addText(c.titulo||'',{x:cx+0.16,y:cy+0.50,w:CW-0.24,h:0.24,fontFace:'Montserrat',fontSize:9.5,bold:true,color:COR.dark,margin:0});
    s1.addText(c.desc||'',{x:cx+0.16,y:cy+0.76,w:CW-0.24,h:0.87,fontFace:'Montserrat',fontSize:7.8,color:'333333',valign:'top',margin:0,wrap:true});
  });

  // Footer dark slide 1
  s1.addShape(pres.ShapeType.rect,{x:0,y:5.18,w:10,h:0.445,fill:{color:COR.dark}});
  s1.addShape(pres.ShapeType.rect,{x:0,y:5.18,w:0.05,h:0.445,fill:{color:COR.laranja}});
  s1.addText([
    {text:(d.s1_footer_bold||'')+' ',options:{bold:true,color:COR.laranja}},
    {text:d.s1_footer_normal||'',options:{bold:false,color:COR.branco}}
  ],{x:0.22,y:5.20,w:7.80,h:0.38,fontFace:'Montserrat',fontSize:9,valign:'middle',margin:0});
  s1.addText('frota162.com.br',{x:8.20,y:5.24,w:1.65,h:0.28,fontFace:'Montserrat',fontSize:8,bold:true,color:COR.laranja,align:'right',valign:'middle',margin:0});

  // ── SLIDE 2 — SERVICE (conceitual, sem valores em R$) ───────────────────
  // Gerado para 1 a 60 placas (matriz de oferta). 1 a 10 placas: coluna única,
  // SEM o contraponto "sem service / você opera" — o lead nem deve saber que
  // existe a opção de operar a plataforma sozinho. 11 a 60: SEM x COM.
  if (temSlideService) {
    const soService = (banda === 'service_somente');
    const s1s = pres.addSlide();
    s1s.background = { color: 'F7F6F4' };
    s1s.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:0.92,fill:{color:COR.laranja}});
    s1s.addText(d.service_header_bold||'',{x:0.38,y:0.05,w:9.3,h:0.58,fontFace:'Montserrat',fontSize:14,bold:true,color:COR.branco,valign:'middle',margin:0,wrap:true});
    s1s.addText(d.service_header_sub||'',{x:0.38,y:0.66,w:9.3,h:0.24,fontFace:'Montserrat',fontSize:10,color:'FFD0C0',valign:'middle',margin:0});

    const SCY = 0.98, SCH = 2.75;
    const SC1X=0.30, SCW=4.55, SC2X=5.15;

    // Coluna SEM SERVICE (neutro, sem borda) — só de 11 a 60 placas
    if (!soService) {
      s1s.addShape(pres.ShapeType.rect,{x:SC1X,y:SCY,w:SCW,h:SCH,fill:{color:'F0EFED'}});
      s1s.addShape(pres.ShapeType.rect,{x:SC1X,y:SCY,w:SCW,h:0.42,fill:{color:'DDDBD8'}});
      s1s.addText(d.service_sem_titulo||'SEM SERVICE',{x:SC1X+0.16,y:SCY,w:SCW-0.32,h:0.42,fontFace:'Montserrat',fontSize:10,bold:true,color:COR.dark,valign:'middle',margin:0});
      (d.service_sem_itens||[]).forEach((item,i)=>{
        const iy = SCY+0.52+i*0.55;
        s1s.addShape(pres.ShapeType.ellipse,{x:SC1X+0.18,y:iy+0.02,w:0.16,h:0.16,fill:{color:'999999'}});
        s1s.addText(item,{x:SC1X+0.44,y:iy-0.06,w:SCW-0.62,h:0.48,fontFace:'Montserrat',fontSize:8,color:'444444',valign:'top',margin:0,wrap:true});
      });
    }

    // Coluna COM SERVICE (verde, sem borda) — largura total quando é só Service
    const COMX = soService ? SC1X : SC2X;
    const COMW = soService ? 9.40 : SCW;
    const comTituloPadrao = soService ? 'O QUE A FROTA162 OPERA POR VOCÊ' : 'COM SERVICE';
    s1s.addShape(pres.ShapeType.rect,{x:COMX,y:SCY,w:COMW,h:SCH,fill:{color:'EAF6EA'}});
    s1s.addShape(pres.ShapeType.rect,{x:COMX,y:SCY,w:COMW,h:0.42,fill:{color:COR.verde}});
    s1s.addText(d.service_com_titulo||comTituloPadrao,{x:COMX+0.16,y:SCY,w:COMW-0.32,h:0.42,fontFace:'Montserrat',fontSize:10,bold:true,color:COR.branco,valign:'middle',margin:0});
    (d.service_com_itens||[]).forEach((item,i)=>{
      const iy = SCY+0.52+i*0.55;
      s1s.addShape(pres.ShapeType.ellipse,{x:COMX+0.18,y:iy+0.02,w:0.16,h:0.16,fill:{color:COR.verde}});
      s1s.addText(item,{x:COMX+0.44,y:iy-0.06,w:COMW-0.62,h:0.48,fontFace:'Montserrat',fontSize:soService?9:8,bold:true,color:COR.dark,valign:'top',margin:0,wrap:true});
    });

    // 3 mini-cards de benefício (Tempo / Custo / Risco) - sem borda, 100% qualitativo
    const BX0 = 0.30, BW = 3.00, BG = 0.20, BY = SCY+SCH+0.10, BH = 0.85;
    (d.service_beneficios||[]).forEach((b,i)=>{
      const bx = BX0 + i*(BW+BG);
      s1s.addShape(pres.ShapeType.rect,{x:bx,y:BY,w:BW,h:BH,fill:{color:'EAF6EA'}});
      s1s.addText(b.stat||'',{x:bx+0.14,y:BY+0.08,w:BW-0.28,h:0.36,fontFace:'Montserrat',fontSize:20,bold:true,color:COR.verde,margin:0});
      s1s.addText(b.titulo||'',{x:bx+0.14,y:BY+0.42,w:BW-0.28,h:0.20,fontFace:'Montserrat',fontSize:8.5,bold:true,color:COR.dark,margin:0});
      s1s.addText(b.desc||'',{x:bx+0.14,y:BY+0.60,w:BW-0.28,h:0.24,fontFace:'Montserrat',fontSize:6.8,color:'555555',valign:'top',margin:0,wrap:true});
    });

    // Nota final - reforço vendedor
    s1s.addText(d.service_nota||'',{x:0.38,y:BY+BH+0.08,w:9.24,h:0.24,fontFace:'Montserrat',fontSize:7.5,italic:true,bold:true,color:COR.verde,margin:0,wrap:true});

    s1s.addShape(pres.ShapeType.rect,{x:0,y:5.18,w:10,h:0.445,fill:{color:COR.dark}});
    s1s.addShape(pres.ShapeType.rect,{x:0,y:5.18,w:0.05,h:0.445,fill:{color:COR.laranja}});
    s1s.addText([
      {text:'Você ganha um especialista dedicado. ',options:{bold:true,color:COR.laranja}},
      {text:'Sem contratar, sem treinar, sem se preocupar.',options:{bold:false,color:COR.branco}}
    ],{x:0.22,y:5.20,w:7.80,h:0.38,fontFace:'Montserrat',fontSize:8.5,valign:'middle',margin:0});
    s1s.addText('frota162.com.br',{x:8.20,y:5.24,w:1.65,h:0.28,fontFace:'Montserrat',fontSize:8,bold:true,color:COR.laranja,align:'right',valign:'middle',margin:0});
  }

  // ── SLIDE 3 — 3 Zonas (HOJE → COMO RESOLVE → RESULTADO) ─────────────
  const s2 = pres.addSlide();
  s2.background = { color: 'F7F6F4' };

  // Header dark
  s2.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:0.84,fill:{color:COR.dark}});
  s2.addShape(pres.ShapeType.rect,{x:0,y:0,w:0.05,h:0.84,fill:{color:COR.laranja}});
  s2.addText(d.s2_header_bold||'',{x:0.22,y:0.04,w:9.4,h:0.42,fontFace:'Montserrat',fontSize:12,bold:true,color:COR.laranja,valign:'middle',margin:0,wrap:true});
  s2.addText(d.s2_header_normal||'',{x:0.22,y:0.50,w:9.4,h:0.26,fontFace:'Montserrat',fontSize:8.5,color:COR.branco,valign:'middle',margin:0});

  const CY=0.84, CH2=4.34; // termina em y=5.18
  const Z1X=0, Z1W=3.10, Z2X=3.10, Z2W=3.80, Z3X=6.90, Z3W=3.10;

  // ZONA 1 — HOJE (vermelho claro)
  s2.addShape(pres.ShapeType.rect,{x:Z1X,y:CY,w:Z1W,h:CH2,fill:{color:'FFEDE7'}});
  s2.addText('HOJE',{x:Z1X+0.18,y:CY+0.16,w:Z1W-0.24,h:0.22,fontFace:'Montserrat',fontSize:9,bold:true,color:COR.vermelho,charSpacing:1,margin:0});

  // Stat 1
  const z1s1Sz=(d.z1_stat1||'').length>10?15:(d.z1_stat1||'').length>6?20:26;
  s2.addText(d.z1_stat1||'',{x:Z1X+0.18,y:CY+0.42,w:Z1W-0.24,h:0.48,fontFace:'Montserrat',fontSize:z1s1Sz,bold:true,color:COR.vermelho,margin:0,wrap:true});
  s2.addText(d.z1_sub1||'',{x:Z1X+0.18,y:CY+0.92,w:Z1W-0.24,h:0.30,fontFace:'Montserrat',fontSize:8,color:'555555',margin:0,wrap:true});

  // Divisor
  s2.addShape(pres.ShapeType.rect,{x:Z1X+0.18,y:CY+1.28,w:Z1W-0.36,h:0.016,fill:{color:'F0B8A5'}});

  // Stat 2
  const z1s2Sz=(d.z1_stat2||'').length>10?12:(d.z1_stat2||'').length>6?15:18;
  s2.addText(d.z1_stat2||'',{x:Z1X+0.18,y:CY+1.34,w:Z1W-0.24,h:0.36,fontFace:'Montserrat',fontSize:z1s2Sz,bold:true,color:COR.vermelho,margin:0,wrap:true});
  s2.addText(d.z1_sub2||'',{x:Z1X+0.18,y:CY+1.72,w:Z1W-0.24,h:0.30,fontFace:'Montserrat',fontSize:8,color:'555555',margin:0,wrap:true});

  // Bullets — posição dinâmica com base no espaço restante
  const bulletY = CY + 2.10;
  (d.z1_bullets||[]).forEach((b,i)=>{
    s2.addText('· '+b,{x:Z1X+0.18,y:bulletY+(i*0.30),w:Z1W-0.24,h:0.28,fontFace:'Montserrat',fontSize:7.5,color:'444444',margin:0,wrap:true});
  });

  // ZONA 2 — COMO A FROTA162 RESOLVE (quase branco)
  s2.addShape(pres.ShapeType.rect,{x:Z2X,y:CY,w:Z2W,h:CH2,fill:{color:'FCFCFB'}});
  s2.addText(parceria ? 'COMO FUNCIONA A PARCERIA' : 'COMO A FROTA162 RESOLVE',{x:Z2X+0.18,y:CY+0.16,w:Z2W-0.24,h:0.22,fontFace:'Montserrat',fontSize:8.5,bold:true,color:COR.dark,charSpacing:0.5,margin:0});

  // Rail vertical + 4 passos numerados
  const SY=CY+0.52, SH=0.74, SG=0.10, rx=Z2X+0.34;
  const railH = (SH+SG)*3 + SH;
  s2.addShape(pres.ShapeType.rect,{x:rx+0.15,y:SY+0.19,w:0.014,h:railH-0.10,fill:{color:COR.divisor}});

  (d.passos||[]).forEach((p,i)=>{
    const py=SY+i*(SH+SG);
    s2.addShape(pres.ShapeType.ellipse,{x:rx,y:py+0.05,w:0.38,h:0.38,fill:{color:COR.laranja}});
    s2.addText(String(i+1),{x:rx,y:py+0.05,w:0.38,h:0.38,fontFace:'Montserrat',fontSize:11,bold:true,color:COR.branco,align:'center',valign:'middle',margin:0});
    s2.addText(p.titulo||'',{x:Z2X+0.82,y:py+0.05,w:Z2W-0.98,h:0.28,fontFace:'Montserrat',fontSize:9,bold:true,color:COR.dark,margin:0,wrap:true});
    s2.addText(p.desc||'',{x:Z2X+0.82,y:py+0.34,w:Z2W-0.98,h:0.36,fontFace:'Montserrat',fontSize:8,color:'555555',valign:'top',margin:0,wrap:true});
  });

  // ZONA 3 — RESULTADO / PLANO RECOMENDADO (verde claro)
  s2.addShape(pres.ShapeType.rect,{x:Z3X,y:CY,w:Z3W,h:CH2,fill:{color:'EAF6EA'}});
  const z3Rotulo = parceria ? 'MODELO DA PARCERIA' : (d.tem_roi ? 'RESULTADO' : (banda === 'planos' ? 'PLANO RECOMENDADO' : 'OFERTA RECOMENDADA'));
  s2.addText(z3Rotulo,{x:Z3X+0.18,y:CY+0.16,w:Z3W-0.24,h:0.22,fontFace:'Montserrat',fontSize:9,bold:true,color:COR.verde,charSpacing:1,margin:0});

  // z3_stat — fonte adaptativa
  const z3Sz = (d.z3_stat||'').length > 8 ? 16 : 26;
  s2.addText(d.z3_stat||'',{x:Z3X+0.18,y:CY+0.40,w:Z3W-0.24,h:0.50,fontFace:'Montserrat',fontSize:z3Sz,bold:true,color:COR.verde,margin:0,wrap:true});
  s2.addText(d.z3_sub1||'',{x:Z3X+0.18,y:CY+0.92,w:Z3W-0.24,h:0.30,fontFace:'Montserrat',fontSize:8,color:'555555',margin:0,wrap:true});
  s2.addShape(pres.ShapeType.rect,{x:Z3X+0.18,y:CY+1.28,w:Z3W-0.36,h:0.016,fill:{color:'BFE3BF'}});
  s2.addText(d.z3_investimento||'A confirmar',{x:Z3X+0.18,y:CY+1.34,w:Z3W-0.24,h:0.36,fontFace:'Montserrat',fontSize:16,bold:true,color:COR.verde,margin:0,wrap:true});

  // Badge diferencial
  s2.addShape(pres.ShapeType.rect,{x:Z3X+0.18,y:CY+1.82,w:Z3W-0.36,h:0.36,fill:{color:'D4EED4'},line:{color:'BFE3BF',width:0.5}});
  s2.addText(d.z3_badge||'',{x:Z3X+0.22,y:CY+1.82,w:Z3W-0.44,h:0.36,fontFace:'Montserrat',fontSize:7.5,bold:true,color:COR.verde,align:'center',valign:'middle',margin:0,wrap:true});

  // Zona 3: linha secundária (opção avulsa, 11 a 60 placas) OU Payback — nunca os dois, evita colisão visual.
  // 1 a 10 placas nunca tem linha secundária (lead não pode saber que existe plano avulso).
  if (banda === 'service_preferencial' && d.z3_alt_investimento) {
    s2.addShape(pres.ShapeType.rect,{x:Z3X+0.18,y:CY+2.32,w:Z3W-0.36,h:0.016,fill:{color:'BFE3BF'}});
    s2.addText(d.z3_alt_label||'Ou só a plataforma - você opera',{x:Z3X+0.18,y:CY+2.40,w:Z3W-0.24,h:0.22,fontFace:'Montserrat',fontSize:7.5,bold:true,color:'558855',margin:0,wrap:true});
    s2.addText(d.z3_alt_investimento,{x:Z3X+0.18,y:CY+2.62,w:Z3W-0.24,h:0.32,fontFace:'Montserrat',fontSize:14,bold:true,color:COR.verde,margin:0,wrap:true});
    if (d.z3_alt_tagline) {
      s2.addText(d.z3_alt_tagline,{x:Z3X+0.18,y:CY+2.96,w:Z3W-0.24,h:0.30,fontFace:'Montserrat',fontSize:7.5,bold:true,italic:true,color:COR.laranja,margin:0,wrap:true});
    }
  } else {
    // Payback na zona 3 — SOMENTE quando tem_roi=true e ROI confirmado na call
    const dp = d.dias_payback||0;
    const invN = d.investimento_mensal_num||0;
    const roiAnualN = d.roi_anual||0;
    const economiaMensalN = roiAnualN > 0 ? roiAnualN / 12 : 0;
    const dpCalcS2 = d.tem_roi && roiAnualN > 0 && invN > 0 ? Math.round(invN / economiaMensalN * 30) : 0;
    if(dpCalcS2 > 0 && dpCalcS2 <= 365){
      const pbLabelS2 = dpCalcS2<=45 ? `↑ Payback: ~${dpCalcS2} dias` : `↑ Payback: ~${Math.round(dpCalcS2/30)} meses`;
      s2.addShape(pres.ShapeType.rect,{x:Z3X+0.18,y:CY+2.28,w:Z3W-0.36,h:0.32,fill:{color:COR.branco},line:{color:COR.verde,width:1.0}});
      s2.addText(pbLabelS2,{x:Z3X+0.22,y:CY+2.28,w:Z3W-0.44,h:0.32,fontFace:'Montserrat',fontSize:8.5,bold:true,color:COR.verde,align:'center',valign:'middle',margin:0});
    }
  }

  // Nota rodapé zona 3
  s2.addText(d.z3_nota||'',{x:Z3X+0.18,y:CY+3.42,w:Z3W-0.24,h:0.60,fontFace:'Montserrat',fontSize:6.5,italic:true,color:'AAAAAA',valign:'top',margin:0,wrap:true});

  // Footer dark slide 2 — mesmo padrão slide 1
  s2.addShape(pres.ShapeType.rect,{x:0,y:5.18,w:10,h:0.445,fill:{color:COR.dark}});
  s2.addShape(pres.ShapeType.rect,{x:0,y:5.18,w:0.05,h:0.445,fill:{color:COR.laranja}});
  s2.addText([
    {text:(d.s2_cta_bold||'')+'  ',options:{bold:true,color:COR.laranja}},
    {text:d.s2_cta_normal||'',options:{bold:false,color:COR.branco}}
  ],{x:0.22,y:5.20,w:7.80,h:0.38,fontFace:'Montserrat',fontSize:8.5,valign:'middle',margin:0});
  s2.addText('frota162.com.br',{x:8.20,y:5.24,w:1.65,h:0.28,fontFace:'Montserrat',fontSize:8,bold:true,color:COR.laranja,align:'right',valign:'middle',margin:0});

  // ── SLIDE 4 — O custo de esperar ──────────────────────────────────────
  // SÓ é gerado quando o ROI foi confirmado na call (tem_roi) E há custo mensal
  // calculado. Sem isso o gráfico saía em branco (barras de R$0) — pior que não ter.
  const cm = Number(d.custo_mensal) || 0;
  const temSlideCusto = !!d.tem_roi && cm > 0;
  if (temSlideCusto) {
    const s3 = pres.addSlide();
    s3.background = { color: 'F7F6F4' };

    // Header laranja
    s3.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:1.00,fill:{color:COR.laranja}});
    s3.addText(d.s3_header_bold||'',{x:0.38,y:0.05,w:9.3,h:0.58,fontFace:'Montserrat',fontSize:14,bold:true,color:COR.branco,valign:'middle',margin:0,wrap:true});
    s3.addText(d.s3_header_sub||'',{x:0.38,y:0.65,w:9.3,h:0.24,fontFace:'Montserrat',fontSize:9,color:'FFD0C0',valign:'middle',margin:0});
    s3.addText(d.s3_formula||'',{x:0.38,y:1.02,w:9.3,h:0.20,fontFace:'Montserrat',fontSize:7.5,italic:true,color:'888888',margin:0,wrap:true});

    // 6 barras: 10, 20, 30, 45, 60, 90 dias
    const DIAS=[10,20,30,45,60,90];
    const VALS=DIAS.map(d=>cm*(d/30));
    const BY=4.52, MH=2.90, BW=1.35, BG=0.16, X0=0.30;

    // Payback slide 3 — SOMENTE quando tem_roi=true e ROI confirmado
    const inv2=d.investimento_mensal_num||0;
    const roiAnualS3 = d.roi_anual||0;
    const economiaMensalS3 = roiAnualS3 > 0 ? roiAnualS3 / 12 : 0;
    const dpCalc = (d.tem_roi && roiAnualS3 > 0 && inv2 > 0) ? Math.round(inv2 / economiaMensalS3 * 30) : 0;

    // Linha baseline
    s3.addShape(pres.ShapeType.rect,{x:0.20,y:BY,w:9.60,h:0.014,fill:{color:'CCCCCC'}});

    // Payback — posição fixa no topo (y=1.28), nunca sobrepõe barras
    let paybackBi = -1;
    if(dpCalc>0){
      paybackBi = DIAS.findIndex(x=>x>=dpCalc); if(paybackBi<0) paybackBi=5;
      const pbx=X0+paybackBi*(BW+BG);
      const pbLabel = dpCalc<=45 ? `↑ Payback: ~${dpCalc} dias` : `↑ Payback: ~${Math.round(dpCalc/30)} meses`;
      const PAYBACK_Y = 1.28;
      s3.addShape(pres.ShapeType.rect,{x:pbx+0.06,y:PAYBACK_Y,w:BW-0.12,h:0.32,fill:{color:COR.branco},line:{color:COR.verde,width:1.5}});
      s3.addText(pbLabel,{x:pbx+0.06,y:PAYBACK_Y,w:BW-0.12,h:0.32,fontFace:'Montserrat',fontSize:7.5,bold:true,color:COR.verde,align:'center',valign:'middle',margin:0});
      // Linha tracejada do payback até a barra
      const bHpb=MH*(DIAS[paybackBi]/90), bToppb=BY-bHpb;
      const lineStart = PAYBACK_Y+0.32;
      const lineHeight = Math.max(0, bToppb - lineStart);
      if(lineHeight>0) s3.addShape(pres.ShapeType.rect,{x:pbx+(BW/2)-0.007,y:lineStart,w:0.014,h:lineHeight,fill:{color:'BFE3BF'}});
    }

    // Barras
    DIAS.forEach((dia,i)=>{
      const bH=MH*(dia/90), bTop=BY-bH, bx=X0+i*(BW+BG);
      const isPayback = (i===paybackBi);
      s3.addShape(pres.ShapeType.rect,{x:bx,y:bTop,w:BW,h:bH,fill:{color:isPayback?'FFB3B3':'FFCDD2'},line:{color:COR.vermelho,width:0.5}});
      s3.addText(`-R$${Math.round(VALS[i]).toLocaleString('pt-BR')}`,{x:bx,y:bTop-0.30,w:BW,h:0.26,fontFace:'Montserrat',fontSize:9.5,bold:true,color:COR.vermelho,align:'center',margin:0});
      s3.addText(`${dia}d`,{x:bx,y:BY+0.06,w:BW,h:0.22,fontFace:'Montserrat',fontSize:8,bold:true,color:'555555',align:'center',margin:0});
    });

    // Nota — abaixo das labels
    s3.addText(d.s3_nota||'',{x:0.30,y:4.84,w:9.40,h:0.24,fontFace:'Montserrat',fontSize:6,italic:true,color:'AAAAAA',valign:'top',margin:0,wrap:true});

    // Footer dark slide 3 — FIXO no rodapé (y=5.18)
    s3.addShape(pres.ShapeType.rect,{x:0,y:5.18,w:10,h:0.445,fill:{color:COR.dark}});
    s3.addShape(pres.ShapeType.rect,{x:0,y:5.18,w:0.05,h:0.445,fill:{color:COR.laranja}});
    s3.addText([
      {text:'Decisão adiada não é decisão neutra.  ',options:{bold:true,color:COR.laranja}},
      {text:`R$${Math.round(cm).toLocaleString('pt-BR')} por mês continuam saindo do caixa — com ou sem contrato assinado.`,options:{bold:false,color:COR.branco}}
    ],{x:0.22,y:5.20,w:7.80,h:0.38,fontFace:'Montserrat',fontSize:8.5,valign:'middle',margin:0});
    s3.addText('frota162.com.br',{x:8.20,y:5.24,w:1.65,h:0.28,fontFace:'Montserrat',fontSize:8,bold:true,color:COR.laranja,align:'right',valign:'middle',margin:0});
  }

  return pres.writeFile({ fileName: outPath });
}

// Remove emojis e chars especiais que quebram pptxgenjs
// Remove caracteres que quebram caminhos de arquivo (barras, dois-pontos, etc.)
// Bug real encontrado: nome de empresa com "/" (ex: "FM Rodrigues / Salfena")
// fazia o sistema de arquivos interpretar como subpasta inexistente -> ENOENT.
function sanitizeFileName(str) {
  if (!str) return 'Prospect';
  return String(str)
    .replace(/[\/\\:*?"<>|]/g, '-')  // caracteres proibidos em nomes de arquivo -> hífen
    .replace(/\s+/g, ' ')
    .trim();
}

function sanitize(str) {
  if (!str) return '';
  return String(str)
    .replace(/[\u{1F000}-\u{1FFFF}]/gu, '')
    .replace(/[\u{2600}-\u{27BF}]/gu, '')
    .replace(/[\u{1F300}-\u{1F9FF}]/gu, '')
    .replace(/[^\x00-\x7F\xC0-\u024F]/g, (c) => {
      const code = c.charCodeAt(0);
      return (code >= 0x00C0 && code <= 0x024F) ? c : '';
    })
    .replace(/\s+/g, ' ').trim();
}

app.get('/', (req, res) => res.json({ status: 'ok', service: 'Frota162 PPTX v5' }));

app.post('/generate', (req, res) => {
  res.json({ ok: true, status: 'processing' });

  (async () => {
    // Declaradas ANTES do try — se o JSON.parse falhar, o catch final ainda
    // consegue montar a mensagem de erro sem quebrar com ReferenceError
    // (bug real que já causou um crash do processo inteiro — "Exited with status 1").
    let titulo, executivo, call_id;
    try {
      let raw = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
      raw = raw.replace(/```json/gi,'').replace(/```/g,'').trim();
      const input = JSON.parse(raw);
      ({ titulo, executivo, call_id } = input);

      // Filtro 1 — título deve conter "Frota162 ><" ou "Frota162 <>"
      const tituloLower = (titulo||'').toLowerCase();
      const ehReuniaoCliente = tituloLower.includes('frota162 ><') || tituloLower.includes('frota162 <>') || tituloLower.includes('frota162><') || tituloLower.includes('frota162<>');
      if (!ehReuniaoCliente) {
        console.log('Descartado — não é reunião com cliente:', titulo);
        return;
      }

      // Filtro 2 — apenas executivos autorizados
      const EXECUTIVOS = ['palloma', 'julio', 'júlio', 'ravila', 'rávila', 'thais', 'william', 'willîam', 'bruno pereira'];
      const execLower = (executivo||'').toLowerCase();
      const ehExecutivoAutorizado = EXECUTIVOS.some(e => execLower.includes(e));
      if (!ehExecutivoAutorizado) {
        console.log('Descartado — executivo não autorizado:', executivo);
        return;
      }

      // Cliente Drive único para toda a execução (usado no controle de duplicatas e no upload)
      const drive = getDriveClient();

      // Controle de duplicatas à prova de corrida (marcador atômico no Drive)
      const callId = call_id || titulo;
      if (callId) {
        const devoProcessar = await claimCall(drive, callId);
        if (!devoProcessar) {
          console.log('Call já processada ou perdeu a corrida, pulando:', callId);
          return;
        }
      }

      // Busca dados reais no Elephan (transcrição + data da call)
      const callData = await fetchCallData(call_id);
      const transcricao = callData.texto;
      const dataRaw = callData.dataRaw;

      // Filtro 0 — apenas calls de HOJE (usa a data REAL da API Elephan, horário Brasília UTC-3)
      const offsetBrasilia = 3 * 60;
      const agoraBrasilia = new Date(Date.now() - offsetBrasilia * 60 * 1000);
      const hojeStr = agoraBrasilia.toISOString().slice(0, 10);

      const dObj = new Date(dataRaw);
      let dataCallStr = '';
      let dataCallFormatada = 'Data não informada';
      if (!isNaN(dObj.getTime())) {
        const dBrasilia = new Date(dObj.getTime() - offsetBrasilia * 60 * 1000);
        dataCallStr = dBrasilia.toISOString().slice(0, 10);
        dataCallFormatada = dBrasilia.toISOString().slice(0, 16).replace('T', ' ');
      }

      // Se a data não é de hoje, descarta silenciosamente
      if (dataCallStr && dataCallStr !== hojeStr) {
        console.log('Descartado — call não é de hoje:', dataCallStr, 'hoje:', hojeStr, titulo);
        return;
      }
      // Se não conseguiu determinar a data, descarta por segurança (evita processar calls antigas sem data)
      if (!dataCallStr) {
        console.log('Descartado — data da call não pôde ser determinada:', titulo);
        return;
      }

      // Filtro transcrição — avisa no Slack, mas SÓ UMA VEZ por call (marcador dedicado
      // evita que a mesma call vazia continue gerando o mesmo aviso a cada execução do Make,
      // já que ela nunca vai ganhar o marcador "processed_" por definição).
      if (!transcricao || transcricao.length < 500) {
        const jaAvisou = await isMarked(drive, `descartada_${callId}`);
        if (!jaAvisou) {
          await postSlack(`:no_entry_sign: *Call descartada — ${titulo||'Sem título'}* (${executivo||''}): transcrição ausente ou muito curta para gerar material.`).catch(()=>{});
          await markGeneric(drive, `descartada_${callId}`);
        } else {
          console.log('Descartado (silencioso, já avisado antes) — transcrição curta:', callId);
        }
        return;
      }

      const conteudo = `Título: ${titulo||'Sem título'}\nData: ${dataCallFormatada}\nExecutivo Frota162: ${executivo||''}\n\nTranscrição:\n${transcricao}`;

      // Retry automático — erros da Anthropic (rate limit/overload) costumam ser transitórios
      let d;
      let lastErr;
      for (let tentativa = 1; tentativa <= 3; tentativa++) {
        try {
          d = await callClaude(conteudo);
          break;
        } catch(e) {
          lastErr = e;
          console.log(`callClaude tentativa ${tentativa} falhou:`, e.message);
          if (tentativa < 3) await new Promise(r => setTimeout(r, 5000 * tentativa));
        }
      }
      if (!d) throw lastErr;

      // Validação pós-Claude — descarta silenciosamente se campos essenciais estiverem vazios
      const empresaValida = d.empresa && d.empresa !== 'Empresa Não Identificada' && d.empresa !== 'Não identificado' && d.empresa !== '';
      const placasValidas = d.placas && d.placas > 0;
      if (!empresaValida || !placasValidas) {
        console.log('Descartado — campos insuficientes após análise Claude:', titulo, '| empresa:', d.empresa, '| placas:', d.placas);
        return;
      }

      const empresa = d.empresa || 'Prospect';
      const nomeArq = `Frota162 >< ${sanitizeFileName(empresa)} (Diretoria).pptx`;
      const outPath = path.join(os.tmpdir(), nomeArq);

      const sanitizeObj = (obj) => {
        if (typeof obj === 'string') return sanitize(obj);
        if (Array.isArray(obj)) return obj.map(sanitizeObj);
        if (obj && typeof obj === 'object') { const r={}; for(const k of Object.keys(obj)) r[k]=sanitizeObj(obj[k]); return r; }
        return obj;
      };
      const dClean = sanitizeObj(d);
      await gerarPPTX(dClean, outPath);

      const pastaId = process.env.PASTA_RAIZ_ID;

      const uploaded = await drive.files.create({
        supportsAllDrives: true,
        requestBody: { name: nomeArq, parents: [pastaId], mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
        media: { mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', body: fs.createReadStream(outPath) },
        fields: 'id,name,webViewLink',
      });

      await drive.permissions.create({ fileId: uploaded.data.id, supportsAllDrives: true, requestBody: { role: 'writer', type: 'anyone' } });
      fs.unlinkSync(outPath);
      // ID já foi reservado no início — não precisa marcar de novo aqui

      // Temperatura com emoji
      const tempEmoji = d.temperatura==='quente' ? '🔴' : d.temperatura==='morno' ? '🟡' : '🔵';

      // Slack ID do executivo
      // O nome do executivo SEMPRE vem do Elephan/Make (fonte confiável) — NUNCA do Claude,
      // que pode confundir com nomes mencionados na transcrição do lado do cliente.
      const slackId = getSlackId(executivo);
      const execMencao = slackId ? `<@${slackId}>` : (executivo || 'N/A');

      // ROI formatado
      const roiAnual = d.roi_anual || 0;
      const roiTexto = roiAnual > 0 ? `R$${Math.round(roiAnual).toLocaleString('pt-BR')}/ano` : 'A calcular';

      const dataHoraReuniao = dataCallFormatada;
      const msg = `:car: *Novo material e análise estratégica* :rocket:\n\n- *Empresa:* ${empresa}\n- *Executivo:* ${execMencao}\n- *Data da reunião:* ${dataHoraReuniao}\n- *Placas e MRR estimado:* ${d.placas||0} placas · ${textoMrr(d)}\n- *ROI estimado:* ${roiTexto}\n- *Material:* <${uploaded.data.webViewLink}|Abrir PPTX>\n- *Temperatura estimada:* ${tempEmoji} ${d.temperatura||'N/A'}\n- *Resumo Geral da negociação:* ${d.slack_resumo||''}`;

      await postSlack(msg);

      // SÓ agora, com o Slack confirmado, marca sucesso definitivo.
      // Se qualquer etapa anterior falhar, este marcador nunca é criado
      // e a call pode ser tentada de novo no próximo lote do Make.
      if (callId) await markProcessed(drive, callId);

    } catch(err) {
      console.error('Background error:', err.message);
      try {
        await postSlack(`:warning: *Erro ao gerar material* — ${titulo||'Sem título'} (${executivo||'?'})\nMotivo: ${err.message}`);
      } catch(e2) {
        console.error('Falha ao avisar erro no Slack:', e2.message);
      }
    }
  })();
});

// ═══════════════════════════════════════════════════════════════════════
// SALESBUD — PIPELINE PARALELO DE TESTE (isolado do fluxo Elephan acima)
// Nada neste bloco modifica ou depende do /generate. Namespace de dedup
// próprio (prefixo "sb_"), postagem em canal de TESTE separado no Slack.
// ═══════════════════════════════════════════════════════════════════════

// Mapa userId (Salesbud) → nome do executivo Frota162
const SALESBUD_USER_MAP = {
  '15361': 'William Duarte',
  '15360': 'Palloma Santos',
  '15359': 'Thais Cristina',
  '15358': 'Julio Mazzetti',
  '15357': 'Rávila Silva',
  '15356': 'Bruno Pereira',
};

// Chave única de uma reunião, igual para o webhook (id numérico) e para a API (id mtg_...):
// título só com letras/números (tolerante a como cada fonte sanitiza "<>" e espaços) +
// minuto do início. É o que impede a mesma reunião de ser processada duas vezes.
function chaveReuniao(titulo, meetingAt) {
  const t = String(titulo || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const ms = new Date(meetingAt).getTime();
  const minuto = isNaN(ms) ? 'sem-data' : Math.floor(ms / 60000);
  return 'mk_' + crypto.createHash('sha1').update(`${t}|${minuto}`).digest('hex').slice(0, 24);
}

// A Salesbud manda a transcrição em HTML (<p><strong>João:</strong> texto...).
// Converte para texto plano preservando quebras de linha por parágrafo.
function stripHtml(html) {
  if (!html) return '';
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// Verificação de assinatura HMAC — BEST EFFORT.
// A doc da Salesbud não especifica o nome exato do header nem o algoritmo.
// Se SALESBUD_WEBHOOK_SECRET não estiver configurado, aceita sem verificar.
// Se estiver configurado mas nenhum header conhecido for encontrado, ACEITA
// e LOGA todos os headers recebidos — isso é o que vamos usar para descobrir
// o header real na primeira entrega de teste, e então enrijecer a checagem.
function verificaAssinaturaSalesbud(req, rawBody) {
  const secret = process.env.SALESBUD_WEBHOOK_SECRET;
  if (!secret) return { ok: true, motivo: 'sem secret configurado - aceito sem verificacao' };

  const candidatos = ['x-salesbud-signature', 'x-signature', 'x-webhook-signature', 'x-hub-signature-256'];
  let headerEncontrado = null, valorHeader = null;
  for (const h of candidatos) {
    if (req.headers[h]) { headerEncontrado = h; valorHeader = req.headers[h]; break; }
  }
  if (!headerEncontrado) {
    console.log('[Salesbud] AVISO: nenhum header de assinatura reconhecido nesta entrega.');
    console.log('[Salesbud] Headers recebidos (usar para identificar o header real):', JSON.stringify(req.headers));
    return { ok: true, motivo: 'header de assinatura nao encontrado - aceito temporariamente' };
  }

  const hashCalculado = crypto.createHmac('sha256', secret).update(rawBody).digest('hex');
  const valorLimpo = String(valorHeader).replace(/^sha256=/, '');
  const valido = hashCalculado === valorLimpo;
  console.log(`[Salesbud] Verificação HMAC via header '${headerEncontrado}': ${valido ? 'OK' : 'FALHOU'}`);
  return { ok: valido, motivo: valido ? 'assinatura valida' : 'assinatura invalida' };
}

// Núcleo do pipeline Salesbud — usado pelo webhook E pela sincronização via API.
// Retorna: 'ok' | 'ja_processada' | 'descartada' (decisão final: não adianta repetir)
// | 'erro' (falha transitória: vale tentar de novo no próximo ciclo).
async function processarPayloadSalesbud(payload) {
  let titulo, executivo, callId, chaveUnica;
  try {
      // Só processamos o payload de "Reunião" (tem transcription + meetingAt).
      // Payloads de VoIP/WhatsApp são ignorados nesta primeira fase.
      if (!payload.transcription || !payload.meetingAt) {
        console.log('[Salesbud] Payload não é do tipo Reunião, ignorando.');
        return 'descartada';
      }

      titulo = payload.title || 'Sem título';
      const userId = String(payload.userId || '');
      executivo = payload._executivo || SALESBUD_USER_MAP[userId] || null;
      callId = `sb_${payload.id}`;

      // Filtro 1 — só reunião concluída (status 3)
      if (payload.status !== 3) {
        console.log('[Salesbud] Descartado — status não é concluído:', payload.status, titulo);
        return 'descartada';
      }

      // Filtro 2 — só reunião externa (com cliente) — a Salesbud já classifica isso
      if (payload.isExternal !== true) {
        console.log('[Salesbud] Descartado — reunião interna:', titulo);
        return 'descartada';
      }

      // Filtro 3 — título deve conter padrão Frota162
      const tituloLower = titulo.toLowerCase();
      const ehReuniaoCliente = tituloLower.includes('frota162 ><') || tituloLower.includes('frota162 <>') || tituloLower.includes('frota162><') || tituloLower.includes('frota162<>') || tituloLower.includes('frota 162');
      if (!ehReuniaoCliente) {
        console.log('[Salesbud] Descartado — não é reunião com cliente:', titulo);
        return 'descartada';
      }

      // Filtro 4 — executivo autorizado (o webhook já pode estar filtrado por usuário
      // na própria Salesbud, mas mantemos esta checagem como segunda camada de defesa)
      if (!executivo) {
        console.log('[Salesbud] Descartado — userId não mapeado:', userId, titulo);
        return 'descartada';
      }

      const drive = getDriveClient();

      // Filtro 5 — já processada / corrida de paralelismo (mesma infra de marcadores
      // do fluxo Elephan, mas com callId prefixado "sb_" = isolamento total)
      // Chave única da reunião (título normalizado + minuto): é a MESMA para o webhook e
      // para a sincronização via API, então as duas fontes disputam o mesmo marcador e
      // só uma processa. O marcador legado (sb_<id numérico>) continua valendo para o webhook.
      chaveUnica = chaveReuniao(payload.title, payload.meetingAt);
      if (!payload._origemApi && await isProcessed(drive, callId)) {
        console.log('[Salesbud] Já processada (marcador legado), pulando:', callId);
        return 'ja_processada';
      }
      const devoProcessar = await claimCall(drive, chaveUnica);
      if (!devoProcessar) {
        console.log('[Salesbud] Já processada ou perdeu a corrida, pulando:', callId, chaveUnica);
        return 'ja_processada';
      }

      // Filtro 6 — data da reunião precisa ser HOJE (Brasília UTC-3).
      // meetingAt já vem em ISO — muito mais simples que o parsing que fazíamos com Elephan.
      const offsetBrasilia = 3 * 60;
      const agoraBrasilia = new Date(Date.now() - offsetBrasilia * 60 * 1000);
      const hojeStr = agoraBrasilia.toISOString().slice(0, 10);
      const dObj = new Date(payload.meetingAt);
      let dataCallStr = '', dataCallFormatada = 'Data não informada';
      if (!isNaN(dObj.getTime())) {
        const dBrasilia = new Date(dObj.getTime() - offsetBrasilia * 60 * 1000);
        dataCallStr = dBrasilia.toISOString().slice(0, 10);
        dataCallFormatada = dBrasilia.toISOString().slice(0, 16).replace('T', ' ');
      }
      // Via API a janela já é controlada pela sincronização (48h, a partir de SALESBUD_SYNC_DESDE);
      // o filtro de "hoje" vale só para o webhook, que entrega na hora.
      if (!dataCallStr || (!payload._origemApi && dataCallStr !== hojeStr)) {
        console.log('[Salesbud] Descartado — reunião não é de hoje:', dataCallStr, 'hoje:', hojeStr, titulo);
        return 'descartada';
      }

      // Transcrição: remove HTML, valida tamanho mínimo. Aviso único se curta demais.
      const transcricao = payload.transcriptionIsPlain ? String(payload.transcription).trim() : stripHtml(payload.transcription);
      if (!transcricao || transcricao.length < 500) {
        const jaAvisou = await isMarked(drive, `descartada_${callId}`);
        if (!jaAvisou) {
          await postSlack(`:no_entry_sign: *[Salesbud] Call descartada — ${titulo}* (${executivo}): transcrição ausente ou muito curta.`, process.env.SLACK_WEBHOOK_URL).catch(()=>{});
          await markGeneric(drive, `descartada_${callId}`);
        } else {
          console.log('[Salesbud] Descartado (silencioso, já avisado antes) — transcrição curta:', callId);
        }
        return 'descartada';
      }

      // Salva a transcrição em .txt no Drive — ANTES do Claude, para termos o
      // registro mesmo que a análise ou o PPTX falhem depois. Falha ao salvar
      // não bloqueia o resto do pipeline (só loga e segue).
      let linkTranscricao = '';
      try {
        linkTranscricao = await salvarTranscricaoDrive(drive, titulo, dataCallFormatada, executivo, callId, transcricao);
      } catch(e) {
        console.error('[Salesbud] Falha ao salvar transcrição no Drive (não bloqueante):', e.message);
      }

      // customerName/company já vêm estruturados da Salesbud — passamos como contexto
      // extra pro Claude, complementando (não substituindo) a extração pela transcrição.
      // competitorMentions também vem da Salesbud (lista de nomes) — passamos os nomes
      // já detectados para o Claude descrever O QUE foi dito sobre cada um (sentimento
      // e conteúdo), já que a Salesbud só entrega a lista, não o contexto qualitativo.
      const concorrentes = (payload.context && Array.isArray(payload.context.competitorMentions)) ? payload.context.competitorMentions : [];
      const contextoConcorrencia = concorrentes.length > 0
        ? `\nConcorrentes já identificados nesta call pela Salesbud: ${concorrentes.join(', ')}. Para cada um, descreva no campo concorrencia_detalhe o sentimento do CLIENTE (elogiou/neutro/criticou) e o que especificamente foi dito sobre ele na transcrição. Se o concorrente apenas foi citado, sem avaliação do cliente (ex.: o executivo o mencionou como contexto histórico), o sentimento é neutro e diga que foi só uma menção.\n`
        : '';
      const contextoExtra = `Nome/email do cliente (Salesbud): ${payload.customerName||'não informado'}\nEmpresa/domínio (Salesbud): ${payload.company||'não informado'}${contextoConcorrencia}\n\n`;
      const conteudo = `Título: ${titulo}\nData: ${dataCallFormatada}\nExecutivo Frota162: ${executivo}\n${contextoExtra}Transcrição:\n${transcricao}`;

      let d, lastErr;
      for (let tentativa = 1; tentativa <= 3; tentativa++) {
        try { d = await callClaude(conteudo); break; }
        catch(e) {
          lastErr = e;
          console.log(`[Salesbud] callClaude tentativa ${tentativa} falhou:`, e.message);
          if (tentativa < 3) await new Promise(r => setTimeout(r, 5000 * tentativa));
        }
      }
      if (!d) throw lastErr;

      const empresaValida = d.empresa && d.empresa !== 'Empresa Não Identificada' && d.empresa !== 'Não identificado' && d.empresa !== '';
      if (!empresaValida) {
        console.log('[Salesbud] Descartado — empresa não identificada após análise Claude:', titulo, '| empresa:', d.empresa);
        return 'descartada';
      }

      // Placas não mapeadas NÃO bloqueiam o material: ele sai com tudo que a call trouxe
      // (placas=0 segue a regra de 1 a 10 placas). Só o MRR exige honestidade (ver textoMrr).
      if (!(d.placas > 0) && d.tipo_reuniao !== 'parceria') {
        console.log('[Salesbud] Placas não mapeadas — material gerado mesmo assim:', titulo);
      }

      const empresa = d.empresa || 'Prospect';
      const nomeArq = `Frota162 >< ${sanitizeFileName(empresa)} (Diretoria).pptx`;
      const outPath = path.join(os.tmpdir(), nomeArq);

      const sanitizeObjSb = (obj) => {
        if (typeof obj === 'string') return sanitize(obj);
        if (Array.isArray(obj)) return obj.map(sanitizeObjSb);
        if (obj && typeof obj === 'object') { const r={}; for(const k of Object.keys(obj)) r[k]=sanitizeObjSb(obj[k]); return r; }
        return obj;
      };
      const dClean = sanitizeObjSb(d);
      await gerarPPTX(dClean, outPath);

      const pastaId = process.env.PASTA_RAIZ_ID;
      const uploaded = await drive.files.create({
        supportsAllDrives: true,
        requestBody: { name: nomeArq, parents: [pastaId], mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
        media: { mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', body: fs.createReadStream(outPath) },
        fields: 'id,name,webViewLink',
      });
      await drive.permissions.create({ fileId: uploaded.data.id, supportsAllDrives: true, requestBody: { role: 'writer', type: 'anyone' } });
      fs.unlinkSync(outPath);

      const tempEmoji = d.temperatura==='quente' ? '🔴' : d.temperatura==='morno' ? '🟡' : '🔵';
      const slackId = getSlackId(executivo);
      const execMencao = slackId ? `<@${slackId}>` : (executivo || 'N/A');
      const roiAnual = d.roi_anual || 0;
      const roiTexto = roiAnual > 0 ? `R$${Math.round(roiAnual).toLocaleString('pt-BR')}/ano` : 'A calcular';

      // Ponto 1 (feedback do Bruno): a Salesbud já entrega concorrentes mencionados
      // e um score de qualidade da call — puxamos direto do payload (não precisa do
      // Claude extrair de novo) e só adicionamos a linha quando há dado real.
      // (variável "concorrentes" já foi extraída mais acima, antes de montar o prompt)
      const linhaConcorrentes = concorrentes.length > 0 ? `\n- *Concorrente mencionado:* ${concorrentes.join(', ')}` : '';

      // Detalhe qualitativo (sentimento + o que foi dito) — vem da análise do Claude,
      // que recebeu os nomes dos concorrentes como contexto (ver contextoConcorrencia acima)
      const linhaConcorrenciaDetalhe = concorrentes.length > 0 && d.concorrencia_detalhe ? `\n- *Sobre a concorrência:* ${d.concorrencia_detalhe}` : '';
      // Só aparece quando o detalhe registra ELOGIO a um concorrente (objeção real a contornar) e a reunião é de venda.
      // O modelo já gerou o campo mesmo com "criticou"/"neutro"; por isso a regra também vale aqui, em código.
      const houveElogio = /elogi/i.test(d.concorrencia_detalhe || '');
      const linhaContraArgumento = concorrentes.length > 0 && d.tipo_reuniao !== 'parceria' && houveElogio && d.concorrencia_contra_argumento ? `\n- *Como contornar:* ${d.concorrencia_contra_argumento}` : '';

      // Desvio da matriz de oferta por placas (ex.: executivo ofereceu plano avulso a lead de até 10 placas)
      const linhaAlertaRegra = d.alerta_regra_placas && String(d.alerta_regra_placas).trim() ? `\n- *Alerta de regra:* ${String(d.alerta_regra_placas).trim()}` : '';

      const scoreSalesbud = payload.analytics && payload.analytics.score != null ? payload.analytics.score : null;
      const justificativaScore = payload.analytics && payload.analytics.justification ? payload.analytics.justification : '';
      const linhaScore = scoreSalesbud != null ? `\n- *Score Salesbud:* ${scoreSalesbud}/10` : '';

      const msg = `:car: *[Salesbud] Novo material e análise estratégica* :rocket:\n\n- *Empresa:* ${empresa}\n- *Executivo:* ${execMencao}\n- *Data da reunião:* ${dataCallFormatada}\n- *Placas e MRR estimado:* ${textoPlacas(d)} · ${textoMrr(d)}${linhaAlertaRegra}\n- *ROI estimado:* ${roiTexto}${linhaConcorrentes}${linhaConcorrenciaDetalhe}${linhaContraArgumento}${linhaScore}\n- *Material:* <${uploaded.data.webViewLink}|Abrir PPTX>\n- *Temperatura estimada:* ${tempEmoji} ${d.temperatura||'N/A'}\n- *Resumo Geral da negociação:* ${d.slack_resumo||''}`;

      console.log(`[Salesbud] SUCESSO — titulo:"${titulo}" empresa:"${empresa}" placas:${d.placas} executivo:${executivo}`);
      await postSlack(msg, process.env.SLACK_WEBHOOK_URL);

      // Grava linha no histórico consultável (Google Sheets). Não bloqueia o
      // pipeline se falhar — a mensagem no Slack e o PPTX já foram entregues.
      const tags = (payload.context && Array.isArray(payload.context.tags)) ? payload.context.tags : [];
      try {
        await salvarHistoricoPlanilha([
          dataCallFormatada,
          executivo,
          empresa,
          d.placas || 0,
          d.temperatura || '',
          roiAnual || 0,
          scoreSalesbud != null ? scoreSalesbud : '',
          concorrentes.join(', '),
          tags.join(', '),
          d.perfil_lead || '',
          d.proximo_passo || '',
          uploaded.data.webViewLink,
          linkTranscricao,
          callId,
          d.mrr_citado && String(d.mrr_citado).trim() ? String(d.mrr_citado).trim() : 'não mapeado',
        ]);
      } catch(e) {
        console.error('[Salesbud] Falha ao gravar no histórico da planilha (não bloqueante):', e.message);
      }

      // Só marca sucesso definitivo depois do Slack confirmar entrega
      await markProcessed(drive, callId);
      await markProcessed(drive, chaveUnica);

      // ── Sugestão de Follow-up (módulo isolado, nunca bloqueia o principal) ──
      // Roda DEPOIS do markProcessed: se falhar, a call já está marcada como
      // processada (o slide estratégico já foi entregue) e não deve ser
      // reprocessada — a falha aqui é só logada por gerarEEnviarFollowup.
      if (d.tipo_reuniao !== 'parceria') {
        await gerarEEnviarFollowup({ empresa, executivo, transcricao, dCall: d, postSlack });
      }

      return 'ok';

    } catch(err) {
      console.error('[Salesbud] Background error:', err.message);
      if (chaveUnica) { try { await liberarClaim(getDriveClient(), chaveUnica); } catch (e) {} }
      try {
        await postSlack(`:warning: *[Salesbud] Erro ao gerar material* — ${titulo||'Sem título'} (${executivo||'?'})\nMotivo: ${err.message}`, process.env.SLACK_WEBHOOK_URL);
      } catch(e2) {
        console.error('[Salesbud] Falha ao avisar erro no Slack:', e2.message);
      }
      return 'erro';
    }
}

app.post('/webhook/salesbud', (req, res) => {
  // Responde rápido, processa em background (mesmo padrão de robustez do /generate)
  res.json({ ok: true, status: 'processing' });

  (async () => {
    try {
      const rawBody = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);

      const verificacao = verificaAssinaturaSalesbud(req, rawBody);
      if (!verificacao.ok) {
        console.log('[Salesbud] Webhook REJEITADO -', verificacao.motivo);
        return;
      }

      const payload = JSON.parse(rawBody);

      // LOG INCONDICIONAL — dispara SEMPRE que um webhook chega, antes de qualquer
      // filtro. É a evidência definitiva de "o webhook chegou" independente do que
      // acontecer depois (sucesso, descarte ou erro). Buscar por "RECEBIDO" no log
      // do Render é a forma confiável de confirmar chegada — nunca buscar pelo nome
      // da empresa, que só existe DEPOIS da análise do Claude e nunca era logado.
      console.log(`[Salesbud] RECEBIDO — id:${payload.id} titulo:"${payload.title||''}" userId:${payload.userId} status:${payload.status} isExternal:${payload.isExternal} meetingAt:${payload.meetingAt}`);

      await processarPayloadSalesbud(payload);
    } catch(err) {
      console.error('[Salesbud] Erro ao receber webhook:', err.message);
    }
  })();
});

// ═══════════════════════════════════════════════════════════════════════
// CHECKLIST DIÁRIO — disparado por um Render Cron Job separado, todo dia às
// 22h (horário de Brasília). Mensagem fixa, sem lógica de agregação — só
// lembra o Bruno de comparar Salesbud x Slack no fim do dia. Isolado de
// tudo o resto: não usa Drive, Claude, nem toca nos outros pipelines.
// ═══════════════════════════════════════════════════════════════════════
const CHECKLIST_DIARIO = `📋 *Checklist diário — Salesbud × Slack*

Use isso agora para confirmar que nenhuma call sumiu silenciosamente hoje.

*1. Conta as reuniões "Frota162 ><" ou "Frota162 <>" de hoje na Salesbud*
Para os 6 executivos (Bruno Pereira, Júlio Mazzetti, Palloma Santos, Rávila Silva, Thais Cristina, William Duarte), quantas reuniões concluídas aparecem hoje?
→ Número A: ____

*2. Conta as mensagens "[Salesbud]" no canal oficial hoje*
Quantas mensagens "Novo material e análise estratégica" chegaram no #sales-slides-estrategicos hoje?
→ Número B: ____

*3. Bateu A = B?*
✅ Se sim, dia limpo, não precisa investigar nada.
⚠️ Se não, vai para o passo 4.

*4. Para cada reunião sem mensagem correspondente:*
Render → Logs → busca pelo nome do cliente ou título da reunião.
- Nada com \`RECEBIDO\` → webhook não chegou (verificar do lado da Salesbud)
- \`RECEBIDO\` sem \`SUCESSO\` → descartada por algum filtro (a linha seguinte diz qual)
- \`RECEBIDO\` seguido de \`Background error\` → erro real, investigar`;

app.post('/cron/checklist-diario', (req, res) => {
  res.json({ ok: true });
  postSlack(CHECKLIST_DIARIO, process.env.CHECKLIST_DM_WEBHOOK_URL)
    .then(() => console.log('[Checklist] Enviado com sucesso'))
    .catch(e => console.error('[Checklist] Falha ao enviar:', e.message));
});

// ==== SUPER BRIEFING (HubSpot) — não mexe na lógica do Salesbud/Elephan acima ====
registrarRotaSuperBriefing(app, getDriveClient, claimCall, markProcessed);
registrarRotaPolling(app, getDriveClient, claimCall, markProcessed);

// ==== SINCRONIZAÇÃO VIA API DA SALESBUD (rede de segurança do webhook + gatilho principal) ====
registrarRotaSalesbudSync(app, {
  processar: processarPayloadSalesbud,
  getDriveClient, isProcessed, isMarked, markGeneric, chaveReuniao,
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Frota162 PPTX Server v30 (sync via API Salesbud + gera sempre + parceria + matriz de placas + MRR so o citado + sem slide de custo em branco + follow-up pos-call + PASTA_RAIZ ${process.env.PASTA_RAIZ_ID}) porta ${PORT}`));
