const crypto = require('crypto');
const express = require('express');
const PptxGenJS = require('pptxgenjs');
const { google } = require('googleapis');
const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');

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

const SYSTEM = `Você é especialista em vendas B2B da Frota162. Analise a transcrição e retorne SOMENTE JSON válido sem markdown sem backticks.

TABELA DE PREÇOS (planos: Enterprise 1, Enterprise 2, Enterprise 3 — NUNCA outro nome): Até 40 placas mínimo E1 R$397 E2 R$549 E3 R$649. 41-99 E1 R$9,48 E2 R$14,23 E3 R$16,60. 100-199 E1 R$9,00 E2 R$13,51 E3 R$15,77. 200-299 E1 R$7,65 E2 R$11,49 E3 R$13,41. 300-399 E1 R$7,27 E2 R$10,91 E3 R$12,74. 400-499 E1 R$6,91 E2 R$10,37 E3 R$12,10. 500-999 E1 R$6,56 E2 R$9,85 E3 R$11,49. 1000-1999 E1 R$5,90 E2 R$8,86 E3 R$10,34. CNPJ adicional R$150/mês. Enterprise 1=multas+SNE+1CNPJ. Enterprise 2=E1+IPVA+indicação+3CNPJs. Enterprise 3=E2+CNH+tox+5CNPJs.

REGRA CRÍTICA ≤40 PLACAS (vigente desde treinamento Service): para clientes com até 40 placas, NUNCA oferecer Enterprise 1 ou Enterprise 2. Existem SOMENTE duas opções: (a) Enterprise 3 sozinho R$649/mês — cliente opera a plataforma; (b) Enterprise 3 + Service R$932/mês — a Frota162 opera tudo para o cliente (indicação de condutor, pagamento, gestão documental completa). z3_stat SEMPRE "Enterprise 3" para placas<=40, nunca E1/E2.

SERVICE — quando incluir: se placas<=40 E a call sinalizar qualquer interesse do cliente em não operar a plataforma, em terceirizar, ou em receber as duas propostas (software e software+service), defina tem_interesse_service=true. Posicionamento correto do Service (do treinamento oficial): NÃO é terceirizar para despachante — é BPO documental, a Frota162 assume o processo com especialistas dedicados. Nunca revelar limites internos de quantidade ao cliente (ex: "144 indicações/ano") — dizer apenas "está incluso". Sempre mencionar SNE (desconto por adesão) como pilar central do Service. Benefícios reais a explorar: zero novas contratações, zero curva de aprendizado da equipe do cliente, redução de risco (prazo/indicação/CNH não dependem mais da memória do cliente), tempo da equipe liberado para a operação.

ROI: economia_multas=multas_mes x valor x 0.4(SNE) ou 0.2. economia_NIC=NIC_tratadas x valor x 0.6. economia_pessoas=(func-1) x 2500. ROI_anual=economia_total_anual - investimento_anual. custo_mensal_base=multas_mes x 130 se sem ROI. payback = investimento_mensal / (ROI_anual / 12). SOMENTE calcular dias_payback quando tem_roi=true e ROI_anual > 0. Se tem_roi=false, definir dias_payback=0.

LINGUAGEM: material apresentado pelo executivo Frota162 à diretoria do cliente. Use linguagem voltada ao cliente: "sua frota", "seu time", "sua operação". NÃO linguagem interna da Frota162.

REGRAS: valor mínimo multa R$130. sinal menos SÓ nas barras do custo de esperar. tem_roi=false se call não confirmou. títulos máx 40 chars. z3_stat deve ser Enterprise 1, Enterprise 2 ou Enterprise 3 (ou Enterprise 3 obrigatório se placas<=40) — NUNCA nome inventado como "Plano Intermediário". headers provocativos e específicos para ESTE cliente. NUNCA usar emojis em nenhum campo. z1_stat1 e z1_stat2 devem ser curtos max 12 chars ex: 'R$90k' '30%' '5.000+' nunca frases longas. NUNCA fabricar números de ROI/economia que a call não confirmou — se Gabriel não revelou gasto, tem_roi=false e os campos de custo usam linguagem qualitativa, nunca valor inventado.

TEMPERATURA: quente=lead engajado perguntas próximos passos decisor envolvido. morno=interesse sem comprometimento claro. frio=pouco engajamento objeções sem próximo passo.

JSON (todos obrigatórios; campos service_* só relevantes quando tem_interesse_service=true e placas<=40):
{"empresa":"","perfil_lead":"decisor ou influenciador","placas":0,"cnpjs":0,"segmento":"","tem_roi":true,"tem_interesse_service":false,"temperatura":"quente ou morno ou frio","roi_anual":0,"s1_header_bold":"[Nome do decisor se identificado],\\nvocês têm X placas [situação específica]. Formato: Nome,\\nvocês têm 22 placas rodando SP sem visibilidade. Se sem nome: frase provocativa com dado real max 70 chars","s1_header_sub":"X placas · Y CNPJs · Região","s1_subtitulo":"contexto segmento voltado ao cliente","cards":[{"stat":"","titulo":"max 40 chars","desc":"2-3 linhas específicas voltadas ao cliente"},{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""}],"s1_footer_bold":"urgência específica com número real max 80 chars","s1_footer_normal":"complemento","service_header_bold":"frase conceitual SEM valores em R$, ex: 'A Frota162 também pode operar tudo isso para você.'","service_header_sub":"Enterprise 3 + Service - disponível para frotas até 40 placas","service_sem_titulo":"SEM SERVICE - você opera","service_sem_itens":["item 1 sem service","item 2","item 3","item 4"],"service_com_titulo":"COM SERVICE - a Frota162 opera","service_com_itens":["item 1 com service orientado a resultado","item 2","item 3","item 4"],"service_beneficios":[{"stat":"0","titulo":"Novas contratações","desc":"curto"},{"stat":"Menos","titulo":"Tempo da equipe","desc":"curto"},{"stat":"100%","titulo":"Do risco sai da mão","desc":"curto"}],"service_nota":"reforço qualitativo sem valores em R$","s2_header_bold":"Da dor de -R$X ao retorno de +R$Y por ano. OU X placas sem visibilidade a recomendacao é o Enterprise 3.","s2_header_normal":"Como a Frota162 resolve, em 3 passos — [Empresa]","z1_stat1":"","z1_sub1":"1 linha","z1_stat2":"","z1_sub2":"1 linha","z1_bullets":["dado específico 1","dado específico 2"],"passos":[{"titulo":"max 35 chars voltado ao cliente","desc":"1 linha no contexto do cliente"},{"titulo":"","desc":""},{"titulo":"","desc":""},{"titulo":"","desc":""}],"z3_stat":"ROI anual R$X OU Enterprise 1 OU Enterprise 2 OU Enterprise 3","z3_sub1":"retorno por ano ou features do plano","z3_investimento":"R$X,XX/mês","z3_badge":"diferencial específico para este cliente","z3_service_label":"Com Service - a Frota162 opera por você (só se tem_interesse_service)","z3_service_investimento":"R$X,XX/mês (só se tem_interesse_service)","z3_service_tagline":"Menos gente, tempo e risco - por só R$X a mais (só se tem_interesse_service, X = diferença real E3+Service menos E3)","z3_nota":"condições comerciais e limitações reais","s2_cta_bold":"próximo passo combinado na call","s2_cta_normal":"ação concreta","custo_mensal":0,"investimento_mensal_num":0,"dias_payback":0,"s3_header_bold":"Cada mês sem a Frota162 é R$X saindo do caixa da [Empresa].","s3_header_sub":"base confirmada na call","s3_formula":"fórmula usada","s3_nota":"metodologia e limitações","slack_resumo":"2 linhas objetivas SEM EMOJIS: dor principal + alerta crucial","proximo_passo":"ação concreta para o executivo"}`;

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

  // ── SLIDE 2 — SERVICE: SEM x COM (conceitual, sem valores em R$) ────────
  // Só é gerado quando placas <= 40 E há sinal de interesse no Service
  if (d.tem_interesse_service && (d.placas||999) <= 40) {
    const s1s = pres.addSlide();
    s1s.background = { color: 'F7F6F4' };
    s1s.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:0.92,fill:{color:COR.laranja}});
    s1s.addText(d.service_header_bold||'',{x:0.38,y:0.05,w:9.3,h:0.58,fontFace:'Montserrat',fontSize:14,bold:true,color:COR.branco,valign:'middle',margin:0,wrap:true});
    s1s.addText(d.service_header_sub||'',{x:0.38,y:0.66,w:9.3,h:0.24,fontFace:'Montserrat',fontSize:10,color:'FFD0C0',valign:'middle',margin:0});

    const SCY = 0.98, SCH = 2.75;
    const SC1X=0.30, SCW=4.55, SC2X=5.15;

    // Coluna SEM SERVICE (neutro, sem borda)
    s1s.addShape(pres.ShapeType.rect,{x:SC1X,y:SCY,w:SCW,h:SCH,fill:{color:'F0EFED'}});
    s1s.addShape(pres.ShapeType.rect,{x:SC1X,y:SCY,w:SCW,h:0.42,fill:{color:'DDDBD8'}});
    s1s.addText(d.service_sem_titulo||'SEM SERVICE',{x:SC1X+0.16,y:SCY,w:SCW-0.32,h:0.42,fontFace:'Montserrat',fontSize:10,bold:true,color:COR.dark,valign:'middle',margin:0});
    (d.service_sem_itens||[]).forEach((item,i)=>{
      const iy = SCY+0.52+i*0.55;
      s1s.addShape(pres.ShapeType.ellipse,{x:SC1X+0.18,y:iy+0.02,w:0.16,h:0.16,fill:{color:'999999'}});
      s1s.addText(item,{x:SC1X+0.44,y:iy-0.06,w:SCW-0.62,h:0.48,fontFace:'Montserrat',fontSize:8,color:'444444',valign:'top',margin:0,wrap:true});
    });

    // Coluna COM SERVICE (verde, sem borda)
    s1s.addShape(pres.ShapeType.rect,{x:SC2X,y:SCY,w:SCW,h:SCH,fill:{color:'EAF6EA'}});
    s1s.addShape(pres.ShapeType.rect,{x:SC2X,y:SCY,w:SCW,h:0.42,fill:{color:COR.verde}});
    s1s.addText(d.service_com_titulo||'COM SERVICE',{x:SC2X+0.16,y:SCY,w:SCW-0.32,h:0.42,fontFace:'Montserrat',fontSize:10,bold:true,color:COR.branco,valign:'middle',margin:0});
    (d.service_com_itens||[]).forEach((item,i)=>{
      const iy = SCY+0.52+i*0.55;
      s1s.addShape(pres.ShapeType.ellipse,{x:SC2X+0.18,y:iy+0.02,w:0.16,h:0.16,fill:{color:COR.verde}});
      s1s.addText(item,{x:SC2X+0.44,y:iy-0.06,w:SCW-0.62,h:0.48,fontFace:'Montserrat',fontSize:8,bold:true,color:COR.dark,valign:'top',margin:0,wrap:true});
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
  s2.addText('COMO A FROTA162 RESOLVE',{x:Z2X+0.18,y:CY+0.16,w:Z2W-0.24,h:0.22,fontFace:'Montserrat',fontSize:8.5,bold:true,color:COR.dark,charSpacing:0.5,margin:0});

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
  s2.addText(d.tem_roi?'RESULTADO':'PLANO RECOMENDADO',{x:Z3X+0.18,y:CY+0.16,w:Z3W-0.24,h:0.22,fontFace:'Montserrat',fontSize:9,bold:true,color:COR.verde,charSpacing:1,margin:0});

  // z3_stat — fonte adaptativa
  const z3Sz = (d.z3_stat||'').length > 8 ? 16 : 26;
  s2.addText(d.z3_stat||'',{x:Z3X+0.18,y:CY+0.40,w:Z3W-0.24,h:0.50,fontFace:'Montserrat',fontSize:z3Sz,bold:true,color:COR.verde,margin:0,wrap:true});
  s2.addText(d.z3_sub1||'',{x:Z3X+0.18,y:CY+0.92,w:Z3W-0.24,h:0.30,fontFace:'Montserrat',fontSize:8,color:'555555',margin:0,wrap:true});
  s2.addShape(pres.ShapeType.rect,{x:Z3X+0.18,y:CY+1.28,w:Z3W-0.36,h:0.016,fill:{color:'BFE3BF'}});
  s2.addText(d.z3_investimento||'',{x:Z3X+0.18,y:CY+1.34,w:Z3W-0.24,h:0.36,fontFace:'Montserrat',fontSize:16,bold:true,color:COR.verde,margin:0,wrap:true});

  // Badge diferencial
  s2.addShape(pres.ShapeType.rect,{x:Z3X+0.18,y:CY+1.82,w:Z3W-0.36,h:0.36,fill:{color:'D4EED4'},line:{color:'BFE3BF',width:0.5}});
  s2.addText(d.z3_badge||'',{x:Z3X+0.22,y:CY+1.82,w:Z3W-0.44,h:0.36,fontFace:'Montserrat',fontSize:7.5,bold:true,color:COR.verde,align:'center',valign:'middle',margin:0,wrap:true});

  // Zona 3: Service (prioridade quando aplicável) OU Payback — nunca os dois, evita colisão visual
  if (d.tem_interesse_service && (d.placas||999) <= 40 && d.z3_service_investimento) {
    // Linha de Service - já apresentado no slide anterior, aqui só reforça a opção
    s2.addShape(pres.ShapeType.rect,{x:Z3X+0.18,y:CY+2.32,w:Z3W-0.36,h:0.016,fill:{color:'BFE3BF'}});
    s2.addText(d.z3_service_label||'Com Service - a Frota162 opera por você',{x:Z3X+0.18,y:CY+2.40,w:Z3W-0.24,h:0.22,fontFace:'Montserrat',fontSize:7.5,bold:true,color:'558855',margin:0,wrap:true});
    s2.addText(d.z3_service_investimento,{x:Z3X+0.18,y:CY+2.62,w:Z3W-0.24,h:0.32,fontFace:'Montserrat',fontSize:14,bold:true,color:COR.verde,margin:0,wrap:true});
    if (d.z3_service_tagline) {
      s2.addText(d.z3_service_tagline,{x:Z3X+0.18,y:CY+2.96,w:Z3W-0.24,h:0.30,fontFace:'Montserrat',fontSize:7.5,bold:true,italic:true,color:COR.laranja,margin:0,wrap:true});
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
  const s3 = pres.addSlide();
  s3.background = { color: 'F7F6F4' };

  // Header laranja
  s3.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:1.00,fill:{color:COR.laranja}});
  s3.addText(d.s3_header_bold||'',{x:0.38,y:0.05,w:9.3,h:0.58,fontFace:'Montserrat',fontSize:14,bold:true,color:COR.branco,valign:'middle',margin:0,wrap:true});
  s3.addText(d.s3_header_sub||'',{x:0.38,y:0.65,w:9.3,h:0.24,fontFace:'Montserrat',fontSize:9,color:'FFD0C0',valign:'middle',margin:0});
  s3.addText(d.s3_formula||'',{x:0.38,y:1.02,w:9.3,h:0.20,fontFace:'Montserrat',fontSize:7.5,italic:true,color:'888888',margin:0,wrap:true});

  // 6 barras: 10, 20, 30, 45, 60, 90 dias
  const cm = d.custo_mensal||0;
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

  return pres.writeFile({ fileName: outPath });
}

// Remove emojis e chars especiais que quebram pptxgenjs
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
      const nomeArq = `Frota162 >< ${empresa} (Diretoria).pptx`;
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
      const msg = `:car: *Novo material e análise estratégica* :rocket:\n\n- *Empresa:* ${empresa}\n- *Executivo:* ${execMencao}\n- *Data da reunião:* ${dataHoraReuniao}\n- *Placas e MRR estimado:* ${d.placas||0} placas · ${d.z3_investimento||'A definir'}\n- *ROI estimado:* ${roiTexto}\n- *Material:* <${uploaded.data.webViewLink}|Abrir PPTX>\n- *Temperatura estimada:* ${tempEmoji} ${d.temperatura||'N/A'}\n- *Resumo Geral da negociação:* ${d.slack_resumo||''}`;

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

app.post('/webhook/salesbud', (req, res) => {
  // Responde rápido, processa em background (mesmo padrão de robustez do /generate)
  res.json({ ok: true, status: 'processing' });

  (async () => {
    let titulo, executivo, callId;
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

      // Só processamos o payload de "Reunião" (tem transcription + meetingAt).
      // Payloads de VoIP/WhatsApp são ignorados nesta primeira fase.
      if (!payload.transcription || !payload.meetingAt) {
        console.log('[Salesbud] Payload não é do tipo Reunião, ignorando.');
        return;
      }

      titulo = payload.title || 'Sem título';
      const userId = String(payload.userId || '');
      executivo = SALESBUD_USER_MAP[userId] || null;
      callId = `sb_${payload.id}`;

      // Filtro 1 — só reunião concluída (status 3)
      if (payload.status !== 3) {
        console.log('[Salesbud] Descartado — status não é concluído:', payload.status, titulo);
        return;
      }

      // Filtro 2 — só reunião externa (com cliente) — a Salesbud já classifica isso
      if (payload.isExternal !== true) {
        console.log('[Salesbud] Descartado — reunião interna:', titulo);
        return;
      }

      // Filtro 3 — título deve conter padrão Frota162
      const tituloLower = titulo.toLowerCase();
      const ehReuniaoCliente = tituloLower.includes('frota162 ><') || tituloLower.includes('frota162 <>') || tituloLower.includes('frota162><') || tituloLower.includes('frota162<>') || tituloLower.includes('frota 162');
      if (!ehReuniaoCliente) {
        console.log('[Salesbud] Descartado — não é reunião com cliente:', titulo);
        return;
      }

      // Filtro 4 — executivo autorizado (o webhook já pode estar filtrado por usuário
      // na própria Salesbud, mas mantemos esta checagem como segunda camada de defesa)
      if (!executivo) {
        console.log('[Salesbud] Descartado — userId não mapeado:', userId, titulo);
        return;
      }

      const drive = getDriveClient();

      // Filtro 5 — já processada / corrida de paralelismo (mesma infra de marcadores
      // do fluxo Elephan, mas com callId prefixado "sb_" = isolamento total)
      const devoProcessar = await claimCall(drive, callId);
      if (!devoProcessar) {
        console.log('[Salesbud] Já processada ou perdeu a corrida, pulando:', callId);
        return;
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
      if (!dataCallStr || dataCallStr !== hojeStr) {
        console.log('[Salesbud] Descartado — reunião não é de hoje:', dataCallStr, 'hoje:', hojeStr, titulo);
        return;
      }

      // Transcrição: remove HTML, valida tamanho mínimo. Aviso único se curta demais.
      const transcricao = stripHtml(payload.transcription);
      if (!transcricao || transcricao.length < 500) {
        const jaAvisou = await isMarked(drive, `descartada_${callId}`);
        if (!jaAvisou) {
          await postSlack(`:no_entry_sign: *[Salesbud] Call descartada — ${titulo}* (${executivo}): transcrição ausente ou muito curta.`, process.env.SLACK_WEBHOOK_URL).catch(()=>{});
          await markGeneric(drive, `descartada_${callId}`);
        } else {
          console.log('[Salesbud] Descartado (silencioso, já avisado antes) — transcrição curta:', callId);
        }
        return;
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
      const contextoExtra = `Nome/email do cliente (Salesbud): ${payload.customerName||'não informado'}\nEmpresa/domínio (Salesbud): ${payload.company||'não informado'}\n\n`;
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
      const placasValidas = d.placas && d.placas > 0;
      if (!empresaValida || !placasValidas) {
        console.log('[Salesbud] Descartado — campos insuficientes após análise Claude:', titulo, '| empresa:', d.empresa, '| placas:', d.placas);
        return;
      }

      const empresa = d.empresa || 'Prospect';
      const nomeArq = `Frota162 >< ${empresa} (Diretoria).pptx`;
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
      const concorrentes = (payload.context && Array.isArray(payload.context.competitorMentions)) ? payload.context.competitorMentions : [];
      const linhaConcorrentes = concorrentes.length > 0 ? `\n- *Concorrente mencionado:* ${concorrentes.join(', ')}` : '';

      const scoreSalesbud = payload.analytics && payload.analytics.score != null ? payload.analytics.score : null;
      const justificativaScore = payload.analytics && payload.analytics.justification ? payload.analytics.justification : '';
      const linhaScore = scoreSalesbud != null ? `\n- *Score Salesbud:* ${scoreSalesbud}/10` : '';

      const msg = `:car: *[Salesbud] Novo material e análise estratégica* :rocket:\n\n- *Empresa:* ${empresa}\n- *Executivo:* ${execMencao}\n- *Data da reunião:* ${dataCallFormatada}\n- *Placas e MRR estimado:* ${d.placas||0} placas · ${d.z3_investimento||'A definir'}\n- *ROI estimado:* ${roiTexto}${linhaConcorrentes}${linhaScore}\n- *Material:* <${uploaded.data.webViewLink}|Abrir PPTX>\n- *Temperatura estimada:* ${tempEmoji} ${d.temperatura||'N/A'}\n- *Resumo Geral da negociação:* ${d.slack_resumo||''}`;

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
        ]);
      } catch(e) {
        console.error('[Salesbud] Falha ao gravar no histórico da planilha (não bloqueante):', e.message);
      }

      // Só marca sucesso definitivo depois do Slack confirmar entrega
      await markProcessed(drive, callId);

    } catch(err) {
      console.error('[Salesbud] Background error:', err.message);
      try {
        await postSlack(`:warning: *[Salesbud] Erro ao gerar material* — ${titulo||'Sem título'} (${executivo||'?'})\nMotivo: ${err.message}`, process.env.SLACK_WEBHOOK_URL);
      } catch(e2) {
        console.error('[Salesbud] Falha ao avisar erro no Slack:', e2.message);
      }
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Frota162 PPTX Server v15 (Service para ate 40 placas + slide dedicado + PASTA_RAIZ ${process.env.PASTA_RAIZ_ID}) porta ${PORT}`));
