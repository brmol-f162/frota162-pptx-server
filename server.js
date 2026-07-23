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

ROI: economia_multas=multas_mes x valor x 0.4(SNE) ou 0.2. economia_NIC=NIC_tratadas x valor x 0.6. economia_pessoas=(func-1) x 2500. ROI_anual=economia_total_anual - investimento_anual. custo_mensal_base=multas_mes x 130 se sem ROI. payback=30 x (investimento_mes / custo_mensal). SEMPRE calcular dias_payback — usar custo_mensal_base quando sem ROI confirmado.

LINGUAGEM: material apresentado pelo executivo Frota162 à diretoria do cliente. Use linguagem voltada ao cliente: "sua frota", "seu time", "sua operação". NÃO linguagem interna da Frota162.

REGRAS: valor mínimo multa R$130. sinal menos SÓ nas barras slide 3. tem_roi=false se call não confirmou. títulos máx 40 chars. z3_stat deve ser Enterprise 1, Enterprise 2 ou Enterprise 3 quando sem ROI — NUNCA nome inventado. headers provocativos e específicos para ESTE cliente.

TEMPERATURA: quente=lead engajado perguntas próximos passos decisor envolvido. morno=interesse sem comprometimento claro. frio=pouco engajamento objeções sem próximo passo.

JSON (todos obrigatórios):
{"empresa":"","executivo":"","perfil_lead":"decisor ou influenciador","placas":0,"cnpjs":0,"segmento":"","tem_roi":true,"temperatura":"quente ou morno ou frio","roi_anual":0,"s1_header_bold":"[Nome do decisor se identificado],\\nvocês têm X placas [situação específica]. Formato: Nome,\\nvocês têm 22 placas rodando SP sem visibilidade. Se sem nome: frase provocativa com dado real max 70 chars","s1_header_sub":"X placas · Y CNPJs · Região","s1_subtitulo":"contexto segmento voltado ao cliente","cards":[{"stat":"","titulo":"max 40 chars","desc":"2-3 linhas específicas voltadas ao cliente"},{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""}],"s1_footer_bold":"urgência específica com número real max 80 chars","s1_footer_normal":"complemento","s2_header_bold":"Da dor de -R$X ao retorno de +R$Y por ano. OU X placas sem visibilidade a recomendacao é o Enterprise N.","s2_header_normal":"Como a Frota162 resolve, em 3 passos — [Empresa]","z1_stat1":"","z1_sub1":"1 linha","z1_stat2":"","z1_sub2":"1 linha","z1_bullets":["dado específico 1","dado específico 2"],"passos":[{"titulo":"max 35 chars voltado ao cliente","desc":"1 linha no contexto do cliente"},{"titulo":"","desc":""},{"titulo":"","desc":""},{"titulo":"","desc":""}],"z3_stat":"ROI anual R$X OU Enterprise 1 OU Enterprise 2 OU Enterprise 3","z3_sub1":"retorno por ano ou features do plano","z3_investimento":"R$X,XX/mês","z3_badge":"diferencial específico para este cliente","z3_nota":"condições comerciais e limitações reais","s2_cta_bold":"próximo passo combinado na call","s2_cta_normal":"ação concreta","custo_mensal":0,"investimento_mensal_num":0,"dias_payback":0,"s3_header_bold":"Cada mês sem a Frota162 é R$X saindo do caixa da [Empresa].","s3_header_sub":"base confirmada na call","s3_formula":"fórmula usada","s3_nota":"metodologia e limitações","slack_resumo":"2 linhas objetivas SEM EMOJIS: dor principal + alerta crucial","proximo_passo":"ação concreta para o executivo"}`;

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

function postSlack(msg) {
  return new Promise((resolve) => {
    const body = JSON.stringify({ text: msg });
    const url = new URL(process.env.SLACK_WEBHOOK_URL);
    const req = https.request({
      hostname: url.hostname, path: url.pathname + url.search, method: 'POST',
      headers: { 'content-type': 'application/json', 'content-length': Buffer.byteLength(body) }
    }, (res) => { res.on('data', ()=>{}); res.on('end', resolve); });
    req.on('error', () => resolve());
    req.write(body); req.end();
  });
}

function gerarPPTX(d, outPath) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9'; // 10 x 5.625

  // ── SLIDE 1 — Dor em linguagem executiva ──────────────────────────────
  const s1 = pres.addSlide();
  s1.background = { color: 'F7F6F4' };

  // Header laranja
  s1.addShape(pres.ShapeType.rect, {x:0,y:0,w:10,h:1.00,fill:{color:COR.laranja}});
  s1.addText(d.s1_header_bold||'', {x:0.38,y:0.06,w:9.3,h:0.55,fontFace:'Montserrat',fontSize:15,bold:true,color:COR.branco,valign:'middle',margin:0});
  s1.addText(d.s1_header_sub||'', {x:0.38,y:0.62,w:9.3,h:0.28,fontFace:'Montserrat',fontSize:10,color:'FFD0C0',valign:'middle',margin:0});
  s1.addText(d.s1_subtitulo||'', {x:0.38,y:1.02,w:9.3,h:0.20,fontFace:'Montserrat',fontSize:8,color:COR.cinza,italic:true,margin:0});

  // 4 cards 2x2 — maiores e mais ricos
  const CPOS=[{cx:0.20,cy:1.30},{cx:5.10,cy:1.30},{cx:0.20,cy:3.22},{cx:5.10,cy:3.22}];
  const CW=4.70, CH=1.80;
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
    // Stat
    s1.addText(c.stat||'',{x:cx+0.16,y:cy+0.08,w:CW-0.24,h:0.46,fontFace:'Montserrat',fontSize:24,bold:true,color:p.stat,margin:0});
    // Título
    s1.addText(c.titulo||'',{x:cx+0.16,y:cy+0.54,w:CW-0.24,h:0.26,fontFace:'Montserrat',fontSize:9.5,bold:true,color:COR.dark,margin:0});
    // Descrição — mais espaço
    s1.addText(c.desc||'',{x:cx+0.16,y:cy+0.80,w:CW-0.24,h:0.92,fontFace:'Montserrat',fontSize:8,color:'333333',valign:'top',margin:0,wrap:true});
  });

  // Footer dark
  s1.addShape(pres.ShapeType.rect,{x:0,y:5.10,w:10,h:0.525,fill:{color:COR.dark}});
  s1.addShape(pres.ShapeType.rect,{x:0,y:5.10,w:0.05,h:0.525,fill:{color:COR.laranja}});
  s1.addText([
    {text:(d.s1_footer_bold||'')+' ',options:{bold:true,color:COR.laranja}},
    {text:d.s1_footer_normal||'',options:{bold:false,color:COR.branco}}
  ],{x:0.22,y:5.12,w:7.80,h:0.44,fontFace:'Montserrat',fontSize:9.5,valign:'middle',margin:0});
  s1.addText('frota162.com.br',{x:8.20,y:5.18,w:1.65,h:0.32,fontFace:'Montserrat',fontSize:8.5,bold:true,color:COR.laranja,align:'right',valign:'middle',margin:0});

  // ── SLIDE 2 — 3 Zonas ─────────────────────────────────────────────────
  const s2 = pres.addSlide();
  s2.background = { color: 'F7F6F4' };

  s2.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:0.84,fill:{color:COR.dark}});
  s2.addShape(pres.ShapeType.rect,{x:0,y:0,w:0.05,h:0.84,fill:{color:COR.laranja}});
  s2.addText(d.s2_header_bold||'',{x:0.22,y:0.04,w:9.3,h:0.44,fontFace:'Montserrat',fontSize:12,bold:true,color:COR.laranja,valign:'middle',margin:0,wrap:true});
  s2.addText(d.s2_header_normal||'',{x:0.22,y:0.50,w:9.3,h:0.28,fontFace:'Montserrat',fontSize:9,color:COR.branco,valign:'middle',margin:0});

  const CY=0.84, CH2=4.16;
  const Z1X=0,Z1W=3.10,Z2X=3.10,Z2W=3.80,Z3X=6.90,Z3W=3.10;

  // Zona 1 HOJE
  s2.addShape(pres.ShapeType.rect,{x:Z1X,y:CY,w:Z1W,h:CH2,fill:{color:'FFEDE7'}});
  s2.addText('HOJE',{x:Z1X+0.16,y:CY+0.16,w:Z1W-0.22,h:0.22,fontFace:'Montserrat',fontSize:9,bold:true,color:COR.vermelho,charSpacing:1,margin:0});
  s2.addText(d.z1_stat1||'',{x:Z1X+0.16,y:CY+0.40,w:Z1W-0.22,h:0.46,fontFace:'Montserrat',fontSize:26,bold:true,color:COR.vermelho,margin:0});
  s2.addText(d.z1_sub1||'',{x:Z1X+0.16,y:CY+0.88,w:Z1W-0.22,h:0.30,fontFace:'Montserrat',fontSize:8,color:'555555',margin:0,wrap:true});
  s2.addShape(pres.ShapeType.rect,{x:Z1X+0.16,y:CY+1.24,w:Z1W-0.32,h:0.016,fill:{color:'F0B8A5'}});
  s2.addText(d.z1_stat2||'',{x:Z1X+0.16,y:CY+1.30,w:Z1W-0.22,h:0.36,fontFace:'Montserrat',fontSize:18,bold:true,color:COR.vermelho,margin:0});
  s2.addText(d.z1_sub2||'',{x:Z1X+0.16,y:CY+1.70,w:Z1W-0.22,h:0.30,fontFace:'Montserrat',fontSize:8,color:'555555',margin:0,wrap:true});
  (d.z1_bullets||[]).forEach((b,i)=>{
    s2.addText('· '+b,{x:Z1X+0.16,y:CY+2.10+(i*0.28),w:Z1W-0.22,h:0.26,fontFace:'Montserrat',fontSize:7.5,color:'444444',margin:0,wrap:true});
  });

  // Setas removidas por design

  // Zona 2 COMO RESOLVE
  s2.addShape(pres.ShapeType.rect,{x:Z2X,y:CY,w:Z2W,h:CH2,fill:{color:'FCFCFB'}});
  s2.addText('COMO A FROTA162 RESOLVE',{x:Z2X+0.16,y:CY+0.16,w:Z2W-0.22,h:0.22,fontFace:'Montserrat',fontSize:8.5,bold:true,color:COR.dark,charSpacing:0.5,margin:0});
  const SY=CY+0.50,SH=0.72,SG=0.12,rx=Z2X+0.36;
  s2.addShape(pres.ShapeType.rect,{x:rx+0.16,y:SY+0.20,w:0.016,h:(SH+SG)*3+0.20,fill:{color:COR.divisor}});
  (d.passos||[]).forEach((p,i)=>{
    const py=SY+i*(SH+SG);
    s2.addShape(pres.ShapeType.ellipse,{x:rx,y:py+0.06,w:0.38,h:0.38,fill:{color:COR.laranja}});
    s2.addText(String(i+1),{x:rx,y:py+0.06,w:0.38,h:0.38,fontFace:'Montserrat',fontSize:11,bold:true,color:COR.branco,align:'center',valign:'middle',margin:0});
    s2.addText(p.titulo||'',{x:Z2X+0.82,y:py+0.06,w:Z2W-0.94,h:0.28,fontFace:'Montserrat',fontSize:9,bold:true,color:COR.dark,margin:0,wrap:true});
    s2.addText(p.desc||'',{x:Z2X+0.82,y:py+0.36,w:Z2W-0.94,h:0.34,fontFace:'Montserrat',fontSize:8,color:'555555',valign:'top',margin:0,wrap:true});
  });

  // Zona 3 RESULTADO
  s2.addShape(pres.ShapeType.rect,{x:Z3X,y:CY,w:Z3W,h:CH2,fill:{color:'EAF6EA'}});
  s2.addText(d.tem_roi?'RESULTADO':'PLANO RECOMENDADO',{x:Z3X+0.16,y:CY+0.16,w:Z3W-0.22,h:0.22,fontFace:'Montserrat',fontSize:9,bold:true,color:COR.verde,charSpacing:1,margin:0});
  // Ajusta fonte do z3_stat: planos (Enterprise N) têm texto mais longo
  const z3FontSize = (d.z3_stat||'').length > 8 ? 16 : 26;
  s2.addText(d.z3_stat||'',{x:Z3X+0.16,y:CY+0.40,w:Z3W-0.22,h:0.46,fontFace:'Montserrat',fontSize:z3FontSize,bold:true,color:COR.verde,margin:0,wrap:true});
  s2.addText(d.z3_sub1||'',{x:Z3X+0.16,y:CY+0.88,w:Z3W-0.22,h:0.30,fontFace:'Montserrat',fontSize:8,color:'555555',margin:0,wrap:true});
  s2.addShape(pres.ShapeType.rect,{x:Z3X+0.16,y:CY+1.24,w:Z3W-0.32,h:0.016,fill:{color:'BFE3BF'}});
  s2.addText(d.z3_investimento||'',{x:Z3X+0.16,y:CY+1.30,w:Z3W-0.22,h:0.36,fontFace:'Montserrat',fontSize:16,bold:true,color:COR.verde,margin:0,wrap:true});
  s2.addShape(pres.ShapeType.rect,{x:Z3X+0.16,y:CY+1.80,w:Z3W-0.32,h:0.38,fill:{color:'D4EED4'},line:{color:'BFE3BF',width:0.5}});
  s2.addText(d.z3_badge||'',{x:Z3X+0.20,y:CY+1.80,w:Z3W-0.40,h:0.38,fontFace:'Montserrat',fontSize:7.5,bold:true,color:COR.verde,align:'center',valign:'middle',margin:0,wrap:true});

  // Payback na Zona 3 se disponível
  const dp = d.dias_payback||0;
  if (dp > 0 && dp <= 90) {
    const pbLabel = dp <= 45 ? `Payback em ~${dp} dias` : `Payback em ~${Math.round(dp/30)} meses`;
    s2.addShape(pres.ShapeType.rect,{x:Z3X+0.16,y:CY+2.28,w:Z3W-0.32,h:0.38,fill:{color:'FFFFFF'},line:{color:COR.verde,width:1.0}});
    s2.addText(`↑ ${pbLabel}`,{x:Z3X+0.20,y:CY+2.28,w:Z3W-0.40,h:0.38,fontFace:'Montserrat',fontSize:8.5,bold:true,color:COR.verde,align:'center',valign:'middle',margin:0});
  }

  s2.addText(d.z3_nota||'',{x:Z3X+0.16,y:CY+3.30,w:Z3W-0.22,h:0.60,fontFace:'Montserrat',fontSize:6.5,italic:true,color:'AAAAAA',valign:'top',margin:0,wrap:true});

  // Banner final
  s2.addShape(pres.ShapeType.rect,{x:0,y:5.20,w:10,h:0.425,fill:{color:COR.dark}});
  s2.addShape(pres.ShapeType.rect,{x:0,y:5.20,w:0.05,h:0.425,fill:{color:COR.laranja}});
  s2.addText([
    {text:(d.s2_cta_bold||'')+'  ',options:{bold:true,color:COR.laranja}},
    {text:d.s2_cta_normal||'',options:{bold:false,color:COR.branco}}
  ],{x:0.22,y:5.22,w:7.80,h:0.36,fontFace:'Montserrat',fontSize:8.5,valign:'middle',margin:0});
  s2.addText('frota162.com.br',{x:8.20,y:5.26,w:1.65,h:0.28,fontFace:'Montserrat',fontSize:8,bold:true,color:COR.laranja,align:'right',valign:'middle',margin:0});

  // ── SLIDE 3 — O custo de esperar ──────────────────────────────────────
  const s3 = pres.addSlide();
  s3.background = { color: 'F7F6F4' };

  s3.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:1.00,fill:{color:COR.laranja}});
  s3.addText(d.s3_header_bold||'',{x:0.38,y:0.06,w:9.3,h:0.56,fontFace:'Montserrat',fontSize:15,bold:true,color:COR.branco,valign:'middle',margin:0,wrap:true});
  s3.addText(d.s3_header_sub||'',{x:0.38,y:0.64,w:9.3,h:0.26,fontFace:'Montserrat',fontSize:9.5,color:'FFD0C0',valign:'middle',margin:0});
  s3.addText(d.s3_formula||'',{x:0.38,y:1.04,w:9.3,h:0.20,fontFace:'Montserrat',fontSize:8,color:COR.cinza,italic:true,margin:0});

  const cm=d.custo_mensal||0;
  // 6 barras: 10, 20, 30, 45, 60, 90 dias
  const DIAS=[10,20,30,45,60,90];
  const VALS=DIAS.map(d=>cm*(d/30));
  const BY=3.90, MH=1.80, BW=1.40, BG=0.17, X0=0.35;

  DIAS.forEach((dia,i)=>{
    const bH=MH*(dia/90), bTop=BY-bH, bx=X0+i*(BW+BG);
    s3.addShape(pres.ShapeType.rect,{x:bx,y:bTop,w:BW,h:bH,fill:{color:'FFCDD2'},line:{color:COR.vermelho,width:0.5}});
    s3.addText(`-R$${Math.round(VALS[i]).toLocaleString('pt-BR')}`,{x:bx,y:bTop-0.32,w:BW,h:0.28,fontFace:'Montserrat',fontSize:10,bold:true,color:COR.vermelho,align:'center',margin:0});
    s3.addText(`${dia}d`,{x:bx,y:BY+0.08,w:BW,h:0.22,fontFace:'Montserrat',fontSize:8.5,bold:true,color:'555555',align:'center',margin:0});
  });

  s3.addShape(pres.ShapeType.rect,{x:0.20,y:BY,w:9.60,h:0.014,fill:{color:'CCCCCC'}});

  // Payback — sempre posicionado no TOPO (y fixo = 1.32) para não sobrepor barras
  const inv=d.investimento_mensal_num||0;
  const dpCalc = (dp > 0) ? dp : (inv > 0 && cm > 0 ? Math.round(30*(inv/cm)) : 0);
  if(dpCalc > 0){
    // Ancora na barra mais próxima
    let bi=DIAS.findIndex(x=>x>=dpCalc); if(bi<0) bi=5;
    const pbx=X0+bi*(BW+BG);
    const pbLabel = dpCalc<=45 ? `↑ Payback: ~${dpCalc} dias` : `↑ Payback: ~${Math.round(dpCalc/30)} meses`;
    // Posição fixa no topo — nunca sobrepõe barra
    const PAYBACK_Y = 1.28;
    s3.addShape(pres.ShapeType.rect,{x:pbx+0.08,y:PAYBACK_Y,w:BW-0.16,h:0.32,fill:{color:COR.branco},line:{color:COR.verde,width:1.5}});
    s3.addText(pbLabel,{x:pbx+0.08,y:PAYBACK_Y,w:BW-0.16,h:0.32,fontFace:'Montserrat',fontSize:8,bold:true,color:COR.verde,align:'center',valign:'middle',margin:0});
    // Linha verde do payback até a barra
    const bHpb=MH*(DIAS[bi]/90), bToppb=BY-bHpb;
    s3.addShape(pres.ShapeType.rect,{x:pbx+(BW/2)-0.007,y:PAYBACK_Y+0.32,w:0.014,h:Math.max(0,bToppb-(PAYBACK_Y+0.32)),fill:{color:'BFE3BF'}});
  }

  // Nota metodológica — abaixo das labels, fora do gráfico
  s3.addText(d.s3_nota||'',{x:0.35,y:4.22,w:9.30,h:0.28,fontFace:'Montserrat',fontSize:6,italic:true,color:'AAAAAA',valign:'top',margin:0,wrap:true});

  // Footer padronizado — altura compacta
  s3.addShape(pres.ShapeType.rect,{x:0,y:4.58,w:10,h:0.425,fill:{color:COR.dark}});
  s3.addShape(pres.ShapeType.rect,{x:0,y:4.58,w:0.05,h:0.425,fill:{color:COR.laranja}});
  s3.addText([
    {text:'Decisão adiada não é decisão neutra.  ',options:{bold:true,color:COR.laranja}},
    {text:`R$${Math.round(cm).toLocaleString('pt-BR')} por mês continuam saindo do caixa — com ou sem contrato assinado.`,options:{bold:false,color:COR.branco}}
  ],{x:0.22,y:4.60,w:7.80,h:0.36,fontFace:'Montserrat',fontSize:8.5,valign:'middle',margin:0});
  s3.addText('frota162.com.br',{x:8.20,y:4.64,w:1.65,h:0.28,fontFace:'Montserrat',fontSize:8,bold:true,color:COR.laranja,align:'right',valign:'middle',margin:0});

  return pres.writeFile({ fileName: outPath });
}

app.get('/', (req, res) => res.json({ status: 'ok', service: 'Frota162 PPTX v5' }));

app.post('/generate', (req, res) => {
  res.json({ ok: true, status: 'processing' });

  (async () => {
    try {
      let raw = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
      raw = raw.replace(/```json/gi,'').replace(/```/g,'').trim();
      const input = JSON.parse(raw);
      const { titulo, executivo, data_call, transcricao, pasta_mes_id } = input;

      // Filtro de qualidade — título deve conter "Frota162 ><" ou "Frota162 <>" (reunião com cliente)
      const tituloLower = (titulo||'').toLowerCase();
      const ehReuniaoCliente = tituloLower.includes('frota162 ><') || tituloLower.includes('frota162 <>') || tituloLower.includes('frota162><') || tituloLower.includes('frota162<>');
      if (!ehReuniaoCliente) {
        console.log('Descartado — não é reunião com cliente:', titulo);
        return;
      }
      if (!transcricao || transcricao.length < 500) {
        await postSlack(`:no_entry_sign: *Call descartada — ${titulo||'Sem título'}* (${executivo||''}): transcrição ausente ou muito curta para gerar material.`).catch(()=>{});
        return;
      }

      const conteudo = `Título: ${titulo||'Sem título'}\nData: ${data_call||''}\nExecutivo Frota162: ${executivo||''}\n\nTranscrição:\n${transcricao}`;
      const d = await callClaude(conteudo);

      const empresa = d.empresa || 'Prospect';
      const nomeArq = `Frota162 >< ${empresa} (Diretoria).pptx`;
      const outPath = path.join(os.tmpdir(), nomeArq);

      await gerarPPTX(d, outPath);

      const auth = new google.auth.GoogleAuth({
        credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
        scopes: ['https://www.googleapis.com/auth/drive'],
      });
      const drive = google.drive({ version: 'v3', auth });
      const pastaId = pasta_mes_id || process.env.PASTA_RAIZ_ID;

      const uploaded = await drive.files.create({
        supportsAllDrives: true,
        requestBody: { name: nomeArq, parents: [pastaId], mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
        media: { mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', body: fs.createReadStream(outPath) },
        fields: 'id,name,webViewLink',
      });

      await drive.permissions.create({ fileId: uploaded.data.id, supportsAllDrives: true, requestBody: { role: 'writer', type: 'anyone' } });
      fs.unlinkSync(outPath);

      // Temperatura com emoji
      const tempEmoji = d.temperatura==='quente' ? '🔴' : d.temperatura==='morno' ? '🟡' : '🔵';

      // Slack ID do executivo
      const slackId = getSlackId(d.executivo || executivo);
      const execMencao = slackId ? `<@${slackId}>` : (d.executivo || executivo || 'N/A');

      // ROI formatado
      const roiAnual = d.roi_anual || 0;
      const roiTexto = roiAnual > 0 ? `R$${Math.round(roiAnual).toLocaleString('pt-BR')}/ano` : 'A calcular';

      const msg = `:car: *Novo material e análise estratégica* :rocket:\n- *Empresa:* ${empresa}\n- *Executivo:* ${execMencao}\n- *Placas e MRR estimado:* ${d.placas||0} placas · ${d.z3_investimento||'A definir'}\n- *ROI estimado:* ${roiTexto}\n- *Material:* <${uploaded.data.webViewLink}|Abrir PPTX>\n- *Temperatura estimada:* ${tempEmoji} ${d.temperatura||'N/A'}\n- *Resumo Geral da negociação:* ${d.slack_resumo||''}`;

      await postSlack(msg);

    } catch(err) {
      console.error('Background error:', err.message);
      await postSlack(`:warning: Erro ao gerar PPTX: ${err.message}`).catch(()=>{});
    }
  })();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Frota162 PPTX Server v5 porta ${PORT}`));
