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
  laranja: 'E8401C', dark: '1A1A1A', branco: 'FFFFFF',
  fundo: 'F7F6F4', divisor: 'E0DFDD', verde: '1E6B1E',
  vermelho: 'CC2200', azul: '1565C0', cinza: '888888',
};

const SYSTEM = `Você é especialista em vendas B2B da Frota162. Analise a transcrição e retorne SOMENTE JSON válido sem markdown sem backticks.

TABELA DE PREÇOS: Até 40 placas mínimo E1 R$397 E2 R$549 E3 R$649. 41-99 E1 R$9,48 E2 R$14,23 E3 R$16,60. 100-199 E1 R$9,00 E2 R$13,51 E3 R$15,77. 200-299 E1 R$7,65 E2 R$11,49 E3 R$13,41. 300-399 E1 R$7,27 E2 R$10,91 E3 R$12,74. 400-499 E1 R$6,91 E2 R$10,37 E3 R$12,10. 500-999 E1 R$6,56 E2 R$9,85 E3 R$11,49. 1000-1999 E1 R$5,90 E2 R$8,86 E3 R$10,34. CNPJ adicional R$150/mês. E1=multas+SNE+1CNPJ. E2=E1+IPVA+indicação+3CNPJs. E3=E2+CNH+tox+5CNPJs.

ROI: economia_multas=multas_mes x valor x 0.4(SNE) ou 0.2. economia_NIC=NIC_tratadas x valor x 0.6. economia_pessoas=(func-1) x 2500. ROI_anual=economia_total_anual - investimento_anual. custo_mensal_base=multas_mes x 130 se sem ROI. payback=30 x (investimento_mes / custo_mensal).

REGRAS: valor mínimo multa R$130. sinal menos SÓ nas barras slide 3. tem_roi=false se call não confirmou. títulos máx 45 chars.

JSON:{"empresa":"","executivo":"","perfil_lead":"","placas":0,"cnpjs":0,"segmento":"","tem_roi":true,"s1_header_bold":"","s1_header_sub":"","s1_subtitulo":"","cards":[{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""}],"s1_footer_bold":"","s1_footer_normal":"","s2_header_bold":"","s2_header_normal":"","z1_stat1":"","z1_sub1":"","z1_stat2":"","z1_sub2":"","z1_bullets":["",""],"passos":[{"titulo":"","desc":""},{"titulo":"","desc":""},{"titulo":"","desc":""},{"titulo":"","desc":""}],"z3_stat":"","z3_sub1":"","z3_investimento":"","z3_badge":"","z3_nota":"","s2_cta_bold":"","s2_cta_normal":"","custo_mensal":0,"investimento_mensal_num":0,"s3_header_bold":"","s3_header_sub":"","s3_formula":"","s3_nota":"","slack_resumo":"","proximo_passo":""}`;

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
          // Tenta corrigir JSON truncado fechando chaves/colchetes abertos
          try {
            resolve(JSON.parse(t));
          } catch(parseErr) {
            // Conta chaves e colchetes para tentar fechar
            let opens = 0, openBrackets = 0;
            for (const ch of t) {
              if (ch === '{') opens++;
              else if (ch === '}') opens--;
              else if (ch === '[') openBrackets++;
              else if (ch === ']') openBrackets--;
            }
            // Remove trailing comma/incomplete field
            t = t.replace(/,\s*$/, '').replace(/,\s*"[^"]*"\s*:\s*[^,}\]]*$/, '');
            // Close open brackets and braces
            t += ']'.repeat(Math.max(0, openBrackets)) + '}'.repeat(Math.max(0, opens));
            try {
              resolve(JSON.parse(t));
            } catch(e2) {
              reject(new Error('Claude parse: ' + parseErr.message));
            }
          }
        } catch(e) { reject(new Error('Claude parse: ' + e.message)); }
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
      hostname: url.hostname, path: url.pathname + url.search,
      method: 'POST',
      headers: { 'content-type': 'application/json', 'content-length': Buffer.byteLength(body) }
    }, (res) => { res.on('data', () => {}); res.on('end', resolve); });
    req.on('error', () => resolve());
    req.write(body); req.end();
  });
}

function gerarPPTX(d, outPath) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';

  // SLIDE 1
  const s1 = pres.addSlide();
  s1.background = { color: 'F7F6F4' };
  s1.addShape(pres.ShapeType.rect, { x:0, y:0, w:10, h:0.92, fill:{ color: COR.laranja } });
  s1.addText([
    { text: (d.s1_header_bold||'') + '\n', options: { bold:true, fontSize:14.5, color:COR.branco } },
    { text: d.s1_header_sub||'', options: { bold:false, fontSize:11, color:'FFD0C0' } }
  ], { x:0.38, y:0, w:9.3, h:0.92, fontFace:'Montserrat', valign:'middle', margin:0 });
  s1.addText(d.s1_subtitulo||'', { x:0.38, y:0.96, w:9.3, h:0.22, fontFace:'Montserrat', fontSize:8.5, color:COR.cinza, italic:true });

  const CPOS = [{cx:0.22,cy:1.27},{cx:5.14,cy:1.27},{cx:0.22,cy:3.13},{cx:5.14,cy:3.13}];
  const CPAL = [
    {fundo:'FFF5F3',strip:COR.vermelho,stat:COR.vermelho},
    {fundo:'F0EFED',strip:COR.cinza,stat:COR.dark},
    {fundo:'F0F5FF',strip:COR.azul,stat:COR.azul},
    {fundo:'F0EFED',strip:COR.cinza,stat:COR.dark},
  ];
  (d.cards||[]).forEach((c,i) => {
    if(i>3) return;
    const {cx,cy}=CPOS[i]; const p=CPAL[i];
    s1.addShape(pres.ShapeType.rect,{x:cx,y:cy,w:4.60,h:1.72,fill:{color:p.fundo},line:{color:COR.divisor,width:0.5}});
    s1.addShape(pres.ShapeType.rect,{x:cx,y:cy+0.16,w:0.045,h:1.40,fill:{color:p.strip}});
    s1.addText(c.stat||'',{x:cx+0.14,y:cy+0.10,w:4.40,h:0.44,fontFace:'Montserrat',fontSize:22,bold:true,color:p.stat,margin:0});
    s1.addText(c.titulo||'',{x:cx+0.14,y:cy+0.52,w:4.38,h:0.26,fontFace:'Montserrat',fontSize:9.5,bold:true,color:COR.dark,margin:0});
    s1.addText(c.desc||'',{x:cx+0.14,y:cy+0.78,w:4.38,h:0.86,fontFace:'Montserrat',fontSize:7.8,color:'333333',valign:'top',margin:0,wrap:true});
  });
  s1.addShape(pres.ShapeType.rect,{x:0,y:4.90,w:10,h:0.725,fill:{color:COR.dark}});
  s1.addShape(pres.ShapeType.rect,{x:0,y:4.90,w:0.045,h:0.725,fill:{color:COR.laranja}});
  s1.addText([
    {text:(d.s1_footer_bold||'')+' ',options:{bold:true,color:COR.laranja}},
    {text:d.s1_footer_normal||'',options:{bold:false,color:COR.branco}}
  ],{x:0.26,y:4.92,w:7.80,h:0.52,fontFace:'Montserrat',fontSize:10,valign:'middle',margin:0});
  s1.addText('frota162.com.br',{x:8.30,y:5.02,w:1.55,h:0.40,fontFace:'Montserrat',fontSize:9,bold:true,color:COR.laranja,align:'right',valign:'middle',margin:0});

  // SLIDE 2
  const s2 = pres.addSlide();
  s2.background = { color: 'F7F6F4' };
  s2.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:0.84,fill:{color:COR.dark}});
  s2.addShape(pres.ShapeType.rect,{x:0,y:0,w:0.045,h:0.84,fill:{color:COR.laranja}});
  s2.addText([
    {text:(d.s2_header_bold||'')+'  ',options:{bold:true,color:COR.laranja}},
    {text:d.s2_header_normal||'',options:{bold:false,color:COR.branco}}
  ],{x:0.38,y:0,w:9.3,h:0.84,fontFace:'Montserrat',fontSize:13,valign:'middle',margin:0});

  const CY=0.84,CH=4.16;
  const Z1X=0,Z1W=3.10,Z2X=3.10,Z2W=3.80,Z3X=6.90,Z3W=3.10;

  s2.addShape(pres.ShapeType.rect,{x:Z1X,y:CY,w:Z1W,h:CH,fill:{color:'FFEDE7'}});
  s2.addText('HOJE',{x:Z1X+0.18,y:CY+0.18,w:Z1W-0.22,h:0.22,fontFace:'Montserrat',fontSize:9.5,bold:true,color:COR.vermelho,charSpacing:1,margin:0});
  s2.addText(d.z1_stat1||'',{x:Z1X+0.18,y:CY+0.44,w:Z1W-0.22,h:0.50,fontFace:'Montserrat',fontSize:28,bold:true,color:COR.vermelho,margin:0});
  s2.addText(d.z1_sub1||'',{x:Z1X+0.18,y:CY+0.96,w:Z1W-0.22,h:0.28,fontFace:'Montserrat',fontSize:8.5,color:'555555',margin:0,wrap:true});
  s2.addShape(pres.ShapeType.rect,{x:Z1X+0.18,y:CY+1.32,w:Z1W-0.36,h:0.018,fill:{color:'F0B8A5'}});
  s2.addText(d.z1_stat2||'',{x:Z1X+0.18,y:CY+1.40,w:Z1W-0.22,h:0.35,fontFace:'Montserrat',fontSize:19,bold:true,color:COR.vermelho,margin:0});
  s2.addText(d.z1_sub2||'',{x:Z1X+0.18,y:CY+1.78,w:Z1W-0.22,h:0.28,fontFace:'Montserrat',fontSize:8.5,color:'555555',margin:0,wrap:true});
  (d.z1_bullets||[]).forEach((b,i)=>{
    s2.addText('· '+b,{x:Z1X+0.18,y:CY+2.18+(i*0.30),w:Z1W-0.22,h:0.28,fontFace:'Montserrat',fontSize:8,color:'444444',margin:0,wrap:true});
  });

  [Z1W-0.08,Z2X+Z2W-0.08].forEach(sx=>{
    s2.addShape(pres.ShapeType.ellipse,{x:sx-0.10,y:CY+(CH/2)-0.16,w:0.32,h:0.32,fill:{color:COR.branco},line:{color:COR.laranja,width:1.5}});
    s2.addText('→',{x:sx-0.10,y:CY+(CH/2)-0.16,w:0.32,h:0.32,fontFace:'Montserrat',fontSize:10,bold:true,color:COR.laranja,align:'center',valign:'middle',margin:0});
  });

  s2.addShape(pres.ShapeType.rect,{x:Z2X,y:CY,w:Z2W,h:CH,fill:{color:'FCFCFB'}});
  s2.addText('COMO A FROTA162 RESOLVE',{x:Z2X+0.18,y:CY+0.18,w:Z2W-0.22,h:0.22,fontFace:'Montserrat',fontSize:9.5,bold:true,color:COR.dark,charSpacing:0.5,margin:0});
  const SY=CY+0.55,SH=0.70,SG=0.14,rx=Z2X+0.38;
  s2.addShape(pres.ShapeType.rect,{x:rx+0.17,y:SY+0.21,w:0.018,h:(SH+SG)*3+0.21,fill:{color:COR.divisor}});
  (d.passos||[]).forEach((p,i)=>{
    const py=SY+i*(SH+SG);
    s2.addShape(pres.ShapeType.ellipse,{x:rx,y:py+0.08,w:0.42,h:0.42,fill:{color:COR.laranja}});
    s2.addText(String(i+1),{x:rx,y:py+0.08,w:0.42,h:0.42,fontFace:'Montserrat',fontSize:13,bold:true,color:COR.branco,align:'center',valign:'middle',margin:0});
    s2.addText(p.titulo||'',{x:Z2X+0.90,y:py+0.08,w:Z2W-1.0,h:0.26,fontFace:'Montserrat',fontSize:9.5,bold:true,color:COR.dark,margin:0,wrap:true});
    s2.addText(p.desc||'',{x:Z2X+0.90,y:py+0.36,w:Z2W-1.0,h:0.36,fontFace:'Montserrat',fontSize:8.5,color:'555555',valign:'top',margin:0,wrap:true});
  });

  s2.addShape(pres.ShapeType.rect,{x:Z3X,y:CY,w:Z3W,h:CH,fill:{color:'EAF6EA'}});
  s2.addText(d.tem_roi?'RESULTADO':'PLANO RECOMENDADO',{x:Z3X+0.18,y:CY+0.18,w:Z3W-0.22,h:0.22,fontFace:'Montserrat',fontSize:9.5,bold:true,color:COR.verde,charSpacing:1,margin:0});
  s2.addText(d.z3_stat||'',{x:Z3X+0.18,y:CY+0.44,w:Z3W-0.22,h:0.50,fontFace:'Montserrat',fontSize:28,bold:true,color:COR.verde,margin:0});
  s2.addText(d.z3_sub1||'',{x:Z3X+0.18,y:CY+0.96,w:Z3W-0.22,h:0.28,fontFace:'Montserrat',fontSize:8.5,color:'555555',margin:0,wrap:true});
  s2.addShape(pres.ShapeType.rect,{x:Z3X+0.18,y:CY+1.32,w:Z3W-0.36,h:0.018,fill:{color:'BFE3BF'}});
  s2.addText(d.z3_investimento||'',{x:Z3X+0.18,y:CY+1.40,w:Z3W-0.22,h:0.35,fontFace:'Montserrat',fontSize:15,bold:true,color:COR.verde,margin:0,wrap:true});
  s2.addShape(pres.ShapeType.rect,{x:Z3X+0.18,y:CY+1.90,w:Z3W-0.36,h:0.38,fill:{color:'D4EED4'},line:{color:'BFE3BF',width:0.5}});
  s2.addText(d.z3_badge||'',{x:Z3X+0.22,y:CY+1.90,w:Z3W-0.44,h:0.38,fontFace:'Montserrat',fontSize:8,bold:true,color:COR.verde,align:'center',valign:'middle',margin:0,wrap:true});
  s2.addText(d.z3_nota||'',{x:Z3X+0.18,y:CY+3.30,w:Z3W-0.22,h:0.60,fontFace:'Montserrat',fontSize:6.5,italic:true,color:'AAAAAA',valign:'top',margin:0,wrap:true});

  s2.addShape(pres.ShapeType.rect,{x:0,y:5.00,w:10,h:0.625,fill:{color:COR.dark}});
  s2.addShape(pres.ShapeType.rect,{x:0,y:5.00,w:0.045,h:0.625,fill:{color:COR.laranja}});
  s2.addText([
    {text:(d.s2_cta_bold||'')+'  ',options:{bold:true,color:COR.laranja}},
    {text:d.s2_cta_normal||'',options:{bold:false,color:COR.branco}}
  ],{x:0.26,y:5.00,w:7.80,h:0.52,fontFace:'Montserrat',fontSize:9,valign:'middle',margin:0});
  s2.addText('frota162.com.br',{x:8.30,y:5.10,w:1.55,h:0.35,fontFace:'Montserrat',fontSize:9,bold:true,color:COR.laranja,align:'right',valign:'middle',margin:0});

  // SLIDE 3
  const s3 = pres.addSlide();
  s3.background = { color: 'F7F6F4' };
  s3.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:0.92,fill:{color:COR.laranja}});
  s3.addText([
    {text:(d.s3_header_bold||'')+'\n',options:{bold:true,fontSize:14.5,color:COR.branco}},
    {text:d.s3_header_sub||'',options:{bold:false,fontSize:10,color:'FFD0C0'}}
  ],{x:0.38,y:0,w:9.3,h:0.92,fontFace:'Montserrat',valign:'middle',margin:0});
  s3.addText(d.s3_formula||'',{x:0.38,y:0.96,w:9.3,h:0.22,fontFace:'Montserrat',fontSize:8.5,color:COR.cinza,italic:true});

  const cm=d.custo_mensal||0;
  const DIAS=[15,30,60,90],VALS=[cm/2,cm,cm*2,cm*3];
  const BY=4.05,MH=2.15,BW=1.98,BG=0.29,X0=0.60;
  DIAS.forEach((dia,i)=>{
    const bH=MH*(dia/90),bTop=BY-bH,bx=X0+i*(BW+BG);
    s3.addShape(pres.ShapeType.rect,{x:bx,y:bTop,w:BW,h:bH,fill:{color:'FFCDD2'},line:{color:COR.vermelho,width:0.5}});
    s3.addText(`-R$${Math.round(VALS[i]).toLocaleString('pt-BR')}`,{x:bx,y:bTop-0.36,w:BW,h:0.32,fontFace:'Montserrat',fontSize:14,bold:true,color:COR.vermelho,align:'center',margin:0});
    s3.addText(`${dia} DIAS`,{x:bx,y:BY+0.10,w:BW,h:0.26,fontFace:'Montserrat',fontSize:9.5,bold:true,color:'555555',align:'center',margin:0});
  });
  s3.addShape(pres.ShapeType.rect,{x:0.40,y:BY,w:9.20,h:0.018,fill:{color:'CCCCCC'}});

  const inv=d.investimento_mensal_num||0;
  if(inv>0&&cm>0){
    const dp=Math.round(30*(inv/cm));
    if(dp<=90){
      let bi=DIAS.findIndex(x=>x>=dp); if(bi<0)bi=3;
      const pbx=X0+bi*(BW+BG),pbH=MH*(DIAS[bi]/90),pbTop=BY-pbH;
      s3.addShape(pres.ShapeType.rect,{x:pbx+0.15,y:pbTop-0.58,w:BW-0.30,h:0.35,fill:{color:COR.branco},line:{color:COR.verde,width:1.5}});
      s3.addText(dp<=45?`↑ Payback: ~${dp} dias`:`↑ Payback: ~${Math.round(dp/30)} meses`,{x:pbx+0.15,y:pbTop-0.58,w:BW-0.30,h:0.35,fontFace:'Montserrat',fontSize:8.5,bold:true,color:COR.verde,align:'center',valign:'middle',margin:0});
    }
  }

  s3.addText(d.s3_nota||'',{x:0.38,y:4.42,w:9.24,h:0.38,fontFace:'Montserrat',fontSize:6.5,italic:true,color:'AAAAAA',valign:'top',margin:0,wrap:true});
  s3.addShape(pres.ShapeType.rect,{x:0,y:4.90,w:10,h:0.725,fill:{color:COR.dark}});
  s3.addShape(pres.ShapeType.rect,{x:0,y:4.90,w:0.045,h:0.725,fill:{color:COR.laranja}});
  s3.addText([
    {text:'Decisão adiada não é decisão neutra.  ',options:{bold:true,color:COR.laranja}},
    {text:`R$${Math.round(cm).toLocaleString('pt-BR')} por mês continuam saindo do caixa, com ou sem contrato assinado.`,options:{bold:false,color:COR.branco}}
  ],{x:0.26,y:4.92,w:7.80,h:0.52,fontFace:'Montserrat',fontSize:10,valign:'middle',margin:0});
  s3.addText('frota162.com.br',{x:8.30,y:5.02,w:1.55,h:0.40,fontFace:'Montserrat',fontSize:9,bold:true,color:COR.laranja,align:'right',valign:'middle',margin:0});

  return pres.writeFile({ fileName: outPath });
}

app.get('/', (req, res) => res.json({ status: 'ok', service: 'Frota162 PPTX v4' }));

app.post('/generate', (req, res) => {
  // Responde imediatamente — processa em background
  res.json({ ok: true, status: 'processing' });

  // Processa de forma assíncrona
  (async () => {
    try {
      let raw = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
      raw = raw.replace(/```json/gi,'').replace(/```/g,'').trim();
      const input = JSON.parse(raw);
      const { titulo, executivo, data_call, transcricao, pasta_mes_id } = input;

      // Filtro de qualidade — descarta calls sem transcrição suficiente
    if (!transcricao || transcricao.length < 500) {
      await postSlack(`:no_entry_sign: *Call descartada — ${titulo||'Sem título'}* (${executivo||''}): transcrição ausente ou muito curta para gerar material.`).catch(()=>{});
      return;
    }

    const conteudo = `Título: ${titulo||'Sem título'}\nData: ${data_call||''}\nExecutivo Frota162: ${executivo||''}\n\nTranscrição:\n${transcricao||''}`;
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
        requestBody: {
          name: nomeArq, parents: [pastaId],
          mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        },
        media: {
          mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          body: fs.createReadStream(outPath),
        },
        fields: 'id,name,webViewLink',
      });

      await drive.permissions.create({
        fileId: uploaded.data.id, supportsAllDrives: true,
        requestBody: { role: 'writer', type: 'anyone' },
      });

      fs.unlinkSync(outPath);

      const msg = `:car: *Frota162 — ${d.empresa} · ${d.executivo}*\n${d.placas} placas · ${d.perfil_lead} · ${data_call||''}\n*MRR estimado:* ${d.z3_investimento}\n\n${d.slack_resumo}\n\n*Próximo passo:* ${d.proximo_passo}\n\n<${uploaded.data.webViewLink}|Abrir PPTX Executivo>`;

      await postSlack(msg);

    } catch(err) {
      console.error('Background error:', err.message);
      await postSlack(`:warning: Erro ao gerar PPTX: ${err.message}`).catch(()=>{});
    }
  })();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Frota162 PPTX Server v4 porta ${PORT}`));
