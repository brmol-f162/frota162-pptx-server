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
  laranja:  'E8401C',
  dark:     '1A1A1A',
  branco:   'FFFFFF',
  fundo:    'F7F6F4',
  divisor:  'E0DFDD',
  verde:    '1E6B1E',
  vermelho: 'CC2200',
  azul:     '1565C0',
  cinza:    '888888',
};

const SYSTEM_PROMPT = `Você é especialista em vendas B2B da Frota162, plataforma SaaS de gestão de frotas (multas, CNH, IPVA, licenciamento, NTT/ANTT). Analise a transcrição da call comercial e retorne SOMENTE JSON válido, sem markdown, sem backticks, sem texto fora do JSON.

TABELA DE PREÇOS FROTA162:
- Até 40 placas (mínimo mensal): E1 R$397,00 | E2 R$549,00 | E3 R$649,00
- 41 a 99 placas (por placa): E1 R$9,48 | E2 R$14,23 | E3 R$16,60
- 100 a 199 placas (por placa): E1 R$9,00 | E2 R$13,51 | E3 R$15,77
- 200 a 299 placas (por placa): E1 R$7,65 | E2 R$11,49 | E3 R$13,41
- 300 a 399 placas (por placa): E1 R$7,27 | E2 R$10,91 | E3 R$12,74
- 400 a 499 placas (por placa): E1 R$6,91 | E2 R$10,37 | E3 R$12,10
- 500 a 999 placas (por placa): E1 R$6,56 | E2 R$9,85 | E3 R$11,49
- 1.000 a 1.999 placas (por placa): E1 R$5,90 | E2 R$8,86 | E3 R$10,34
- 2.000 a 2.999 placas (por placa): E1 R$5,31 | E2 R$7,98 | E3 R$9,31
- 3.000 a 3.999 placas (por placa): E1 R$4,78 | E2 R$7,18 | E3 R$8,38
- 4.000 a 4.999 placas (por placa): E1 R$4,30 | E2 R$6,46 | E3 R$7,54
- Acima de 5.000 placas (por placa): E1 R$3,87 | E2 R$5,82 | E3 R$6,79
- CNPJ adicional acima do incluso: R$150/CNPJ/mês
- Planos: E1=notificações+multas+SNE+1CNPJ | E2=E1+IPVA+indicação+3CNPJs | E3=E2+CNH+toxicológico+5CNPJs

LÓGICA DE CÁLCULO ROI:
1. ECONOMIA COM MULTAS: multas_mes x valor_multa x desconto_SNE (40% com SNE ativo, 20% sem SNE)
2. ECONOMIA COM NIC: NIC_tratadas_mes x valor_multa x (1 - desconto). Sem Frota162 o custo por NIC é 3x o valor da multa.
3. ECONOMIA COM PESSOAS: (funcionarios_atuais - 1) x R$2.500/mês
4. ECONOMIA TOTAL ANUAL = (economia_multas_mes + economia_nic_mes + economia_pessoas_mes) x 12
5. INVESTIMENTO MENSAL = calcular pela tabela acima para a faixa de placas e plano recomendado
6. ROI ANUAL = ECONOMIA TOTAL ANUAL - (INVESTIMENTO MENSAL x 12)
7. CUSTO MENSAL BASE PARA SLIDE 3: usar o número mais específico da call. Se não houver ROI calculado, usar: multas_mes x R$130 como base
8. PAYBACK = 30 x (investimento_mensal / custo_mensal_base). Se > 90 dias, não forçar marcador.
9. PLANO RECOMENDADO: E1 para frotas simples. E2 para múltiplos CNPJs ou indicação frequente. E3 para CNH, toxicológico ou 4+ CNPJs.

REGRAS CRÍTICAS:
- Valor mínimo de multa quando não especificado: R$130 (NUNCA dizer valor médio)
- Sinal de menos APENAS nos valores das barras do slide 3 (-R$X)
- Se a call não confirmou ROI: tem_roi=false, não fabricar número
- Verdades desconfortáveis vão escritas nos cards
- Títulos de cards e passos: máximo 45 caracteres

Retorne SOMENTE este JSON sem nenhum texto fora:
{"empresa":"nome da empresa prospect","executivo":"nome do AE da Frota162","perfil_lead":"decisor ou influenciador","placas":0,"cnpjs":0,"segmento":"descricao","tem_roi":true,"s1_header_bold":"frase gancho especifica max 65 chars","s1_header_sub":"X placas Y CNPJs Regiao","s1_subtitulo":"contexto segmento","cards":[{"stat":"","titulo":"max 45 chars","desc":""},{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""},{"stat":"","titulo":"","desc":""}],"s1_footer_bold":"urgencia max 80 chars","s1_footer_normal":"complemento","s2_header_bold":"Da dor de -R$X ao retorno de +R$Y por ano. OU X placas processo manual a recomendacao e o Plano N.","s2_header_normal":"Como a Frota162 resolve em 3 passos [Empresa]","z1_stat1":"","z1_sub1":"","z1_stat2":"","z1_sub2":"","z1_bullets":["",""],"passos":[{"titulo":"max 40 chars","desc":""},{"titulo":"","desc":""},{"titulo":"","desc":""},{"titulo":"","desc":""}],"z3_stat":"ROI anual ou Plano N","z3_sub1":"retorno estimado por ano ou features do plano","z3_investimento":"R$X,XX/mes","z3_badge":"diferencial forte","z3_nota":"condicoes comerciais e limitacoes","s2_cta_bold":"proximo passo combinado","s2_cta_normal":"acao concreta","custo_mensal":0,"investimento_mensal_num":0,"s3_header_bold":"Cada mes sem a Frota162 e R$X saindo do caixa da [Empresa].","s3_header_sub":"com base no prejuizo que a propria call confirmou","s3_formula":"formula usada","s3_nota":"metodologia e limitacoes","slack_resumo":"2 linhas max","proximo_passo":"acao concreta para o executivo"}`;

function callClaude(transcricao) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 2000,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: transcricao }]
    });

    const req = https.request({
      hostname: 'api.anthropic.com',
      path: '/v1/messages',
      method: 'POST',
      headers: {
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
        'content-length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          const text = parsed.content[0].text;
          const clean = text.replace(/```json/gi, '').replace(/```/g, '').trim();
          resolve(JSON.parse(clean));
        } catch(e) {
          reject(new Error('Claude parse error: ' + e.message + ' | raw: ' + data.substring(0, 200)));
        }
      });
    });

    req.on('error', reject);
    req.setTimeout(120000, () => { req.destroy(); reject(new Error('Claude timeout 120s')); });
    req.write(body);
    req.end();
  });
}

app.get('/', (req, res) => res.json({ status: 'ok', service: 'Frota162 PPTX 3 Slides v3' }));

app.post('/generate', async (req, res) => {
  try {
    // Recebe dados do Make: titulo, executivo, data, transcricao, pasta_mes_id
    let raw = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
    raw = raw.replace(/```json/gi, '').replace(/```/g, '').trim();
    const input = JSON.parse(raw);

    const { titulo, executivo, data_call, transcricao, pasta_mes_id } = input;

    // Chama Claude com a transcrição completa
    const conteudo = `Título da call: ${titulo || 'Sem título'}\nData: ${data_call || 'Sem data'}\nExecutivo Frota162: ${executivo || 'Sem nome'}\n\nTranscrição completa:\n${transcricao || 'Sem transcrição'}`;

    const d = await callClaude(conteudo);

    const empresa = d.empresa || 'Prospect';
    const nomeArq = `Frota162 >< ${empresa} (Diretoria).pptx`;
    const outPath = path.join(os.tmpdir(), nomeArq);

    const pres = new PptxGenJS();
    pres.layout = 'LAYOUT_16x9';

    // ── SLIDE 1 ──────────────────────────────────
    const s1 = pres.addSlide();
    s1.background = { color: 'F7F6F4' };
    s1.addShape(pres.ShapeType.rect, { x:0, y:0, w:10, h:0.92, fill:{ color: COR.laranja } });
    s1.addText([
      { text: (d.s1_header_bold || '') + '\n', options: { bold: true, fontSize: 14.5, color: COR.branco } },
      { text: d.s1_header_sub || '', options: { bold: false, fontSize: 11, color: 'FFD0C0' } }
    ], { x:0.38, y:0, w:9.3, h:0.92, fontFace:'Montserrat', valign:'middle', margin:0 });
    s1.addText(d.s1_subtitulo || '', { x:0.38, y:0.96, w:9.3, h:0.22, fontFace:'Montserrat', fontSize:8.5, color:COR.cinza, italic:true });

    const CARDS = [{ cx:0.22, cy:1.27 },{ cx:5.14, cy:1.27 },{ cx:0.22, cy:3.13 },{ cx:5.14, cy:3.13 }];
    const CW = 4.60, CH = 1.72;
    const PALETAS = [
      { fundo:'FFF5F3', strip:COR.vermelho, stat:COR.vermelho },
      { fundo:'F0EFED', strip:COR.cinza, stat:COR.dark },
      { fundo:'F0F5FF', strip:COR.azul, stat:COR.azul },
      { fundo:'F0EFED', strip:COR.cinza, stat:COR.dark },
    ];

    (d.cards || []).forEach((c, i) => {
      if (i > 3) return;
      const { cx, cy } = CARDS[i];
      const p = PALETAS[i];
      s1.addShape(pres.ShapeType.roundRect, { x:cx, y:cy, w:CW, h:CH, fill:{ color: p.fundo }, line:{ color: COR.divisor, width:0.5 }, rectRadius:0.06 });
      s1.addShape(pres.ShapeType.rect, { x:cx, y:cy+0.16, w:0.045, h:CH-0.32, fill:{ color: p.strip } });
      s1.addText(c.stat || '', { x:cx+0.14, y:cy+0.10, w:CW-0.20, h:0.44, fontFace:'Montserrat', fontSize:22, bold:true, color:p.stat, margin:0 });
      s1.addText(c.titulo || '', { x:cx+0.14, y:cy+0.52, w:CW-0.22, h:0.26, fontFace:'Montserrat', fontSize:9.5, bold:true, color:COR.dark, margin:0 });
      s1.addText(c.desc || '', { x:cx+0.14, y:cy+0.78, w:CW-0.22, h:0.86, fontFace:'Montserrat', fontSize:7.8, color:'333333', valign:'top', margin:0, wrap:true });
    });

    s1.addShape(pres.ShapeType.rect, { x:0, y:4.90, w:10, h:0.725, fill:{ color: COR.dark } });
    s1.addShape(pres.ShapeType.rect, { x:0, y:4.90, w:0.045, h:0.725, fill:{ color: COR.laranja } });
    s1.addText([
      { text: (d.s1_footer_bold || '') + ' ', options: { bold: true, color: COR.laranja } },
      { text: d.s1_footer_normal || '', options: { bold: false, color: COR.branco } }
    ], { x:0.26, y:4.92, w:7.80, h:0.52, fontFace:'Montserrat', fontSize:10, valign:'middle', margin:0 });
    s1.addText('frota162.com.br', { x:8.30, y:5.02, w:1.55, h:0.40, fontFace:'Montserrat', fontSize:9, bold:true, color:COR.laranja, align:'right', valign:'middle', margin:0 });

    // ── SLIDE 2 ──────────────────────────────────
    const s2 = pres.addSlide();
    s2.background = { color: 'F7F6F4' };
    s2.addShape(pres.ShapeType.rect, { x:0, y:0, w:10, h:0.84, fill:{ color: COR.dark } });
    s2.addShape(pres.ShapeType.rect, { x:0, y:0, w:0.045, h:0.84, fill:{ color: COR.laranja } });
    s2.addText([
      { text: (d.s2_header_bold || '') + '  ', options: { bold: true, color: COR.laranja } },
      { text: d.s2_header_normal || '', options: { bold: false, color: COR.branco } }
    ], { x:0.38, y:0, w:9.3, h:0.84, fontFace:'Montserrat', fontSize:13, valign:'middle', margin:0 });

    const CY = 0.84, CH2 = 4.16;
    const Z1X=0, Z1W=3.10, Z2X=3.10, Z2W=3.80, Z3X=6.90, Z3W=3.10;

    // Zona 1 HOJE
    s2.addShape(pres.ShapeType.rect, { x:Z1X, y:CY, w:Z1W, h:CH2, fill:{ color:'FFEDE7' } });
    s2.addText('HOJE', { x:Z1X+0.18, y:CY+0.18, w:Z1W-0.22, h:0.22, fontFace:'Montserrat', fontSize:9.5, bold:true, color:COR.vermelho, charSpacing:1, margin:0 });
    s2.addText(d.z1_stat1 || '', { x:Z1X+0.18, y:CY+0.44, w:Z1W-0.22, h:0.50, fontFace:'Montserrat', fontSize:28, bold:true, color:COR.vermelho, margin:0 });
    s2.addText(d.z1_sub1 || '', { x:Z1X+0.18, y:CY+0.96, w:Z1W-0.22, h:0.28, fontFace:'Montserrat', fontSize:8.5, color:'555555', margin:0, wrap:true });
    s2.addShape(pres.ShapeType.rect, { x:Z1X+0.18, y:CY+1.32, w:Z1W-0.36, h:0.018, fill:{ color:'F0B8A5' } });
    s2.addText(d.z1_stat2 || '', { x:Z1X+0.18, y:CY+1.40, w:Z1W-0.22, h:0.35, fontFace:'Montserrat', fontSize:19, bold:true, color:COR.vermelho, margin:0 });
    s2.addText(d.z1_sub2 || '', { x:Z1X+0.18, y:CY+1.78, w:Z1W-0.22, h:0.28, fontFace:'Montserrat', fontSize:8.5, color:'555555', margin:0, wrap:true });
    (d.z1_bullets || []).forEach((b, i) => {
      s2.addText('· ' + b, { x:Z1X+0.18, y:CY+2.18+(i*0.30), w:Z1W-0.22, h:0.28, fontFace:'Montserrat', fontSize:8, color:'444444', margin:0, wrap:true });
    });

    // Setas
    [Z1W-0.08, Z2X+Z2W-0.08].forEach(sx => {
      s2.addShape(pres.ShapeType.oval, { x:sx-0.10, y:CY+(CH2/2)-0.16, w:0.32, h:0.32, fill:{ color: COR.branco }, line:{ color: COR.laranja, width:1.5 } });
      s2.addText('→', { x:sx-0.10, y:CY+(CH2/2)-0.16, w:0.32, h:0.32, fontFace:'Montserrat', fontSize:10, bold:true, color:COR.laranja, align:'center', valign:'middle', margin:0 });
    });

    // Zona 2 COMO RESOLVE
    s2.addShape(pres.ShapeType.rect, { x:Z2X, y:CY, w:Z2W, h:CH2, fill:{ color:'FCFCFB' } });
    s2.addText('COMO A FROTA162 RESOLVE', { x:Z2X+0.18, y:CY+0.18, w:Z2W-0.22, h:0.22, fontFace:'Montserrat', fontSize:9.5, bold:true, color:COR.dark, charSpacing:0.5, margin:0 });

    const STEP_Y0 = CY+0.55, STEP_H = 0.70, STEP_GAP = 0.14;
    const rail_x = Z2X+0.38;
    s2.addShape(pres.ShapeType.rect, { x:rail_x+0.17, y:STEP_Y0+0.21, w:0.018, h:(STEP_H+STEP_GAP)*3+0.21, fill:{ color: COR.divisor } });

    (d.passos || []).forEach((p, i) => {
      const py = STEP_Y0 + i*(STEP_H+STEP_GAP);
      s2.addShape(pres.ShapeType.oval, { x:rail_x, y:py+0.08, w:0.42, h:0.42, fill:{ color: COR.laranja } });
      s2.addText(String(i+1), { x:rail_x, y:py+0.08, w:0.42, h:0.42, fontFace:'Montserrat', fontSize:13, bold:true, color:COR.branco, align:'center', valign:'middle', margin:0 });
      s2.addText(p.titulo || '', { x:Z2X+0.90, y:py+0.08, w:Z2W-1.0, h:0.26, fontFace:'Montserrat', fontSize:9.5, bold:true, color:COR.dark, margin:0, wrap:true });
      s2.addText(p.desc || '', { x:Z2X+0.90, y:py+0.36, w:Z2W-1.0, h:0.36, fontFace:'Montserrat', fontSize:8.5, color:'555555', valign:'top', margin:0, wrap:true });
    });

    // Zona 3 RESULTADO
    s2.addShape(pres.ShapeType.rect, { x:Z3X, y:CY, w:Z3W, h:CH2, fill:{ color:'EAF6EA' } });
    s2.addText(d.tem_roi ? 'RESULTADO' : 'PLANO RECOMENDADO', { x:Z3X+0.18, y:CY+0.18, w:Z3W-0.22, h:0.22, fontFace:'Montserrat', fontSize:9.5, bold:true, color:COR.verde, charSpacing:1, margin:0 });
    s2.addText(d.z3_stat || '', { x:Z3X+0.18, y:CY+0.44, w:Z3W-0.22, h:0.50, fontFace:'Montserrat', fontSize:28, bold:true, color:COR.verde, margin:0 });
    s2.addText(d.z3_sub1 || '', { x:Z3X+0.18, y:CY+0.96, w:Z3W-0.22, h:0.28, fontFace:'Montserrat', fontSize:8.5, color:'555555', margin:0, wrap:true });
    s2.addShape(pres.ShapeType.rect, { x:Z3X+0.18, y:CY+1.32, w:Z3W-0.36, h:0.018, fill:{ color:'BFE3BF' } });
    s2.addText(d.z3_investimento || '', { x:Z3X+0.18, y:CY+1.40, w:Z3W-0.22, h:0.35, fontFace:'Montserrat', fontSize:15, bold:true, color:COR.verde, margin:0, wrap:true });
    s2.addShape(pres.ShapeType.roundRect, { x:Z3X+0.18, y:CY+1.90, w:Z3W-0.36, h:0.38, fill:{ color:'D4EED4' }, line:{ color:'BFE3BF', width:0.5 }, rectRadius:0.06 });
    s2.addText(d.z3_badge || '', { x:Z3X+0.22, y:CY+1.90, w:Z3W-0.44, h:0.38, fontFace:'Montserrat', fontSize:8, bold:true, color:COR.verde, align:'center', valign:'middle', margin:0, wrap:true });
    s2.addText(d.z3_nota || '', { x:Z3X+0.18, y:CY+3.30, w:Z3W-0.22, h:0.60, fontFace:'Montserrat', fontSize:6.5, italic:true, color:'AAAAAA', valign:'top', margin:0, wrap:true });

    s2.addShape(pres.ShapeType.rect, { x:0, y:5.00, w:10, h:0.625, fill:{ color: COR.dark } });
    s2.addShape(pres.ShapeType.rect, { x:0, y:5.00, w:0.045, h:0.625, fill:{ color: COR.laranja } });
    s2.addText([
      { text: (d.s2_cta_bold || '') + '  ', options: { bold: true, color: COR.laranja } },
      { text: d.s2_cta_normal || '', options: { bold: false, color: COR.branco } }
    ], { x:0.26, y:5.00, w:7.80, h:0.52, fontFace:'Montserrat', fontSize:9, valign:'middle', margin:0 });
    s2.addText('frota162.com.br', { x:8.30, y:5.10, w:1.55, h:0.35, fontFace:'Montserrat', fontSize:9, bold:true, color:COR.laranja, align:'right', valign:'middle', margin:0 });

    // ── SLIDE 3 ──────────────────────────────────
    const s3 = pres.addSlide();
    s3.background = { color: 'F7F6F4' };
    s3.addShape(pres.ShapeType.rect, { x:0, y:0, w:10, h:0.92, fill:{ color: COR.laranja } });
    s3.addText([
      { text: (d.s3_header_bold || '') + '\n', options: { bold: true, fontSize: 14.5, color: COR.branco } },
      { text: d.s3_header_sub || '', options: { bold: false, fontSize: 10, color: 'FFD0C0' } }
    ], { x:0.38, y:0, w:9.3, h:0.92, fontFace:'Montserrat', valign:'middle', margin:0 });
    s3.addText(d.s3_formula || '', { x:0.38, y:0.96, w:9.3, h:0.22, fontFace:'Montserrat', fontSize:8.5, color:COR.cinza, italic:true });

    const custo_mensal = d.custo_mensal || 0;
    const DIAS = [15, 30, 60, 90];
    const VALORES = [custo_mensal/2, custo_mensal, custo_mensal*2, custo_mensal*3];
    const BASELINE_Y = 4.05, MAX_H = 2.15, bw = 1.98, gap = 0.29, x0 = 0.60;

    DIAS.forEach((dia, i) => {
      const ratio = dia / 90;
      const barH = MAX_H * ratio;
      const barTop = BASELINE_Y - barH;
      const bx = x0 + i * (bw + gap);
      s3.addShape(pres.ShapeType.rect, { x:bx, y:barTop, w:bw, h:barH, fill:{ color:'FFCDD2' }, line:{ color:COR.vermelho, width:0.5 } });
      s3.addText(`-R$${Math.round(VALORES[i]).toLocaleString('pt-BR')}`, { x:bx, y:barTop-0.36, w:bw, h:0.32, fontFace:'Montserrat', fontSize:14, bold:true, color:COR.vermelho, align:'center', margin:0 });
      s3.addText(`${dia} DIAS`, { x:bx, y:BASELINE_Y+0.10, w:bw, h:0.26, fontFace:'Montserrat', fontSize:9.5, bold:true, color:'555555', align:'center', margin:0 });
    });

    s3.addShape(pres.ShapeType.rect, { x:0.40, y:BASELINE_Y, w:9.20, h:0.018, fill:{ color:'CCCCCC' } });

    const investimento = d.investimento_mensal_num || 0;
    if (investimento > 0 && custo_mensal > 0) {
      const dias_payback = Math.round(30 * (investimento / custo_mensal));
      if (dias_payback <= 90) {
        let barra_idx = DIAS.findIndex(d => d >= dias_payback);
        if (barra_idx < 0) barra_idx = 3;
        const pbx = x0 + barra_idx * (bw + gap);
        const pb_barH = MAX_H * (DIAS[barra_idx] / 90);
        const pb_barTop = BASELINE_Y - pb_barH;
        s3.addShape(pres.ShapeType.roundRect, { x:pbx+0.15, y:pb_barTop-0.58, w:bw-0.30, h:0.35, fill:{ color: COR.branco }, line:{ color: COR.verde, width:1.5 }, rectRadius:0.06 });
        const label = dias_payback <= 45 ? `↑ Payback: ~${dias_payback} dias` : `↑ Payback: ~${Math.round(dias_payback/30)} meses`;
        s3.addText(label, { x:pbx+0.15, y:pb_barTop-0.58, w:bw-0.30, h:0.35, fontFace:'Montserrat', fontSize:8.5, bold:true, color:COR.verde, align:'center', valign:'middle', margin:0 });
      }
    }

    s3.addText(d.s3_nota || '', { x:0.38, y:4.42, w:9.24, h:0.38, fontFace:'Montserrat', fontSize:6.5, italic:true, color:'AAAAAA', valign:'top', margin:0, wrap:true });
    s3.addShape(pres.ShapeType.rect, { x:0, y:4.90, w:10, h:0.725, fill:{ color: COR.dark } });
    s3.addShape(pres.ShapeType.rect, { x:0, y:4.90, w:0.045, h:0.725, fill:{ color: COR.laranja } });
    s3.addText([
      { text: 'Decisão adiada não é decisão neutra.  ', options: { bold: true, color: COR.laranja } },
      { text: `R$${Math.round(custo_mensal).toLocaleString('pt-BR')} por mês continuam saindo do caixa, com ou sem contrato assinado.`, options: { bold: false, color: COR.branco } }
    ], { x:0.26, y:4.92, w:7.80, h:0.52, fontFace:'Montserrat', fontSize:10, valign:'middle', margin:0 });
    s3.addText('frota162.com.br', { x:8.30, y:5.02, w:1.55, h:0.40, fontFace:'Montserrat', fontSize:9, bold:true, color:COR.laranja, align:'right', valign:'middle', margin:0 });

    // ── SALVA E UPLOAD DRIVE ──────────────────────
    await pres.writeFile({ fileName: outPath });

    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
      scopes: ['https://www.googleapis.com/auth/drive'],
    });
    const drive = google.drive({ version: 'v3', auth });
    const pastaId = pasta_mes_id || process.env.PASTA_RAIZ_ID;

    const uploaded = await drive.files.create({
      supportsAllDrives: true,
      requestBody: {
        name: nomeArq,
        parents: [pastaId],
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      },
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        body: fs.createReadStream(outPath),
      },
      fields: 'id,name,webViewLink',
    });

    await drive.permissions.create({
      fileId: uploaded.data.id,
      supportsAllDrives: true,
      requestBody: { role: 'writer', type: 'anyone' },
    });

    fs.unlinkSync(outPath);

    res.json({
      ok: true,
      fileId: uploaded.data.id,
      fileName: uploaded.data.name,
      link: uploaded.data.webViewLink,
      empresa: d.empresa,
      executivo: d.executivo,
      perfil_lead: d.perfil_lead,
      placas: d.placas,
      z3_investimento: d.z3_investimento,
      slack_resumo: d.slack_resumo,
      proximo_passo: d.proximo_passo
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Frota162 PPTX Server v3 porta ${PORT}`));
