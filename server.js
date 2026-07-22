const express = require('express');
const PptxGenJS = require('pptxgenjs');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const os = require('os');

const app = express();
app.use(express.text({ type: '*/*', limit: '10mb' }));
app.use(express.json({ limit: '10mb' }));

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

app.get('/', (req, res) => res.json({ status: 'ok', service: 'Frota162 PPTX 3 Slides' }));

app.post('/generate', async (req, res) => {
  try {
    let raw = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
    raw = raw.replace(/```json/gi, '').replace(/```/g, '').trim();
    const d = JSON.parse(raw);

    if (!d || !d.empresa) return res.status(400).json({ error: 'Campo empresa ausente' });

    const empresa = d.empresa;
    const nomeArq = `Frota162 >< ${empresa} (Diretoria).pptx`;
    const outPath = path.join(os.tmpdir(), nomeArq);

    const pres = new PptxGenJS();
    pres.layout = 'LAYOUT_16x9'; // 10 x 5.625

    // ─────────────────────────────────────────────
    // SLIDE 1 — A dor em linguagem executiva
    // ─────────────────────────────────────────────
    const s1 = pres.addSlide();
    s1.background = { color: 'F7F6F4' };

    // Header laranja
    s1.addShape(pres.ShapeType.rect, { x:0, y:0, w:10, h:0.92, fill:{ color: COR.laranja } });
    s1.addText([
      { text: d.s1_header_bold + '\n', options: { bold: true, fontSize: 14.5, color: COR.branco } },
      { text: d.s1_header_sub, options: { bold: false, fontSize: 11, color: 'FFD0C0' } }
    ], { x:0.38, y:0, w:9.3, h:0.92, fontFace:'Montserrat', valign:'middle', margin:0 });

    // Subtítulo
    s1.addText(d.s1_subtitulo || '', {
      x:0.38, y:0.96, w:9.3, h:0.22,
      fontFace:'Montserrat', fontSize:8.5, color:COR.cinza, italic:true
    });

    // 4 Cards 2x2
    const CARDS = [
      { cx:0.22, cy:1.27 },
      { cx:5.14, cy:1.27 },
      { cx:0.22, cy:3.13 },
      { cx:5.14, cy:3.13 },
    ];
    const CW = 4.60, CH = 1.72;
    const PALETAS = [
      { fundo:'FFF5F3', strip:COR.vermelho, stat:COR.vermelho },
      { fundo:'F0EFED', strip:COR.cinza,    stat:COR.dark },
      { fundo:'F0F5FF', strip:COR.azul,     stat:COR.azul },
      { fundo:'F0EFED', strip:COR.cinza,    stat:COR.dark },
    ];
    const cards = d.cards || [];

    cards.forEach((c, i) => {
      if (i > 3) return;
      const { cx, cy } = CARDS[i];
      const p = PALETAS[i];
      // fundo
      s1.addShape(pres.ShapeType.roundRect, { x:cx, y:cy, w:CW, h:CH, fill:{ color: p.fundo }, line:{ color: COR.divisor, width:0.5 }, rectRadius:0.06 });
      // strip
      s1.addShape(pres.ShapeType.rect, { x:cx, y:cy+0.16, w:0.045, h:CH-0.32, fill:{ color: p.strip } });
      // stat
      s1.addText(c.stat || '', { x:cx+0.14, y:cy+0.10, w:CW-0.20, h:0.44, fontFace:'Montserrat', fontSize:22, bold:true, color:p.stat, margin:0 });
      // titulo
      s1.addText(c.titulo || '', { x:cx+0.14, y:cy+0.52, w:CW-0.22, h:0.26, fontFace:'Montserrat', fontSize:9.5, bold:true, color:COR.dark, margin:0 });
      // desc
      s1.addText(c.desc || '', { x:cx+0.14, y:cy+0.78, w:CW-0.22, h:0.86, fontFace:'Montserrat', fontSize:7.8, color:'333333', valign:'top', margin:0, wrap:true });
    });

    // Footer dark slide 1
    s1.addShape(pres.ShapeType.rect, { x:0, y:4.90, w:10, h:0.725, fill:{ color: COR.dark } });
    s1.addShape(pres.ShapeType.rect, { x:0, y:4.90, w:0.045, h:0.725, fill:{ color: COR.laranja } });
    s1.addText([
      { text: d.s1_footer_bold + ' ', options: { bold: true, color: COR.laranja } },
      { text: d.s1_footer_normal || '', options: { bold: false, color: COR.branco } }
    ], { x:0.26, y:4.92, w:7.80, h:0.52, fontFace:'Montserrat', fontSize:10, valign:'middle', margin:0 });
    s1.addText('frota162.com.br', { x:8.30, y:5.02, w:1.55, h:0.40, fontFace:'Montserrat', fontSize:9, bold:true, color:COR.laranja, align:'right', valign:'middle', margin:0 });

    // ─────────────────────────────────────────────
    // SLIDE 2 — Da dor à solução (3 zonas)
    // ─────────────────────────────────────────────
    const s2 = pres.addSlide();
    s2.background = { color: 'F7F6F4' };

    // Header dark
    s2.addShape(pres.ShapeType.rect, { x:0, y:0, w:10, h:0.84, fill:{ color: COR.dark } });
    s2.addShape(pres.ShapeType.rect, { x:0, y:0, w:0.045, h:0.84, fill:{ color: COR.laranja } });
    s2.addText([
      { text: d.s2_header_bold + '  ', options: { bold: true, color: COR.laranja } },
      { text: d.s2_header_normal || '', options: { bold: false, color: COR.branco } }
    ], { x:0.38, y:0, w:9.3, h:0.84, fontFace:'Montserrat', fontSize:13, valign:'middle', margin:0 });

    const CORPO_Y = 0.84;
    const CORPO_H = 4.16; // até y=5.00
    const Z1X=0, Z1W=3.10;
    const Z2X=3.10, Z2W=3.80;
    const Z3X=6.90, Z3W=3.10;

    // ZONA 1 — HOJE (vermelho claro)
    s2.addShape(pres.ShapeType.rect, { x:Z1X, y:CORPO_Y, w:Z1W, h:CORPO_H, fill:{ color:'FFEDE7' } });
    s2.addText('HOJE', { x:Z1X+0.18, y:CORPO_Y+0.18, w:Z1W-0.22, h:0.22, fontFace:'Montserrat', fontSize:9.5, bold:true, color:COR.vermelho, charSpacing:1, margin:0 });
    s2.addText(d.z1_stat1 || '', { x:Z1X+0.18, y:CORPO_Y+0.44, w:Z1W-0.22, h:0.50, fontFace:'Montserrat', fontSize:28, bold:true, color:COR.vermelho, margin:0 });
    s2.addText(d.z1_sub1 || '', { x:Z1X+0.18, y:CORPO_Y+0.96, w:Z1W-0.22, h:0.28, fontFace:'Montserrat', fontSize:8.5, color:'555555', margin:0, wrap:true });
    // divisor fino
    s2.addShape(pres.ShapeType.rect, { x:Z1X+0.18, y:CORPO_Y+1.32, w:Z1W-0.36, h:0.018, fill:{ color:'F0B8A5' } });
    s2.addText(d.z1_stat2 || '', { x:Z1X+0.18, y:CORPO_Y+1.40, w:Z1W-0.22, h:0.35, fontFace:'Montserrat', fontSize:19, bold:true, color:COR.vermelho, margin:0 });
    s2.addText(d.z1_sub2 || '', { x:Z1X+0.18, y:CORPO_Y+1.78, w:Z1W-0.22, h:0.28, fontFace:'Montserrat', fontSize:8.5, color:'555555', margin:0, wrap:true });
    // bullets
    const z1bullets = d.z1_bullets || [];
    z1bullets.forEach((b, i) => {
      s2.addText('· ' + b, { x:Z1X+0.18, y:CORPO_Y+2.18+(i*0.30), w:Z1W-0.22, h:0.28, fontFace:'Montserrat', fontSize:8, color:'444444', margin:0, wrap:true });
    });

    // Setas de conexão zona 1→2 e 2→3
    [Z1W-0.08, Z2X+Z2W-0.08].forEach(sx => {
      s2.addShape(pres.ShapeType.oval, { x:sx-0.10, y:CORPO_Y+(CORPO_H/2)-0.16, w:0.32, h:0.32, fill:{ color: COR.branco }, line:{ color: COR.laranja, width:1.5 } });
      s2.addText('→', { x:sx-0.10, y:CORPO_Y+(CORPO_H/2)-0.16, w:0.32, h:0.32, fontFace:'Montserrat', fontSize:10, bold:true, color:COR.laranja, align:'center', valign:'middle', margin:0 });
    });

    // ZONA 2 — COMO A FROTA162 RESOLVE (quase branco)
    s2.addShape(pres.ShapeType.rect, { x:Z2X, y:CORPO_Y, w:Z2W, h:CORPO_H, fill:{ color:'FCFCFB' } });
    s2.addText('COMO A FROTA162 RESOLVE', { x:Z2X+0.18, y:CORPO_Y+0.18, w:Z2W-0.22, h:0.22, fontFace:'Montserrat', fontSize:9.5, bold:true, color:COR.dark, charSpacing:0.5, margin:0 });

    // Rail vertical
    const STEP_Y0 = CORPO_Y+0.55;
    const STEP_H = 0.70;
    const STEP_GAP = 0.14;
    const rail_x = Z2X+0.38;
    s2.addShape(pres.ShapeType.rect, { x:rail_x+0.17, y:STEP_Y0+0.21, w:0.018, h:(STEP_H+STEP_GAP)*3+0.21, fill:{ color: COR.divisor } });

    const passos = d.passos || [];
    passos.forEach((p, i) => {
      const py = STEP_Y0 + i*(STEP_H+STEP_GAP);
      s2.addShape(pres.ShapeType.oval, { x:rail_x, y:py+0.08, w:0.42, h:0.42, fill:{ color: COR.laranja } });
      s2.addText(String(i+1), { x:rail_x, y:py+0.08, w:0.42, h:0.42, fontFace:'Montserrat', fontSize:13, bold:true, color:COR.branco, align:'center', valign:'middle', margin:0 });
      s2.addText(p.titulo || '', { x:Z2X+0.90, y:py+0.08, w:Z2W-1.0, h:0.26, fontFace:'Montserrat', fontSize:9.5, bold:true, color:COR.dark, margin:0, wrap:true });
      s2.addText(p.desc || '', { x:Z2X+0.90, y:py+0.36, w:Z2W-1.0, h:0.36, fontFace:'Montserrat', fontSize:8.5, color:'555555', valign:'top', margin:0, wrap:true });
    });

    // ZONA 3 — RESULTADO ou PLANO RECOMENDADO (verde claro)
    s2.addShape(pres.ShapeType.rect, { x:Z3X, y:CORPO_Y, w:Z3W, h:CORPO_H, fill:{ color:'EAF6EA' } });
    const z3label = d.tem_roi ? 'RESULTADO' : 'PLANO RECOMENDADO';
    s2.addText(z3label, { x:Z3X+0.18, y:CORPO_Y+0.18, w:Z3W-0.22, h:0.22, fontFace:'Montserrat', fontSize:9.5, bold:true, color:COR.verde, charSpacing:1, margin:0 });
    s2.addText(d.z3_stat || '', { x:Z3X+0.18, y:CORPO_Y+0.44, w:Z3W-0.22, h:0.50, fontFace:'Montserrat', fontSize:28, bold:true, color:COR.verde, margin:0 });
    s2.addText(d.z3_sub1 || '', { x:Z3X+0.18, y:CORPO_Y+0.96, w:Z3W-0.22, h:0.28, fontFace:'Montserrat', fontSize:8.5, color:'555555', margin:0, wrap:true });
    s2.addShape(pres.ShapeType.rect, { x:Z3X+0.18, y:CORPO_Y+1.32, w:Z3W-0.36, h:0.018, fill:{ color:'BFE3BF' } });
    s2.addText(d.z3_investimento || '', { x:Z3X+0.18, y:CORPO_Y+1.40, w:Z3W-0.22, h:0.35, fontFace:'Montserrat', fontSize:15, bold:true, color:COR.verde, margin:0, wrap:true });
    // badge prova social ou badge diferencial
    s2.addShape(pres.ShapeType.roundRect, { x:Z3X+0.18, y:CORPO_Y+1.90, w:Z3W-0.36, h:0.38, fill:{ color:'D4EED4' }, line:{ color:'BFE3BF', width:0.5 }, rectRadius:0.06 });
    s2.addText(d.z3_badge || '', { x:Z3X+0.22, y:CORPO_Y+1.90, w:Z3W-0.44, h:0.38, fontFace:'Montserrat', fontSize:8, bold:true, color:COR.verde, align:'center', valign:'middle', margin:0, wrap:true });
    // nota rodapé zona 3
    s2.addText(d.z3_nota || '', { x:Z3X+0.18, y:CORPO_Y+3.30, w:Z3W-0.22, h:0.60, fontFace:'Montserrat', fontSize:6.5, italic:true, color:'AAAAAA', valign:'top', margin:0, wrap:true });

    // Banner final full-width
    s2.addShape(pres.ShapeType.rect, { x:0, y:5.00, w:10, h:0.625, fill:{ color: COR.dark } });
    s2.addShape(pres.ShapeType.rect, { x:0, y:5.00, w:0.045, h:0.625, fill:{ color: COR.laranja } });
    s2.addText([
      { text: d.s2_cta_bold + '  ', options: { bold: true, color: COR.laranja } },
      { text: d.s2_cta_normal || '', options: { bold: false, color: COR.branco } }
    ], { x:0.26, y:5.00, w:7.80, h:0.52, fontFace:'Montserrat', fontSize:9, valign:'middle', margin:0 });
    s2.addText('frota162.com.br', { x:8.30, y:5.10, w:1.55, h:0.35, fontFace:'Montserrat', fontSize:9, bold:true, color:COR.laranja, align:'right', valign:'middle', margin:0 });

    // ─────────────────────────────────────────────
    // SLIDE 3 — O custo de esperar
    // ─────────────────────────────────────────────
    const s3 = pres.addSlide();
    s3.background = { color: 'F7F6F4' };

    // Header laranja
    s3.addShape(pres.ShapeType.rect, { x:0, y:0, w:10, h:0.92, fill:{ color: COR.laranja } });
    s3.addText([
      { text: d.s3_header_bold + '\n', options: { bold: true, fontSize: 14.5, color: COR.branco } },
      { text: d.s3_header_sub || '', options: { bold: false, fontSize: 10, color: 'FFD0C0' } }
    ], { x:0.38, y:0, w:9.3, h:0.92, fontFace:'Montserrat', valign:'middle', margin:0 });

    // Subtítulo com fórmula
    s3.addText(d.s3_formula || '', {
      x:0.38, y:0.96, w:9.3, h:0.22,
      fontFace:'Montserrat', fontSize:8.5, color:COR.cinza, italic:true
    });

    // 4 barras crescentes
    const custo_mensal = d.custo_mensal || 0;
    const DIAS = [15, 30, 60, 90];
    const VALORES = [custo_mensal/2, custo_mensal, custo_mensal*2, custo_mensal*3];
    const BASELINE_Y = 4.05;
    const MAX_H = 2.15;
    const bw = 1.98;
    const gap = 0.29;
    const x0 = 0.60;

    DIAS.forEach((dia, i) => {
      const val = VALORES[i];
      const ratio = dia / 90;
      const barH = MAX_H * ratio;
      const barTop = BASELINE_Y - barH;
      const bx = x0 + i * (bw + gap);

      s3.addShape(pres.ShapeType.rect, { x:bx, y:barTop, w:bw, h:barH, fill:{ color:'FFCDD2' }, line:{ color:COR.vermelho, width:0.5 } });
      // valor acima
      s3.addText(`-R$${Math.round(val).toLocaleString('pt-BR')}`, {
        x:bx, y:barTop-0.36, w:bw, h:0.32,
        fontFace:'Montserrat', fontSize:14, bold:true, color:COR.vermelho, align:'center', margin:0
      });
      // label abaixo
      s3.addText(`${dia} DIAS`, {
        x:bx, y:BASELINE_Y+0.10, w:bw, h:0.26,
        fontFace:'Montserrat', fontSize:9.5, bold:true, color:'555555', align:'center', margin:0
      });
    });

    // Linha baseline
    s3.addShape(pres.ShapeType.rect, { x:0.40, y:BASELINE_Y, w:9.20, h:0.018, fill:{ color:'CCCCCC' } });

    // Marcador de payback
    const investimento = d.investimento_mensal_num || 0;
    if (investimento > 0 && custo_mensal > 0) {
      const dias_payback = Math.round(30 * (investimento / custo_mensal));
      if (dias_payback <= 90) {
        // Ancora na primeira barra cujo marco >= dias_payback
        let barra_idx = DIAS.findIndex(d => d >= dias_payback);
        if (barra_idx < 0) barra_idx = 3;
        const pbx = x0 + barra_idx * (bw + gap);
        const pb_ratio = DIAS[barra_idx] / 90;
        const pb_barH = MAX_H * pb_ratio;
        const pb_barTop = BASELINE_Y - pb_barH;

        s3.addShape(pres.ShapeType.roundRect, { x:pbx+0.15, y:pb_barTop-0.58, w:bw-0.30, h:0.35, fill:{ color: COR.branco }, line:{ color: COR.verde, width:1.5 }, rectRadius:0.06 });
        const payback_label = dias_payback <= 45 ? `↑ Payback: ~${dias_payback} dias` : `↑ Payback: ~${Math.round(dias_payback/30)} meses`;
        s3.addText(payback_label, {
          x:pbx+0.15, y:pb_barTop-0.58, w:bw-0.30, h:0.35,
          fontFace:'Montserrat', fontSize:8.5, bold:true, color:COR.verde, align:'center', valign:'middle', margin:0
        });
      }
    }

    // Nota rodapé slide 3
    s3.addText(d.s3_nota || '', {
      x:0.38, y:4.42, w:9.24, h:0.38,
      fontFace:'Montserrat', fontSize:6.5, italic:true, color:'AAAAAA', valign:'top', margin:0, wrap:true
    });

    // Footer dark slide 3
    s3.addShape(pres.ShapeType.rect, { x:0, y:4.90, w:10, h:0.725, fill:{ color: COR.dark } });
    s3.addShape(pres.ShapeType.rect, { x:0, y:4.90, w:0.045, h:0.725, fill:{ color: COR.laranja } });
    s3.addText([
      { text: 'Decisão adiada não é decisão neutra.  ', options: { bold: true, color: COR.laranja } },
      { text: `R$${Math.round(custo_mensal).toLocaleString('pt-BR')} por mês continuam saindo do caixa, com ou sem contrato assinado.`, options: { bold: false, color: COR.branco } }
    ], { x:0.26, y:4.92, w:7.80, h:0.52, fontFace:'Montserrat', fontSize:10, valign:'middle', margin:0 });
    s3.addText('frota162.com.br', { x:8.30, y:5.02, w:1.55, h:0.40, fontFace:'Montserrat', fontSize:9, bold:true, color:COR.laranja, align:'right', valign:'middle', margin:0 });

    // ─────────────────────────────────────────────
    // SALVA E UPLOAD DRIVE
    // ─────────────────────────────────────────────
    await pres.writeFile({ fileName: outPath });

    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
      scopes: ['https://www.googleapis.com/auth/drive'],
    });
    const drive = google.drive({ version: 'v3', auth });
    const pastaId = d.pasta_mes_id || process.env.PASTA_RAIZ_ID;

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

    res.json({ ok: true, fileId: uploaded.data.id, fileName: uploaded.data.name, link: uploaded.data.webViewLink });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Frota162 PPTX Server porta ${PORT}`));
