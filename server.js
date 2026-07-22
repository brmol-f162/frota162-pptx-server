const express = require('express');
const PptxGenJS = require('pptxgenjs');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const os = require('os');

const app = express();
app.use(express.json({ limit: '10mb' }));

// ─── CORES FROTA162 ───────────────────────────────────────────────
const COR = {
  laranja:  'E8401C',
  dark:     '1A1A1A',
  branco:   'FFFFFF',
  fundo:    'F7F6F4',
  card:     'F0EFED',
  divisor:  'E0DFDD',
  verde:    '1E6B1E',
  vermelho: 'CC2200',
};

// ─── HEALTH CHECK ─────────────────────────────────────────────────
app.get('/', (req, res) => res.json({ status: 'ok', service: 'Frota162 PPTX Generator' }));

// ─── ENDPOINT PRINCIPAL ───────────────────────────────────────────
app.post('/generate', async (req, res) => {
  try {
    const d = req.body;
    if (!d || !d.prospect_empresa) {
      return res.status(400).json({ error: 'JSON inválido — campo prospect_empresa ausente' });
    }

    const empresa = d.prospect_empresa || 'Prospect';
    const nomeArquivo = `Frota162 - ${empresa} (Diretoria).pptx`;
    const outputPath = path.join(os.tmpdir(), nomeArquivo);

    // ── INSTÂNCIA PPTX ─────────────────────────────────────────────
    const pres = new PptxGenJS();
    pres.layout = 'LAYOUT_WIDE'; // 13.33" x 7.5"
    pres.author = 'Frota162';
    pres.subject = `Material Executivo — ${empresa}`;

    // ── SLIDE 1: DOR ───────────────────────────────────────────────
    const s1 = pres.addSlide();
    s1.background = { color: COR.fundo };

    // Header laranja
    s1.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 1.35, fill: { color: COR.laranja } });
    s1.addText(d.slide1_titulo || `${empresa}: por que agir agora?`, {
      x: 0.35, y: 0.08, w: 12.6, h: 0.65,
      fontSize: 26, bold: true, color: COR.branco, fontFace: 'Montserrat',
      valign: 'middle', margin: 0,
    });
    s1.addText(d.slide1_subtitulo || d.placas + ' placas', {
      x: 0.35, y: 0.75, w: 12.6, h: 0.45,
      fontSize: 13, color: 'FFD0C0', fontFace: 'Montserrat',
      valign: 'middle', margin: 0,
    });

    // 4 cards 2x2
    const cards = [
      { stat: d.card1_stat, titulo: d.card1_titulo, desc: d.card1_desc },
      { stat: d.card2_stat, titulo: d.card2_titulo, desc: d.card2_desc },
      { stat: d.card3_stat, titulo: d.card3_titulo, desc: d.card3_desc },
      { stat: d.card4_stat, titulo: d.card4_titulo, desc: d.card4_desc },
    ];
    const cardW = 5.9, cardH = 2.45;
    const posicoes = [
      { x: 0.30, y: 1.55 }, { x: 6.93, y: 1.55 },
      { x: 0.30, y: 4.15 }, { x: 6.93, y: 4.15 },
    ];
    cards.forEach((c, i) => {
      const p = posicoes[i];
      // card bg
      s1.addShape(pres.ShapeType.rect, { x: p.x, y: p.y, w: cardW, h: cardH, fill: { color: COR.card }, line: { color: COR.divisor, width: 1 } });
      // stat
      s1.addText(c.stat || '—', { x: p.x + 0.20, y: p.y + 0.18, w: 5.5, h: 0.70, fontSize: 32, bold: true, color: COR.laranja, fontFace: 'Montserrat', margin: 0 });
      // titulo
      s1.addText(c.titulo || '', { x: p.x + 0.20, y: p.y + 0.90, w: 5.5, h: 0.40, fontSize: 13, bold: true, color: COR.dark, fontFace: 'Montserrat', margin: 0 });
      // desc
      s1.addText(c.desc || '', { x: p.x + 0.20, y: p.y + 1.32, w: 5.5, h: 0.90, fontSize: 11, color: '444444', fontFace: 'Montserrat', valign: 'top', margin: 0, wrap: true });
    });

    // Footer slide 1
    s1.addShape(pres.ShapeType.rect, { x: 0, y: 6.88, w: 13.33, h: 0.62, fill: { color: COR.dark } });
    s1.addShape(pres.ShapeType.rect, { x: 0, y: 6.88, w: 0.18, h: 0.62, fill: { color: COR.laranja } });
    s1.addText(d.slide1_footer || 'Cada dia sem visibilidade é um risco que cresce silenciosamente.', {
      x: 0.30, y: 6.90, w: 11.0, h: 0.55,
      fontSize: 11, color: COR.branco, fontFace: 'Montserrat', valign: 'middle', margin: 0,
    });
    s1.addText('frota162.com.br', {
      x: 11.50, y: 6.90, w: 1.65, h: 0.55,
      fontSize: 10, color: COR.laranja, fontFace: 'Montserrat', align: 'right', valign: 'middle', margin: 0,
    });

    // ── SLIDE 2: SOLUÇÃO + FINANCEIRO ─────────────────────────────
    const s2 = pres.addSlide();
    s2.background = { color: COR.fundo };

    // Header dark
    s2.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 1.10, fill: { color: COR.dark } });
    s2.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: 1.10, fill: { color: COR.laranja } });
    s2.addText(`O que a Frota162 entrega para ${empresa}`, {
      x: 0.35, y: 0.10, w: 12.6, h: 0.60,
      fontSize: 22, bold: true, color: COR.branco, fontFace: 'Montserrat', valign: 'middle', margin: 0,
    });
    s2.addText('Solução + Retorno Financeiro', {
      x: 0.35, y: 0.70, w: 12.6, h: 0.32,
      fontSize: 12, color: 'AAAAAA', fontFace: 'Montserrat', valign: 'middle', margin: 0,
    });

    // Divisor vertical
    s2.addShape(pres.ShapeType.rect, { x: 5.30, y: 1.25, w: 0.03, h: 5.40, fill: { color: COR.divisor } });

    // ── COLUNA ESQUERDA: FINANCEIRO ────────────────────────────────
    // Box vermelho — situação atual
    s2.addShape(pres.ShapeType.rect, { x: 0.30, y: 1.30, w: 4.70, h: 1.65, fill: { color: 'FFF0EE' }, line: { color: COR.vermelho, width: 1.5 } });
    s2.addText('SITUAÇÃO ATUAL', { x: 0.45, y: 1.38, w: 4.40, h: 0.30, fontSize: 10, bold: true, color: COR.vermelho, fontFace: 'Montserrat', margin: 0 });
    s2.addText(d.slide2_custo_atual || 'Custos não mapeados com multas e irregularidades', {
      x: 0.45, y: 1.68, w: 4.40, h: 1.15,
      fontSize: 11, color: COR.dark, fontFace: 'Montserrat', valign: 'top', margin: 0, wrap: true,
    });

    // Seta
    s2.addText('▼', { x: 0.30, y: 3.05, w: 4.70, h: 0.40, fontSize: 18, color: COR.laranja, fontFace: 'Montserrat', align: 'center', margin: 0 });

    // Box verde — com Frota162
    s2.addShape(pres.ShapeType.rect, { x: 0.30, y: 3.50, w: 4.70, h: 1.80, fill: { color: 'F0FFF0' }, line: { color: COR.verde, width: 1.5 } });
    s2.addText('COM FROTA162', { x: 0.45, y: 3.58, w: 4.40, h: 0.30, fontSize: 10, bold: true, color: COR.verde, fontFace: 'Montserrat', margin: 0 });
    s2.addText(d.slide2_investimento || 'Investimento mensal a confirmar', {
      x: 0.45, y: 3.88, w: 4.40, h: 0.42,
      fontSize: 18, bold: true, color: COR.verde, fontFace: 'Montserrat', margin: 0,
    });
    s2.addText(d.slide2_economia || 'Economia estimada a calcular', {
      x: 0.45, y: 4.32, w: 4.40, h: 0.80,
      fontSize: 11, color: COR.dark, fontFace: 'Montserrat', valign: 'top', margin: 0, wrap: true,
    });

    // Nota rodapé financeiro
    s2.addText('Pós-pago · Sem fidelidade · Aviso prévio 30 dias', {
      x: 0.30, y: 5.40, w: 4.70, h: 0.35,
      fontSize: 9, color: '888888', fontFace: 'Montserrat', align: 'center', margin: 0,
    });

    // ── COLUNA DIREITA: FEATURES ───────────────────────────────────
    const features = [
      { titulo: d.feat1_titulo, desc: d.feat1_desc, cor: COR.laranja },
      { titulo: d.feat2_titulo, desc: d.feat2_desc, cor: '2E86AB' },
      { titulo: d.feat3_titulo, desc: d.feat3_desc, cor: COR.verde },
      { titulo: d.feat4_titulo, desc: d.feat4_desc, cor: '7B2D8B' },
    ];
    features.forEach((f, i) => {
      const fy = 1.30 + i * 1.20;
      s2.addShape(pres.ShapeType.rect, { x: 5.55, y: fy, w: 0.45, h: 0.45, fill: { color: f.cor } });
      s2.addText(f.titulo || '', { x: 6.10, y: fy, w: 6.90, h: 0.38, fontSize: 13, bold: true, color: COR.dark, fontFace: 'Montserrat', valign: 'middle', margin: 0 });
      s2.addText(f.desc || '', { x: 6.10, y: fy + 0.40, w: 6.90, h: 0.65, fontSize: 11, color: '555555', fontFace: 'Montserrat', valign: 'top', margin: 0, wrap: true });
    });

    // CTA pill
    s2.addShape(pres.ShapeType.rect, { x: 5.55, y: 5.95, w: 7.48, h: 0.60, fill: { color: COR.laranja } });
    s2.addText(d.cta || 'Solicite uma demonstração para sua equipe', {
      x: 5.55, y: 5.95, w: 7.48, h: 0.60,
      fontSize: 14, bold: true, color: COR.branco, fontFace: 'Montserrat', align: 'center', valign: 'middle', margin: 0,
    });

    // Footer slide 2
    s2.addShape(pres.ShapeType.rect, { x: 0, y: 6.88, w: 13.33, h: 0.62, fill: { color: COR.dark } });
    s2.addShape(pres.ShapeType.rect, { x: 0, y: 6.88, w: 0.18, h: 0.62, fill: { color: COR.laranja } });
    s2.addText('frota162.com.br', {
      x: 11.50, y: 6.90, w: 1.65, h: 0.55,
      fontSize: 10, color: COR.laranja, fontFace: 'Montserrat', align: 'right', valign: 'middle', margin: 0,
    });

    // ── SALVA PPTX ────────────────────────────────────────────────
    await pres.writeFile({ fileName: outputPath });

    // ── UPLOAD DRIVE ──────────────────────────────────────────────
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
      scopes: ['https://www.googleapis.com/auth/drive'],
    });
    const drive = google.drive({ version: 'v3', auth });

    const pastaId = req.body.pasta_mes_id || process.env.PASTA_RAIZ_ID;

    const uploaded = await drive.files.create({
      requestBody: {
        name: nomeArquivo,
        parents: [pastaId],
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      },
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        body: fs.createReadStream(outputPath),
      },
      fields: 'id, name, webViewLink',
    });

    // Permissão pública de edição
    await drive.permissions.create({
      fileId: uploaded.data.id,
      requestBody: { role: 'writer', type: 'anyone' },
    });

    // Remove arquivo temporário
    fs.unlinkSync(outputPath);

    res.json({
      ok: true,
      fileId: uploaded.data.id,
      fileName: uploaded.data.name,
      link: uploaded.data.webViewLink,
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Frota162 PPTX Server rodando na porta ${PORT}`));
