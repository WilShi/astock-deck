/**
 * A-share research deck generator.
 *
 * Produces a native, fully editable .pptx (no screenshots, no HTML): text is text,
 * tables are tables, and the charts are native PowerPoint charts that the user can
 * re-point at their own data.
 */

import PptxGenJS from "pptxgenjs";
import type { FinancialPeriod, Kline, Quote } from "./datasource/types.ts";

// pptxgenjs ships UMD-style typings whose namespace members are not importable as
// ESM types, so derive them from the class API instead of naming them directly.
type Slide = ReturnType<PptxGenJS["addSlide"]>;
type TableRow = Parameters<Slide["addTable"]>[0][number];
import { computeTtm, formatMoney, formatPct, impliedGrowth, pickValuationModel, runDcf, runPbRoe, type DcfAssumptions } from "./analysis/valuation.ts";

const C = {
  ink: "12161F",
  muted: "6B7280",
  line: "E5E7EB",
  accent: "2E6BE6",
  accentSoft: "EAF0FD",
  up: "C0392B",
  down: "1E8E5A",
  cover: "0B1220",
  coverSoft: "93A0B8",
  white: "FFFFFF",
};

const FONT = "PingFang SC";
const M = 0.55;
const CONTENT_W = 13.333 - M * 2;

export interface DeckInput {
  quote: Quote;
  financials: FinancialPeriod[];
  klines: Kline[];
  assumptions: DcfAssumptions;
}

export interface DeckSummary {
  model: string;
  fairValuePerShare: number | null;
  impliedGrowthPct: number | null;
}

function addSectionTitle(slide: Slide, title: string, subtitle?: string) {
  slide.addText(title, { x: M, y: 0.34, w: CONTENT_W, h: 0.42, fontSize: 21, bold: true, color: C.ink, fontFace: FONT });
  if (subtitle) {
    slide.addText(subtitle, { x: M, y: 0.78, w: CONTENT_W, h: 0.28, fontSize: 10.5, color: C.muted, fontFace: FONT });
  }
  slide.addShape("line", { x: M, y: 1.1, w: CONTENT_W, h: 0, line: { color: C.line, width: 1 } });
}

function addFooter(slide: Slide, text: string) {
  slide.addText(text, { x: M, y: 7.06, w: CONTENT_W, h: 0.24, fontSize: 8, color: C.muted, fontFace: FONT });
}

/** One headline metric block: big number over a small caption. */
function addMetric(slide: Slide, x: number, y: number, w: number, label: string, value: string, color = C.ink) {
  slide.addShape("roundRect", { x, y, w, h: 1.02, fill: { color: "F7F8FA" }, line: { color: C.line, width: 1 }, rectRadius: 0.06 });
  slide.addText(label, { x: x + 0.16, y: y + 0.12, w: w - 0.32, h: 0.24, fontSize: 9.5, color: C.muted, fontFace: FONT });
  slide.addText(value, { x: x + 0.16, y: y + 0.38, w: w - 0.32, h: 0.46, fontSize: 17, bold: true, color, fontFace: FONT });
}

export async function buildDeck(input: DeckInput, outPath: string): Promise<DeckSummary> {
  const { quote: q, financials, klines, assumptions } = input;
  const latest = financials[0];
  const ttmNp = computeTtm(financials, "netProfit");
  const ttmRev = computeTtm(financials, "revenue");
  const shares = q.totalMarketCap && q.price ? q.totalMarketCap / q.price : null;
  const choice = pickValuationModel(q.name, q.industry, ttmNp.value);

  let fairValuePerShare: number | null = null;
  let valuationBasis = choice.reason;
  let impliedGrowthPct: number | null = null;
  if (choice.model === "dcf") {
    const dcf = runDcf(ttmNp.value, shares, assumptions);
    fairValuePerShare = dcf.fairValuePerShare;
    valuationBasis = dcf.basis;
    const ig = impliedGrowth(q.price, ttmNp.value, shares, assumptions);
    impliedGrowthPct = ig === null ? null : ig * 100;
  } else if (choice.model === "pb-roe") {
    const pb = runPbRoe(latest?.bps ?? null, latest?.roe ?? null, {
      requiredReturn: assumptions.discountRate,
      sustainableGrowth: assumptions.terminalGrowth,
    });
    fairValuePerShare = pb.fairValuePerShare;
    valuationBasis = pb.basis;
  }

  const upside = fairValuePerShare !== null && fairValuePerShare > 0 ? ((fairValuePerShare - q.price) / q.price) * 100 : null;
  const verdict = upside === null ? "估值模型不适用" : upside >= 15 ? "相对低估" : upside <= -15 ? "相对高估" : "估值大致合理";
  const today = new Date().toISOString().slice(0, 10);

  const pptx = new PptxGenJS();
  // LAYOUT_WIDE is 13.333 x 7.5in. LAYOUT_16x9 is only 10 x 5.625in, which is
  // narrower than the coordinate system used below and clips the right column.
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "astock-deck";
  pptx.title = `${q.name}（${q.code}）研究简报`;

  // ---------- 1. 封面 ----------
  const cover = pptx.addSlide();
  cover.background = { color: C.cover };
  cover.addText("A 股研究简报", { x: M, y: 1.15, w: CONTENT_W, h: 0.34, fontSize: 11, color: C.coverSoft, fontFace: FONT, charSpacing: 2 });
  cover.addText(q.name, { x: M, y: 1.6, w: CONTENT_W, h: 0.95, fontSize: 44, bold: true, color: C.white, fontFace: FONT });
  cover.addText(`${q.code}  ·  ${q.industry ?? "A 股"}`, { x: M, y: 2.6, w: CONTENT_W, h: 0.36, fontSize: 15, color: C.coverSoft, fontFace: FONT });
  cover.addShape("line", { x: M, y: 3.25, w: 2.2, h: 0, line: { color: C.accent, width: 3 } });
  cover.addText(
    `${q.price} 元   ·   ${verdict}` + (fairValuePerShare ? `   ·   内在价值 ${fairValuePerShare.toFixed(2)} 元` : ""),
    { x: M, y: 3.6, w: CONTENT_W, h: 0.5, fontSize: 17, color: C.white, fontFace: FONT },
  );
  cover.addText(`数据日期 ${today}　·　行情来源 ${q.source}　·　本简报由 astock-deck 生成，不构成投资建议`, {
    x: M, y: 6.5, w: CONTENT_W, h: 0.3, fontSize: 9.5, color: C.coverSoft, fontFace: FONT,
  });

  // ---------- 2. 核心结论 ----------
  const summary = pptx.addSlide();
  addSectionTitle(summary, "核心结论", `一句话：${verdict}`);
  const gap = 0.22;
  const mw = (CONTENT_W - gap * 3) / 4;
  addMetric(summary, M, 1.45, mw, "最新价", `${q.price} 元`);
  addMetric(summary, M + (mw + gap), 1.45, mw, "PE(TTM)", q.peTtm === null ? "—" : q.peTtm.toFixed(2));
  addMetric(summary, M + (mw + gap) * 2, 1.45, mw, "PB", q.pb === null ? "—" : q.pb.toFixed(2));
  addMetric(summary, M + (mw + gap) * 3, 1.45, mw, "TTM 归母净利", formatMoney(ttmNp.value, "亿元"));

  addMetric(summary, M, 2.7, mw, "TTM 营收", formatMoney(ttmRev.value, "亿元"));
  addMetric(summary, M + (mw + gap), 2.7, mw, "ROE(加权)", formatPct(latest?.roe ?? null));
  addMetric(summary, M + (mw + gap) * 2, 2.7, mw, "毛利率", formatPct(latest?.grossMargin ?? null));
  addMetric(
    summary, M + (mw + gap) * 3, 2.7, mw,
    fairValuePerShare ? "内在价值" : "估值模型",
    fairValuePerShare ? `${fairValuePerShare.toFixed(2)} 元` : "不适用",
    C.accent,
  );

  const bullets: string[] = [];
  bullets.push(`营收同比 ${formatPct(latest?.revenueYoY ?? null)}，净利同比 ${formatPct(latest?.netProfitYoY ?? null)}`);
  if (upside !== null) bullets.push(`现价相对模型内在价值${upside >= 0 ? "折价" : "溢价"} ${Math.abs(upside).toFixed(1)}%`);
  if (impliedGrowthPct !== null) bullets.push(`当前股价隐含 5 年净利润增速约 ${impliedGrowthPct.toFixed(1)}%`);
  summary.addText(bullets.map((b) => ({ text: b, options: { bullet: true } })), {
    x: M, y: 4.15, w: CONTENT_W, h: 1.3, fontSize: 13, color: C.ink, fontFace: FONT, lineSpacingMultiple: 1.5,
  });
  addFooter(summary, `数据来源：${q.source}（公开接口，免 API Key）`);

  // ---------- 3. 财务趋势 ----------
  const trend = pptx.addSlide();
  addSectionTitle(trend, "财务趋势", "各报告期为累计口径，单位为亿元");
  const recent = [...financials].slice(0, 6).reverse();
  const labels = recent.map((p) => p.reportDate.slice(2, 7));
  trend.addChart(
    "bar",
    [
      { name: "营业收入", labels, values: recent.map((p) => Number(((p.revenue ?? 0) / 1e8).toFixed(1))) },
      { name: "归母净利润", labels, values: recent.map((p) => Number(((p.netProfit ?? 0) / 1e8).toFixed(1))) },
    ],
    {
      x: M, y: 1.5, w: CONTENT_W, h: 2.6,
      barDir: "col", barGrouping: "clustered",
      chartColors: [C.accent, "9DB6EA"],
      catAxisLabelColor: C.muted, valAxisLabelColor: C.muted,
      catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
      showLegend: true, legendPos: "t", legendColor: C.muted, legendFontSize: 10,
      showValue: false,
      valGridLine: { color: C.line, style: "solid", size: 0.5 },
    },
  );
  trend.addChart(
    "line",
    [
      { name: "ROE(%)", labels, values: recent.map((p) => Number((p.roe ?? 0).toFixed(2))) },
      { name: "毛利率(%)", labels, values: recent.map((p) => Number((p.grossMargin ?? 0).toFixed(2))) },
    ],
    {
      x: M, y: 4.25, w: CONTENT_W, h: 2.35,
      chartColors: [C.accent, "F0A93B"],
      lineSize: 2.2, lineSmooth: false,
      catAxisLabelColor: C.muted, valAxisLabelColor: C.muted,
      catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
      showLegend: true, legendPos: "t", legendColor: C.muted, legendFontSize: 10,
      valGridLine: { color: C.line, style: "solid", size: 0.5 },
    },
  );
  addFooter(trend, "注：A 股财报为累计口径，本期研究结论采用 TTM 折算值");

  // ---------- 4. 经营与估值 ----------
  const val = pptx.addSlide();
  addSectionTitle(val, "估值测算", `模型：${choice.model === "dcf" ? "两阶段 DCF" : choice.model === "pb-roe" ? "PB-ROE 合理估值" : "不适用"}`);
  const rows: TableRow[] = [
    [
      { text: "项目", options: { bold: true, color: C.white, fill: { color: C.ink } } },
      { text: "数值 / 说明", options: { bold: true, color: C.white, fill: { color: C.ink } } },
    ],
    [{ text: "估值模型选择", options: {} }, { text: choice.reason, options: {} }],
    [{ text: "TTM 归母净利", options: {} }, { text: `${formatMoney(ttmNp.value, "亿元")}（${ttmNp.basis}）`, options: {} }],
    [{ text: "关键假设", options: {} }, { text: `前 ${assumptions.years} 年增速 ${(assumptions.growthRate * 100).toFixed(1)}% · 永续 ${(assumptions.terminalGrowth * 100).toFixed(1)}% · 折现 ${(assumptions.discountRate * 100).toFixed(1)}%`, options: {} }],
    [{ text: "模型内在价值", options: { bold: true } }, { text: fairValuePerShare ? `${fairValuePerShare.toFixed(2)} 元/股` : "—", options: { bold: true, color: C.accent } }],
    [{ text: "相对现价", options: {} }, { text: upside === null ? "—" : `${upside >= 0 ? "折价" : "溢价"} ${Math.abs(upside).toFixed(1)}%`, options: { bold: true, color: C.ink } }],
    [{ text: "现价隐含增速", options: {} }, { text: impliedGrowthPct === null ? "—（该模型不计算）" : `${impliedGrowthPct.toFixed(1)}%（5 年净利润 CAGR）`, options: {} }],
  ];
  val.addTable(rows, {
    x: M, y: 1.45, w: CONTENT_W,
    colW: [2.4, CONTENT_W - 2.4],
    fontSize: 11, fontFace: FONT, color: C.ink,
    border: { type: "solid", color: C.line, pt: 1 },
    align: "left", valign: "middle", rowH: 0.44,
    autoPage: false,
  });
  val.addText(valuationBasis, { x: M, y: 5.35, w: CONTENT_W, h: 0.9, fontSize: 10, color: C.muted, fontFace: FONT });
  addFooter(val, `假设可由用户覆盖：--growth / --terminal / --discount`);

  // ---------- 5. 价格走势 ----------
  if (klines.length >= 20) {
    const price = pptx.addSlide();
    addSectionTitle(price, "价格走势", `最近 ${klines.length} 个交易日（前复权收盘价）`);
    price.addChart(
      "line",
      [{ name: "收盘价", labels: klines.map((k) => k.date.slice(5)), values: klines.map((k) => k.close) }],
      {
        x: M, y: 1.5, w: CONTENT_W, h: 4.6,
        chartColors: [C.accent], lineSize: 2, lineSmooth: false,
        // Plot from just under the window low: a 0-based axis flattens the trend line.
        valAxisMinVal: Math.floor(Math.min(...klines.map((k) => k.low)) * 0.96),
        lineDataSymbol: "none",
        catAxisLabelColor: C.muted, valAxisLabelColor: C.muted,
        catAxisLabelFontSize: 8, valAxisLabelFontSize: 9,
        // The published typings declare this option as string; the runtime expects a number.
        catAxisLabelFrequency: Math.max(1, Math.floor(klines.length / 12)) as unknown as string,
        showLegend: false,
        valGridLine: { color: C.line, style: "solid", size: 0.5 },
      },
    );
    addFooter(price, "数据来源：公开行情接口，可能延迟");
  }

  // ---------- 6. 风险提示 ----------
  const risk = pptx.addSlide();
  addSectionTitle(risk, "风险提示与免责声明");
  const notes = [
    "估值模型基于公开数据与简化假设，对增速、折现率的选取高度敏感，换个参数结论可能反转。",
    "净利润不等于自由现金流，资本开支重的行业 DCF 结果会偏乐观；银行、保险、证券已改用 PB-ROE 模型。",
    "行情与财务数据来自第三方公开接口，可能存在延迟、缺失或口径差异，重大决策请以公司公告为准。",
    "本简报由程序自动生成，不构成任何投资建议，据此操作风险自担。",
  ];
  risk.addText(notes.map((n) => ({ text: n, options: { bullet: true } })), {
    x: M, y: 1.55, w: CONTENT_W, h: 3.2, fontSize: 13, color: C.ink, fontFace: FONT, lineSpacingMultiple: 1.9,
  });
  addFooter(risk, `生成时间 ${today}　·　astock-deck`);

  await pptx.writeFile({ fileName: outPath });
  return { model: choice.model, fairValuePerShare, impliedGrowthPct };
}
