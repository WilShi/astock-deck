#!/usr/bin/env node
/** astock-deck command line entry. */

import { getFinancials, getKline, getQuote, type Kline } from "./datasource/index.ts";
import { buildDeck } from "./deck.ts";

const USAGE = [
  "astock-deck — 把 A 股数据直接生成原生可编辑 PPTX 研究简报",
  "",
  "用法：astock-deck <股票代码> [选项]",
  "",
  "选项：",
  "  -o, --out <文件>      输出路径，默认 <代码>-研究简报.pptx",
  "  --growth <r>          净利润增速假设，默认 0.08",
  "  --terminal <r>        永续增速假设，默认 0.025",
  "  --discount <r>        折现率假设，默认 0.10",
  "  --years <n>           显式预测年数，默认 5",
  "  --bars <n>            走势图取的交易日数，默认 120",
  "  -h, --help            显示帮助",
  "",
  "示例：",
  "  astock-deck 600519",
  "  astock-deck 600519 -o 茅台.pptx --discount 0.12",
].join("\n");

interface Opts {
  code: string;
  out: string;
  growth: number;
  terminal: number;
  discount: number;
  years: number;
  bars: number;
}

function parse(argv: string[]): Opts | null {
  const positional: string[] = [];
  const flags: Record<string, string> = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "-h" || a === "--help") return null;
    if (a === "-o" || a === "--out") { flags.out = argv[++i] ?? ""; continue; }
    if (a.startsWith("--")) { flags[a.slice(2)] = argv[++i] ?? ""; continue; }
    positional.push(a);
  }
  const code = positional[0];
  if (!code) throw new Error("缺少股票代码，例如 astock-deck 600519");

  const num = (key: string, fallback: number): number => {
    const raw = flags[key];
    if (raw === undefined) return fallback;
    const n = Number(raw);
    if (!Number.isFinite(n)) throw new Error(`--${key} 需要数字，收到 "${raw}"`);
    return n;
  };

  return {
    code,
    out: flags.out || `${code}-研究简报.pptx`,
    growth: num("growth", 0.08),
    terminal: num("terminal", 0.025),
    discount: num("discount", 0.10),
    years: Math.max(1, Math.round(num("years", 5))),
    bars: Math.min(Math.max(Math.round(num("bars", 120)), 20), 250),
  };
}

async function main(argv: string[]): Promise<number> {
  const opts = parse(argv);
  if (!opts) { console.log(USAGE); return 0; }

  process.stderr.write(`正在获取 ${opts.code} 的数据…\n`);
  const [quote, financials, klines] = await Promise.all([
    getQuote(opts.code),
    getFinancials(opts.code, 8),
    getKline(opts.code, opts.bars).catch(() => [] as Kline[]),
  ]);

  process.stderr.write(`正在生成 PPTX…\n`);
  const summary = await buildDeck(
    {
      quote, financials, klines,
      assumptions: {
        growthRate: opts.growth,
        terminalGrowth: opts.terminal,
        discountRate: opts.discount,
        years: opts.years,
      },
    },
    opts.out,
  );

  console.log(`已生成：${opts.out}`);
  console.log(`  ${quote.name}（${quote.code}）  ${quote.price} 元`);
  console.log(`  估值模型：${summary.model}`);
  if (summary.fairValuePerShare) console.log(`  内在价值：${summary.fairValuePerShare.toFixed(2)} 元/股`);
  if (summary.impliedGrowthPct !== null) console.log(`  现价隐含增速：${summary.impliedGrowthPct.toFixed(1)}%`);
  return 0;
}

main(process.argv.slice(2)).then(
  (code) => process.exit(code),
  (err: unknown) => { console.error(`错误：${(err as Error).message}`); process.exit(1); },
);
