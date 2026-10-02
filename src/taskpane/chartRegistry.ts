import type { FilterRule } from "./excelImport";

export type ChartTypeId = "bar" | "donut" | "treemap";

export interface ChartFontSizes {
  title: number;
  /** X-axis labels for bar charts; slice/tile labels for donut and treemap charts. */
  xLabels: number;
  /** Y-axis labels for bar charts; unused for donut and treemap charts. */
  yLabels: number;
}

export const DEFAULT_FONT_SIZES: ChartFontSizes = { title: 16, xLabels: 10, yLabels: 10 };

export interface ChartConfig {
  chartType: ChartTypeId;
  columnIndex: number | null;
  color: string;
  name: string;
  title: string;
  categoryOrder: string[];
  /** Per-category bar color, keyed by category name. Falls back to `color` when absent. */
  categoryColors: Record<string, string>;
  fontSizes: ChartFontSizes;
  /** Extra filter rules that apply only to this chart (AND-combined with global + slide filters). */
  chartFilters: FilterRule[];
}

// The config is stored as a tag on the chart's picture shape, so it's saved
// inside the .pptx file itself and survives close/reopen.
export const CHART_CONFIG_TAG_KEY = "SURVEYSTUDIO_CHART_CONFIG";

export function serializeChartConfig(config: ChartConfig): string {
  return JSON.stringify(config);
}

export function parseChartConfig(raw: string): ChartConfig | null {
  try {
    const config = JSON.parse(raw) as ChartConfig;
    // Backfill for configs saved before these fields existed.
    if (!config.categoryColors) config.categoryColors = {};
    if (!config.fontSizes) config.fontSizes = { ...DEFAULT_FONT_SIZES };
    if (!config.chartFilters) config.chartFilters = [];
    return config;
  } catch {
    return null;
  }
}
