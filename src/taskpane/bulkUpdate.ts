/* global PowerPoint */

import { SheetGrid, FilterRule, applyFilters, buildFrequencyData } from "./excelImport";
import { CHART_CONFIG_TAG_KEY, ChartConfig, parseChartConfig } from "./chartRegistry";
import { orderBarData, updateChart } from "./chartToSlide";
import { SLIDE_FILTERS_TAG_KEY, parseSlideFilters } from "./slideFilters";

interface ChartEntry {
  shapeName: string;
  config: ChartConfig;
  slideId: string;
  slideFilters: FilterRule[];
}

async function findAllCharts(): Promise<ChartEntry[]> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    slides.items.forEach((slide) => {
      slide.shapes.load("items/name");
      slide.tags.load("items/key,items/value");
    });
    await context.sync();

    const allShapes = slides.items.flatMap((slide) => slide.shapes.items);
    allShapes.forEach((shape) => shape.tags.load("items/key,items/value"));
    await context.sync();

    const found: ChartEntry[] = [];
    for (const slide of slides.items) {
      const slideFilterTag = slide.tags.items.find((t) => t.key === SLIDE_FILTERS_TAG_KEY);
      const slideFilters = parseSlideFilters(slideFilterTag?.value);

      for (const shape of slide.shapes.items) {
        const tag = shape.tags.items.find((t) => t.key === CHART_CONFIG_TAG_KEY);
        if (!tag) continue;
        const config = parseChartConfig(tag.value);
        if (config) found.push({ shapeName: shape.name, config, slideId: slide.id, slideFilters });
      }
    }
    return found;
  });
}

async function updateCharts(
  charts: ChartEntry[],
  grid: SheetGrid,
  hasHeaderRow: boolean,
  globalFilters: FilterRule[]
): Promise<number> {
  let updatedCount = 0;
  for (const { shapeName, config, slideFilters } of charts) {
    if (config.columnIndex === null) continue;
    const effectiveFilters = [...globalFilters, ...slideFilters, ...config.chartFilters];
    const filteredGrid = applyFilters(grid, effectiveFilters, hasHeaderRow);
    const data = orderBarData(
      buildFrequencyData(filteredGrid, config.columnIndex, hasHeaderRow),
      config.categoryOrder
    );
    if (data.length === 0) continue;

    await updateChart(shapeName, data, config);
    updatedCount++;
  }
  return updatedCount;
}

/**
 * Re-renders every chart shape found across all slides (identified via the
 * chart config tag) using its own saved config, against the current grid
 * and the full filter stack for that chart (global + its slide's filter +
 * its own chart-specific filter, all AND-combined) — e.g. after adjusting
 * a global filter. Returns how many charts were updated.
 */
export async function updateAllCharts(
  grid: SheetGrid,
  hasHeaderRow: boolean,
  globalFilters: FilterRule[]
): Promise<number> {
  const charts = await findAllCharts();
  return updateCharts(charts, grid, hasHeaderRow, globalFilters);
}

/** Same as updateAllCharts, but limited to charts on the given slide. */
export async function updateChartsOnSlide(
  slideId: string,
  grid: SheetGrid,
  hasHeaderRow: boolean,
  globalFilters: FilterRule[]
): Promise<number> {
  const charts = (await findAllCharts()).filter((c) => c.slideId === slideId);
  return updateCharts(charts, grid, hasHeaderRow, globalFilters);
}
