/* global PowerPoint */

import { SheetGrid, FilterRule, columnLabel } from "./excelImport";
import { CHART_CONFIG_TAG_KEY, parseChartConfig } from "./chartRegistry";
import { SLIDE_FILTERS_TAG_KEY, parseSlideFilters } from "./slideFilters";

// Marks every shape this overlay draws, so hideInfoOverlay() can find and
// remove exactly those again without touching the user's own shapes.
export const OVERLAY_TAG_KEY = "SURVEYSTUDIO_OVERLAY";

const FRAME_COLOR = "#D13438";
const LABEL_FONT_SIZE = 10;

function formatFilterList(grid: SheetGrid, hasHeaderRow: boolean, filters: FilterRule[]): string {
  if (filters.length === 0) return "keine";
  return filters
    .map(
      (f) =>
        `${columnLabel(grid, f.columnIndex, hasHeaderRow)}: ${f.selectedValues.join(", ") || "(nichts ausgewählt)"}`
    )
    .join(" | ");
}

function buildFilterSummary(
  grid: SheetGrid,
  hasHeaderRow: boolean,
  globalFilters: FilterRule[],
  slideFilters: FilterRule[]
): string {
  return (
    `Global — ${formatFilterList(grid, hasHeaderRow, globalFilters)}` +
    `  |  Seite — ${formatFilterList(grid, hasHeaderRow, slideFilters)}`
  );
}

/**
 * Draws a red frame + ID label above every chart shape (found via the chart
 * config tag), plus one filter-summary text box per slide that has at least
 * one chart. Clears any previous overlay first, so this is safe to call
 * repeatedly (e.g. to refresh after filters changed).
 */
export async function showInfoOverlay(
  grid: SheetGrid,
  hasHeaderRow: boolean,
  globalFilters: FilterRule[]
): Promise<void> {
  await hideInfoOverlay();

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    slides.items.forEach((slide) => {
      slide.shapes.load("items/name,items/left,items/top,items/width,items/height");
      slide.tags.load("items/key,items/value");
    });
    await context.sync();

    const allShapes = slides.items.flatMap((slide) =>
      slide.shapes.items.map((shape) => ({ slide, shape }))
    );
    allShapes.forEach(({ shape }) => shape.tags.load("items/key,items/value"));
    await context.sync();

    const slidesWithCharts = new Set<PowerPoint.Slide>();

    for (const { slide, shape } of allShapes) {
      const configTag = shape.tags.items.find((t) => t.key === CHART_CONFIG_TAG_KEY);
      if (!configTag) continue;
      const config = parseChartConfig(configTag.value);
      if (!config) continue;

      slidesWithCharts.add(slide);

      const frame = slide.shapes.addGeometricShape("Rectangle", {
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      });
      frame.fill.clear();
      frame.lineFormat.color = FRAME_COLOR;
      frame.lineFormat.weight = 2;
      frame.tags.add(OVERLAY_TAG_KEY, "1");

      const variableName =
        config.columnIndex !== null
          ? columnLabel(grid, config.columnIndex, hasHeaderRow)
          : "(keine Variable)";
      const label = slide.shapes.addTextBox(`${config.name} — ${variableName}`, {
        left: shape.left,
        top: Math.max(0, shape.top - 18),
        width: shape.width,
        height: 16,
      });
      label.fill.clear();
      label.lineFormat.visible = false;
      label.textFrame.textRange.font.size = LABEL_FONT_SIZE;
      label.textFrame.textRange.font.color = FRAME_COLOR;
      label.textFrame.textRange.font.bold = true;
      label.tags.add(OVERLAY_TAG_KEY, "1");
    }

    for (const slide of slidesWithCharts) {
      const slideFilterTag = slide.tags.items.find((t) => t.key === SLIDE_FILTERS_TAG_KEY);
      const slideFilters = parseSlideFilters(slideFilterTag?.value);
      const filterText = buildFilterSummary(grid, hasHeaderRow, globalFilters, slideFilters);

      const info = slide.shapes.addTextBox(filterText, {
        left: 10,
        top: 10,
        width: 460,
        height: 20,
      });
      info.fill.setSolidColor("#FFFFFF");
      info.lineFormat.color = FRAME_COLOR;
      info.lineFormat.weight = 1;
      info.textFrame.textRange.font.size = LABEL_FONT_SIZE;
      info.tags.add(OVERLAY_TAG_KEY, "1");
    }

    await context.sync();
  });
}

/** Removes every shape tagged with OVERLAY_TAG_KEY, across all slides. */
export async function hideInfoOverlay(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    slides.items.forEach((slide) => slide.shapes.load("items/name"));
    await context.sync();

    const allShapes = slides.items.flatMap((slide) => slide.shapes.items);
    allShapes.forEach((shape) => shape.tags.load("items/key"));
    await context.sync();

    for (const shape of allShapes) {
      if (shape.tags.items.some((t) => t.key === OVERLAY_TAG_KEY)) {
        shape.delete();
      }
    }

    await context.sync();
  });
}
