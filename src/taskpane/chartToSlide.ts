/* global PowerPoint document Image XMLSerializer SVGSVGElement btoa */

import * as d3 from "d3";
import {
  CHART_CONFIG_TAG_KEY,
  ChartConfig,
  ChartFontSizes,
  ChartTypeId,
  serializeChartConfig,
} from "./chartRegistry";

const PNG_SCALE = 2;
export const DEFAULT_BAR_COLOR = "#4472C4";
export const DEFAULT_CHART_WIDTH = 400;
export const DEFAULT_CHART_HEIGHT = 240;

const CHART_TYPE_LABELS: Record<ChartTypeId, string> = {
  bar: "Balkendiagramm",
  donut: "Donut-Diagramm",
  treemap: "Treemap",
};

export interface BarDatum {
  category: string;
  value: number;
}

interface TreemapNode {
  category?: string;
  value?: number;
  children?: TreemapNode[];
}

/**
 * Reorders `data` to match `order` (a list of category names). Categories
 * not present in `order` (e.g. new/changed data) are appended at the end,
 * alphabetically.
 */
export function orderBarData(data: BarDatum[], order: string[]): BarDatum[] {
  const byCategory = new Map(data.map((d) => [d.category, d]));
  const ordered: BarDatum[] = [];
  for (const category of order) {
    const d = byCategory.get(category);
    if (d) {
      ordered.push(d);
      byCategory.delete(category);
    }
  }
  const leftovers = Array.from(byCategory.values()).sort((a, b) =>
    a.category.localeCompare(b.category, "de")
  );
  return [...ordered, ...leftovers];
}

function renderBarChartSvg(
  data: BarDatum[],
  defaultColor: string,
  title: string,
  categoryColors: Record<string, string>,
  fontSizes: ChartFontSizes,
  width: number,
  height: number
): SVGSVGElement {
  // d3 needs the node attached to the document to measure text for axis ticks,
  // so render into an off-screen host and clone the result before removing it.
  const host = document.createElement("div");
  host.style.position = "absolute";
  host.style.left = "-9999px";
  host.style.top = "-9999px";
  document.body.appendChild(host);

  const margin = {
    top: title ? Math.max(24, fontSizes.title + 20) : 20,
    right: 20,
    bottom: 30,
    left: 40,
  };

  const svg = d3
    .select(host)
    .append("svg")
    .attr("xmlns", "http://www.w3.org/2000/svg")
    .attr("width", width)
    .attr("height", height);

  svg.append("rect").attr("width", width).attr("height", height).attr("fill", "#ffffff");

  if (title) {
    svg
      .append("text")
      .attr("x", width / 2)
      .attr("y", 24)
      .attr("text-anchor", "middle")
      .attr("font-family", "sans-serif")
      .attr("font-size", `${fontSizes.title}px`)
      .attr("font-weight", "bold")
      .text(title);
  }

  const x = d3
    .scaleBand()
    .domain(data.map((d) => d.category))
    .range([margin.left, width - margin.right])
    .padding(0.2);

  const y = d3
    .scaleLinear()
    .domain([0, (d3.max(data, (d) => d.value) ?? 0) * 1.1])
    .range([height - margin.bottom, margin.top]);

  svg
    .append("g")
    .selectAll("rect.bar")
    .data(data)
    .join("rect")
    .attr("x", (d) => x(d.category) ?? 0)
    .attr("y", (d) => y(d.value))
    .attr("width", x.bandwidth())
    .attr("height", (d) => y(0) - y(d.value))
    .attr("fill", (d) => categoryColors[d.category] ?? defaultColor);

  svg
    .append("g")
    .attr("transform", `translate(0,${height - margin.bottom})`)
    .call(d3.axisBottom(x))
    .selectAll("text")
    .attr("font-family", "sans-serif")
    .attr("font-size", `${fontSizes.xLabels}px`);

  svg
    .append("g")
    .attr("transform", `translate(${margin.left},0)`)
    .call(d3.axisLeft(y))
    .selectAll("text")
    .attr("font-family", "sans-serif")
    .attr("font-size", `${fontSizes.yLabels}px`);

  const clone = svg.node()!.cloneNode(true) as SVGSVGElement;
  document.body.removeChild(host);
  return clone;
}

function renderDonutChartSvg(
  data: BarDatum[],
  defaultColor: string,
  title: string,
  categoryColors: Record<string, string>,
  fontSizes: ChartFontSizes,
  width: number,
  height: number
): SVGSVGElement {
  const host = document.createElement("div");
  host.style.position = "absolute";
  host.style.left = "-9999px";
  host.style.top = "-9999px";
  document.body.appendChild(host);

  const titleHeight = title ? Math.max(24, fontSizes.title + 20) : 0;
  const chartAreaHeight = height - titleHeight;
  const centerX = width / 2;
  const centerY = titleHeight + chartAreaHeight / 2;
  // Leave just enough room for the leader-line labels: horizontally for the
  // label text itself (outer labels sit at radius*1.35 plus text width),
  // vertically just for the line/outer-arc radius (labels have no vertical
  // text stack). Whichever dimension is tighter wins.
  const radius = Math.max(10, Math.min(width / 2 / 1.35 - 30, chartAreaHeight / 2 / 1.1 - 10));

  const svg = d3
    .select(host)
    .append("svg")
    .attr("xmlns", "http://www.w3.org/2000/svg")
    .attr("width", width)
    .attr("height", height);

  svg.append("rect").attr("width", width).attr("height", height).attr("fill", "#ffffff");

  if (title) {
    svg
      .append("text")
      .attr("x", centerX)
      .attr("y", 24)
      .attr("text-anchor", "middle")
      .attr("font-family", "sans-serif")
      .attr("font-size", `${fontSizes.title}px`)
      .attr("font-weight", "bold")
      .text(title);
  }

  const total = d3.sum(data, (d) => d.value) || 1;
  const arcs = d3
    .pie<BarDatum>()
    .value((d) => d.value)
    .sort(null)(data);

  const arc = d3
    .arc<d3.PieArcDatum<BarDatum>>()
    .innerRadius(radius * 0.5)
    .outerRadius(radius);
  const outerArc = d3
    .arc<d3.PieArcDatum<BarDatum>>()
    .innerRadius(radius * 1.1)
    .outerRadius(radius * 1.1);
  const midAngle = (d: d3.PieArcDatum<BarDatum>) => d.startAngle + (d.endAngle - d.startAngle) / 2;

  const g = svg.append("g").attr("transform", `translate(${centerX},${centerY})`);

  g.append("g")
    .selectAll("path")
    .data(arcs)
    .join("path")
    .attr("d", arc)
    .attr("fill", (d) => categoryColors[d.data.category] ?? defaultColor)
    .attr("stroke", "#ffffff")
    .attr("stroke-width", 2);

  g.append("g")
    .selectAll("polyline")
    .data(arcs)
    .join("polyline")
    .attr("points", (d) => {
      const labelPos = outerArc.centroid(d);
      labelPos[0] = radius * 1.3 * (midAngle(d) < Math.PI ? 1 : -1);
      return [arc.centroid(d), outerArc.centroid(d), labelPos].map((p) => p.join(",")).join(" ");
    })
    .attr("stroke", "#999999")
    .attr("stroke-width", 1)
    .attr("fill", "none")
    .attr("opacity", 0.6);

  g.append("g")
    .selectAll("text")
    .data(arcs)
    .join("text")
    .attr("transform", (d) => {
      const labelPos = outerArc.centroid(d);
      labelPos[0] = radius * 1.35 * (midAngle(d) < Math.PI ? 1 : -1);
      return `translate(${labelPos})`;
    })
    .attr("text-anchor", (d) => (midAngle(d) < Math.PI ? "start" : "end"))
    .attr("dy", "0.35em")
    .attr("font-family", "sans-serif")
    .attr("font-size", `${fontSizes.xLabels}px`)
    .text((d) => `${d.data.category} (${Math.round((d.data.value / total) * 100)}%)`);

  const clone = svg.node()!.cloneNode(true) as SVGSVGElement;
  document.body.removeChild(host);
  return clone;
}

function renderTreemapSvg(
  data: BarDatum[],
  defaultColor: string,
  title: string,
  categoryColors: Record<string, string>,
  fontSizes: ChartFontSizes,
  width: number,
  height: number
): SVGSVGElement {
  const host = document.createElement("div");
  host.style.position = "absolute";
  host.style.left = "-9999px";
  host.style.top = "-9999px";
  document.body.appendChild(host);

  const titleHeight = title ? Math.max(24, fontSizes.title + 20) : 0;
  const chartAreaHeight = height - titleHeight;

  const svg = d3
    .select(host)
    .append("svg")
    .attr("xmlns", "http://www.w3.org/2000/svg")
    .attr("width", width)
    .attr("height", height);

  svg.append("rect").attr("width", width).attr("height", height).attr("fill", "#ffffff");

  if (title) {
    svg
      .append("text")
      .attr("x", width / 2)
      .attr("y", 24)
      .attr("text-anchor", "middle")
      .attr("font-family", "sans-serif")
      .attr("font-size", `${fontSizes.title}px`)
      .attr("font-weight", "bold")
      .text(title);
  }

  const root = d3.hierarchy<TreemapNode>({ children: data }).sum((d) => d.value ?? 0);
  const laidOut = d3.treemap<TreemapNode>().size([width, chartAreaHeight]).paddingInner(2)(root);
  const leaves = laidOut.leaves();

  const g = svg.append("g").attr("transform", `translate(0,${titleHeight})`);

  const cell = g
    .selectAll("g")
    .data(leaves)
    .join("g")
    .attr("transform", (d) => `translate(${d.x0},${d.y0})`);

  cell
    .append("rect")
    .attr("width", (d) => d.x1 - d.x0)
    .attr("height", (d) => d.y1 - d.y0)
    .attr("fill", (d) => categoryColors[d.data.category ?? ""] ?? defaultColor);

  cell
    .filter((d) => d.x1 - d.x0 > 30 && d.y1 - d.y0 > 14)
    .append("text")
    .attr("x", 4)
    .attr("y", 16)
    .attr("font-family", "sans-serif")
    .attr("font-size", `${fontSizes.xLabels}px`)
    .attr("font-weight", "bold")
    .attr("fill", "#ffffff")
    .text((d) => d.data.category ?? "");

  const clone = svg.node()!.cloneNode(true) as SVGSVGElement;
  document.body.removeChild(host);
  return clone;
}

const PLACEHOLDER_FILL = "#E0E0E0";
const PLACEHOLDER_BORDER = "#9E9E9E";
const PLACEHOLDER_TEXT = "#616161";

function renderPlaceholderSvg(
  chartType: ChartTypeId,
  width: number,
  height: number
): SVGSVGElement {
  const host = document.createElement("div");
  host.style.position = "absolute";
  host.style.left = "-9999px";
  host.style.top = "-9999px";
  document.body.appendChild(host);

  const svg = d3
    .select(host)
    .append("svg")
    .attr("xmlns", "http://www.w3.org/2000/svg")
    .attr("width", width)
    .attr("height", height);

  const borderWidth = 2;
  svg
    .append("rect")
    .attr("x", borderWidth / 2)
    .attr("y", borderWidth / 2)
    .attr("width", width - borderWidth)
    .attr("height", height - borderWidth)
    .attr("fill", PLACEHOLDER_FILL)
    .attr("stroke", PLACEHOLDER_BORDER)
    .attr("stroke-width", borderWidth)
    .attr("stroke-dasharray", "8,5");

  svg
    .append("text")
    .attr("x", width / 2)
    .attr("y", height / 2)
    .attr("text-anchor", "middle")
    .attr("dy", "0.35em")
    .attr("font-family", "sans-serif")
    .attr("font-size", "20px")
    .attr("font-weight", "bold")
    .attr("fill", PLACEHOLDER_TEXT)
    .text(CHART_TYPE_LABELS[chartType]);

  const clone = svg.node()!.cloneNode(true) as SVGSVGElement;
  document.body.removeChild(host);
  return clone;
}

/**
 * Dispatches to the right renderer for `config.chartType`, at the given
 * logical SVG size (callers pass the actual on-slide size so a manually
 * resized shape keeps its aspect ratio instead of being stretched/distorted
 * on the next update).
 */
function renderChartSvg(
  data: BarDatum[],
  config: ChartConfig,
  width: number,
  height: number
): SVGSVGElement {
  if (data.length === 0) {
    return renderPlaceholderSvg(config.chartType, width, height);
  }
  if (config.chartType === "donut") {
    return renderDonutChartSvg(
      data,
      config.color,
      config.title,
      config.categoryColors,
      config.fontSizes,
      width,
      height
    );
  }
  if (config.chartType === "treemap") {
    return renderTreemapSvg(
      data,
      config.color,
      config.title,
      config.categoryColors,
      config.fontSizes,
      width,
      height
    );
  }
  return renderBarChartSvg(
    data,
    config.color,
    config.title,
    config.categoryColors,
    config.fontSizes,
    width,
    height
  );
}

async function svgToPngBase64(
  svgEl: SVGSVGElement,
  width: number,
  height: number
): Promise<string> {
  const svgString = new XMLSerializer().serializeToString(svgEl);
  const svgDataUrl = `data:image/svg+xml;charset=utf-8;base64,${btoa(unescape(encodeURIComponent(svgString)))}`;

  const img = new Image();
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error("Failed to rasterize chart SVG"));
    img.src = svgDataUrl;
  });

  const canvas = document.createElement("canvas");
  canvas.width = width * PNG_SCALE;
  canvas.height = height * PNG_SCALE;
  const ctx = canvas.getContext("2d")!;
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

  return canvas.toDataURL("image/png").split(",")[1];
}

let insertCounter = 0;

/** Shared offset counter so charts inserted from different UI entry points don't stack. */
export function nextInsertIndex(): number {
  return insertCounter++;
}

let chartIdCounter = 0;

// ponytail: counter resets each session, so a reopened file could in theory
// already contain "id_1" etc. from a previous session — the ID field is
// user-editable, so rename on collision. Upgrade to scanning existing shape
// names if that turns out to matter in practice.
export function generateChartId(): string {
  chartIdCounter += 1;
  return `id_${chartIdCounter}`;
}

/**
 * Renders a D3 chart (bar or donut, per `config.chartType`) for the given
 * data and inserts it as a picture on the currently selected slide.
 * `clickIndex` offsets each inserted picture so repeated inserts don't stack
 * exactly on top of one another. `config` is stored as a tag on the shape
 * (persisted in the .pptx file), so the pane can restore it later —
 * including after the file is closed and reopened. Returns the generated
 * shape name.
 */
export async function insertChart(
  data: BarDatum[],
  clickIndex: number,
  config: ChartConfig
): Promise<string> {
  const svgEl = renderChartSvg(data, config, DEFAULT_CHART_WIDTH, DEFAULT_CHART_HEIGHT);
  const base64Png = await svgToPngBase64(svgEl, DEFAULT_CHART_WIDTH, DEFAULT_CHART_HEIGHT);
  const shapeName = config.name.trim() || generateChartId();

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const step = (clickIndex % 6) * 25;
    const picture = slide.shapes.addPicture(base64Png, {
      left: 50 + step,
      top: 80 + step,
      width: DEFAULT_CHART_WIDTH,
      height: DEFAULT_CHART_HEIGHT,
    });
    picture.name = shapeName;
    picture.tags.add(CHART_CONFIG_TAG_KEY, serializeChartConfig({ ...config, name: shapeName }));
    await context.sync();
  });

  return shapeName;
}

/**
 * Re-renders the chart (bar or donut, per `config.chartType`) and replaces
 * the shape with the given name: deletes it and inserts a fresh picture at
 * the same position/size on the same slide (same approach as insertChart,
 * which is known to work — unlike ShapeFill.setImage, which is less certain
 * on a preview-created shape). Looks the shape up by name across all slides
 * rather than relying on the host's current selection, since a freshly
 * inserted shape isn't necessarily selected. If `config.name` differs from
 * the shape's current name, the new shape uses it (the name doubles as the
 * shape's ID). The config is (re-)persisted as a tag on the new shape.
 * Returns the resulting name, so callers can keep their local state in sync.
 */
export async function updateChart(
  shapeName: string,
  data: BarDatum[],
  config: ChartConfig
): Promise<string> {
  let resultingName = "";

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    slides.items.forEach((slide) => slide.shapes.load("items/name"));
    await context.sync();

    let targetSlide: PowerPoint.Slide | undefined;
    let shape: PowerPoint.Shape | undefined;
    for (const slide of slides.items) {
      const match = slide.shapes.items.find((s) => s.name === shapeName);
      if (match) {
        targetSlide = slide;
        shape = match;
        break;
      }
    }
    if (!targetSlide || !shape) {
      throw new Error(`Diagramm "${shapeName}" wurde auf keiner Folie gefunden.`);
    }

    shape.load("left,top,width,height");
    await context.sync();
    const { left, top, width, height } = shape;

    // Render at the shape's current on-slide size so a manually resized
    // (stretched) chart keeps its aspect ratio instead of being distorted.
    const svgEl = renderChartSvg(data, config, width, height);
    const base64Png = await svgToPngBase64(svgEl, width, height);

    shape.delete();

    resultingName = config.name.trim() || shapeName;
    const picture = targetSlide.shapes.addPicture(base64Png, { left, top, width, height });
    picture.name = resultingName;
    picture.tags.add(
      CHART_CONFIG_TAG_KEY,
      serializeChartConfig({ ...config, name: resultingName })
    );
    picture.load("id");
    await context.sync();

    // Replacing a shape drops the host's selection, so re-select the new one.
    targetSlide.setSelectedShapes([picture.id]);
    await context.sync();
  });

  return resultingName;
}
