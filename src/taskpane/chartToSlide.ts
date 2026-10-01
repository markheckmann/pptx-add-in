/* global PowerPoint document Image XMLSerializer SVGSVGElement btoa */

import * as d3 from "d3";

const CHART_WIDTH = 500;
const CHART_HEIGHT = 300;
const PNG_SCALE = 2;

interface BarDatum {
  category: string;
  value: number;
}

function generateRandomData(): BarDatum[] {
  const count = 5 + Math.floor(Math.random() * 4); // 5-8 bars
  return Array.from({ length: count }, (_, i) => ({
    category: `Cat ${i + 1}`,
    value: Math.round(10 + Math.random() * 90),
  }));
}

function renderBarChartSvg(data: BarDatum[]): SVGSVGElement {
  // d3 needs the node attached to the document to measure text for axis ticks,
  // so render into an off-screen host and clone the result before removing it.
  const host = document.createElement("div");
  host.style.position = "absolute";
  host.style.left = "-9999px";
  host.style.top = "-9999px";
  document.body.appendChild(host);

  const margin = { top: 20, right: 20, bottom: 30, left: 40 };

  const svg = d3
    .select(host)
    .append("svg")
    .attr("xmlns", "http://www.w3.org/2000/svg")
    .attr("width", CHART_WIDTH)
    .attr("height", CHART_HEIGHT);

  svg
    .append("rect")
    .attr("width", CHART_WIDTH)
    .attr("height", CHART_HEIGHT)
    .attr("fill", "#ffffff");

  const x = d3
    .scaleBand()
    .domain(data.map((d) => d.category))
    .range([margin.left, CHART_WIDTH - margin.right])
    .padding(0.2);

  const y = d3
    .scaleLinear()
    .domain([0, (d3.max(data, (d) => d.value) ?? 0) * 1.1])
    .range([CHART_HEIGHT - margin.bottom, margin.top]);

  svg
    .append("g")
    .selectAll("rect.bar")
    .data(data)
    .join("rect")
    .attr("x", (d) => x(d.category) ?? 0)
    .attr("y", (d) => y(d.value))
    .attr("width", x.bandwidth())
    .attr("height", (d) => y(0) - y(d.value))
    .attr("fill", "#4472C4");

  svg
    .append("g")
    .attr("transform", `translate(0,${CHART_HEIGHT - margin.bottom})`)
    .call(d3.axisBottom(x))
    .selectAll("text")
    .attr("font-family", "sans-serif")
    .attr("font-size", "10px");

  svg
    .append("g")
    .attr("transform", `translate(${margin.left},0)`)
    .call(d3.axisLeft(y))
    .selectAll("text")
    .attr("font-family", "sans-serif")
    .attr("font-size", "10px");

  const clone = svg.node()!.cloneNode(true) as SVGSVGElement;
  document.body.removeChild(host);
  return clone;
}

async function svgToPngBase64(svgEl: SVGSVGElement): Promise<string> {
  const svgString = new XMLSerializer().serializeToString(svgEl);
  const svgDataUrl = `data:image/svg+xml;charset=utf-8;base64,${btoa(unescape(encodeURIComponent(svgString)))}`;

  const img = new Image();
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error("Failed to rasterize chart SVG"));
    img.src = svgDataUrl;
  });

  const canvas = document.createElement("canvas");
  canvas.width = CHART_WIDTH * PNG_SCALE;
  canvas.height = CHART_HEIGHT * PNG_SCALE;
  const ctx = canvas.getContext("2d")!;
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

  return canvas.toDataURL("image/png").split(",")[1];
}

/**
 * Renders a D3 bar chart with random data and inserts it as a picture on the
 * currently selected slide. `clickIndex` offsets each inserted picture so
 * repeated clicks don't stack exactly on top of one another.
 */
export async function insertRandomBarChart(clickIndex: number): Promise<void> {
  const data = generateRandomData();
  const svgEl = renderBarChartSvg(data);
  const base64Png = await svgToPngBase64(svgEl);

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const step = (clickIndex % 6) * 25;
    const picture = slide.shapes.addPicture(base64Png, {
      left: 50 + step,
      top: 80 + step,
      width: 400,
      height: 240,
    });
    picture.name = `D3BarChart_${Date.now()}`;
    await context.sync();
  });
}
