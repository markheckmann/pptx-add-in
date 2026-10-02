/* global PowerPoint */

import { FilterRule } from "./excelImport";

// Stored on the slide itself, so a per-slide filter travels with the slide
// and survives close/reopen.
export const SLIDE_FILTERS_TAG_KEY = "SURVEYSTUDIO_SLIDE_FILTERS";

export function parseSlideFilters(raw: string | undefined): FilterRule[] {
  if (!raw) return [];
  try {
    return JSON.parse(raw) as FilterRule[];
  } catch {
    return [];
  }
}

export async function loadSlideFilters(slideId: string): Promise<FilterRule[]> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const slide = slides.items.find((s) => s.id === slideId);
    if (!slide) return [];

    const tag = slide.tags.getItemOrNullObject(SLIDE_FILTERS_TAG_KEY);
    tag.load("value,isNullObject");
    await context.sync();

    return tag.isNullObject ? [] : parseSlideFilters(tag.value);
  });
}

export async function saveSlideFilters(slideId: string, filters: FilterRule[]): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const slide = slides.items.find((s) => s.id === slideId);
    if (!slide) return;

    slide.tags.add(SLIDE_FILTERS_TAG_KEY, JSON.stringify(filters));
    await context.sync();
  });
}

/** Finds the slide that currently hosts the given shape and returns its filters. */
export async function getSlideFiltersForShape(shapeName: string): Promise<FilterRule[]> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    slides.items.forEach((slide) => {
      slide.shapes.load("items/name");
      slide.tags.load("items/key,items/value");
    });
    await context.sync();

    const targetSlide = slides.items.find((slide) =>
      slide.shapes.items.some((s) => s.name === shapeName)
    );
    if (!targetSlide) return [];

    const tag = targetSlide.tags.items.find((t) => t.key === SLIDE_FILTERS_TAG_KEY);
    return parseSlideFilters(tag?.value);
  });
}
