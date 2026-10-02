/* global PowerPoint */

import { FilterRule } from "./excelImport";

// Stored on the presentation itself (not a shape), since filters are a
// global setting. Saved inside the .pptx file, so it survives close/reopen.
const FILTERS_TAG_KEY = "SURVEYSTUDIO_FILTERS";

export async function saveFilters(filters: FilterRule[]): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.tags.add(FILTERS_TAG_KEY, JSON.stringify(filters));
    await context.sync();
  });
}

export async function loadFilters(): Promise<FilterRule[]> {
  return PowerPoint.run(async (context) => {
    const tag = context.presentation.tags.getItemOrNullObject(FILTERS_TAG_KEY);
    tag.load("value,isNullObject");
    await context.sync();

    if (tag.isNullObject) return [];
    try {
      return JSON.parse(tag.value) as FilterRule[];
    } catch {
      return [];
    }
  });
}
