/* global PowerPoint */

async function listAllShapeNames(): Promise<Set<string>> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    slides.items.forEach((slide) => slide.shapes.load("items/name"));
    await context.sync();

    const names = new Set<string>();
    for (const slide of slides.items) {
      for (const shape of slide.shapes.items) {
        names.add(shape.name);
      }
    }
    return names;
  });
}

/**
 * Returns `desiredId` if no other shape in the presentation already uses it
 * (ignoring `excludeShapeName` — typically the shape being renamed itself,
 * since renaming a shape to its own current name isn't a collision).
 * Otherwise appends/increments a numeric suffix until the name is free.
 * Checked against *all* shape names, not just chart shapes, since a
 * collision with any shape would confuse our name-based lookups.
 */
export async function ensureUniqueChartId(
  desiredId: string,
  excludeShapeName?: string
): Promise<string> {
  const usedNames = await listAllShapeNames();
  if (excludeShapeName) usedNames.delete(excludeShapeName);

  if (!usedNames.has(desiredId)) return desiredId;

  let suffix = 2;
  while (usedNames.has(`${desiredId}_${suffix}`)) suffix++;
  return `${desiredId}_${suffix}`;
}
