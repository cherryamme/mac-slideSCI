import type { GridLayoutOptions, ShapeBox, AlignmentMode, DistributionMode, EqualizeMode } from "../../shared/types";

export function sortByPosition(boxes: ShapeBox[]): ShapeBox[] {
  return [...boxes].sort((a, b) => {
    const rowDelta = a.top - b.top;
    if (Math.abs(rowDelta) > 8) {
      return rowDelta;
    }
    return a.left - b.left;
  });
}

export function arrangeHorizontal(boxes: ShapeBox[], spacing: number): ShapeBox[] {
  const sorted = sortByPosition(boxes);
  if (sorted.length === 0) {
    return [];
  }

  let nextLeft = Math.min(...sorted.map((box) => box.left));
  const top = Math.min(...sorted.map((box) => box.top));

  return sorted.map((box) => {
    const next = { ...box, left: nextLeft, top };
    nextLeft += box.width + spacing;
    return next;
  });
}

export function arrangeVertical(boxes: ShapeBox[], spacing: number): ShapeBox[] {
  const sorted = sortByPosition(boxes);
  if (sorted.length === 0) {
    return [];
  }

  const left = Math.min(...sorted.map((box) => box.left));
  let nextTop = Math.min(...sorted.map((box) => box.top));

  return sorted.map((box) => {
    const next = { ...box, left, top: nextTop };
    nextTop += box.height + spacing;
    return next;
  });
}

export function arrangeGrid(boxes: ShapeBox[], options: GridLayoutOptions): ShapeBox[] {
  const sorted = sortByPosition(boxes);
  if (sorted.length === 0) {
    return [];
  }

  const columns = Math.max(1, Math.min(options.columns, sorted.length));
  const originLeft = Math.min(...sorted.map((box) => box.left));
  const originTop = Math.min(...sorted.map((box) => box.top));
  const cellWidth = Math.max(...sorted.map((box) => box.width));
  const cellHeight = Math.max(...sorted.map((box) => box.height));

  return sorted.map((box, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    return {
      ...box,
      left: originLeft + column * (cellWidth + options.spacing.horizontal),
      top: originTop + row * (cellHeight + options.spacing.vertical),
      width: options.equalizeSize ? cellWidth : box.width,
      height: options.equalizeSize ? cellHeight : box.height,
    };
  });
}

export function alignShapes(boxes: ShapeBox[], mode: AlignmentMode): ShapeBox[] {
  if (boxes.length === 0) {
    return [];
  }

  const minLeft = Math.min(...boxes.map((box) => box.left));
  const minTop = Math.min(...boxes.map((box) => box.top));
  const maxRight = Math.max(...boxes.map((box) => box.left + box.width));
  const maxBottom = Math.max(...boxes.map((box) => box.top + box.height));
  const centerX = minLeft + (maxRight - minLeft) / 2;
  const centerY = minTop + (maxBottom - minTop) / 2;

  return boxes.map((box) => {
    switch (mode) {
      case "left":
        return { ...box, left: minLeft };
      case "center":
        return { ...box, left: centerX - box.width / 2 };
      case "right":
        return { ...box, left: maxRight - box.width };
      case "top":
        return { ...box, top: minTop };
      case "middle":
        return { ...box, top: centerY - box.height / 2 };
      case "bottom":
        return { ...box, top: maxBottom - box.height };
    }
  });
}

export function distributeShapes(boxes: ShapeBox[], mode: DistributionMode): ShapeBox[] {
  if (boxes.length < 3) {
    return boxes;
  }

  const sorted = [...boxes].sort((a, b) => (mode === "horizontal" ? a.left - b.left : a.top - b.top));

  if (mode === "horizontal") {
    const minLeft = sorted[0].left;
    const maxRight = sorted[sorted.length - 1].left + sorted[sorted.length - 1].width;
    const totalWidth = sorted.reduce((sum, box) => sum + box.width, 0);
    const gap = (maxRight - minLeft - totalWidth) / (sorted.length - 1);
    let nextLeft = minLeft;

    return sorted.map((box) => {
      const next = { ...box, left: nextLeft };
      nextLeft += box.width + gap;
      return next;
    });
  }

  const minTop = sorted[0].top;
  const maxBottom = sorted[sorted.length - 1].top + sorted[sorted.length - 1].height;
  const totalHeight = sorted.reduce((sum, box) => sum + box.height, 0);
  const gap = (maxBottom - minTop - totalHeight) / (sorted.length - 1);
  let nextTop = minTop;

  return sorted.map((box) => {
    const next = { ...box, top: nextTop };
    nextTop += box.height + gap;
    return next;
  });
}

export function equalizeShapes(boxes: ShapeBox[], mode: EqualizeMode): ShapeBox[] {
  if (boxes.length === 0) {
    return [];
  }

  const maxWidth = Math.max(...boxes.map((box) => box.width));
  const maxHeight = Math.max(...boxes.map((box) => box.height));

  return boxes.map((box) => ({
    ...box,
    width: mode === "height" ? box.width : maxWidth,
    height: mode === "width" ? box.height : maxHeight,
  }));
}
