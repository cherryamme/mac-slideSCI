import { describe, expect, it } from "vitest";
import { normalizeSelectedSlidePageIndices } from "../office/exportPdf";

describe("exportPdf", () => {
  it("converts 1-based Office slide indices to 0-based PDF page indices", () => {
    const slides = [
      { id: 2, title: "Second", index: 2 },
      { id: 1, title: "First", index: 1 },
      { id: 2, title: "Second again", index: 2 },
    ];

    expect(normalizeSelectedSlidePageIndices(slides)).toEqual([0, 1]);
  });

  it("ignores invalid slide indices", () => {
    const slides = [
      { id: 0, title: "Bad", index: 0 },
      { id: 1, title: "Good", index: 3 },
    ];

    expect(normalizeSelectedSlidePageIndices(slides)).toEqual([2]);
  });
});
