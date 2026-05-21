import { describe, expect, it } from "vitest";
import { defaultLabelSettings, getPanelLabel } from "../features/labels/labelGenerator";

describe("labelGenerator", () => {
  it("generates panel labels from A by default", () => {
    expect([0, 1, 2].map((index) => getPanelLabel(index, defaultLabelSettings))).toEqual(["A", "B", "C"]);
  });

  it("supports prefixes and custom start index", () => {
    expect(getPanelLabel(0, { ...defaultLabelSettings, prefix: "Fig. ", startIndex: 2 })).toBe("Fig. C");
  });

  it("keeps labels deterministic after Z", () => {
    expect(getPanelLabel(26, defaultLabelSettings)).toBe("A2");
  });
});
