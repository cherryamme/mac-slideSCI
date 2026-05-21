import { describe, expect, it } from "vitest";
import { getPngFileName, getSelectedObjectsPngFileName, getZipFileName } from "../features/export/download";

describe("export download naming", () => {
  it("formats PNG filenames with padded slide numbers and DPI", () => {
    expect(getPngFileName(1, 300)).toBe("mac-slideSCI-slide-001-300dpi.png");
    expect(getPngFileName(12, 600)).toBe("mac-slideSCI-slide-012-600dpi.png");
  });

  it("formats selected-object PNG filenames with DPI", () => {
    expect(getSelectedObjectsPngFileName(300)).toBe("mac-slideSCI-selection-300dpi.png");
  });

  it("formats ZIP filenames with DPI", () => {
    expect(getZipFileName(300)).toBe("mac-slideSCI-png-export-300dpi.zip");
  });
});
