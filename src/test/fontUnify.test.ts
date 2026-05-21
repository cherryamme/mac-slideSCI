import { describe, expect, it } from "vitest";
import { defaultFontUnifySettings, presetFontNames, resolveFontName } from "../shared/settings";

describe("fontUnify", () => {
  it("uses preset font when no custom name", () => {
    expect(resolveFontName({ ...defaultFontUnifySettings, fontName: "Calibri" })).toBe("Calibri");
  });

  it("prefers custom font name when provided", () => {
    expect(resolveFontName({ ...defaultFontUnifySettings, fontName: "Arial", customFontName: "Source Han Serif" })).toBe("Source Han Serif");
  });

  it("falls back to default when both are empty", () => {
    expect(resolveFontName({ ...defaultFontUnifySettings, fontName: "", customFontName: "" })).toBe(defaultFontUnifySettings.fontName);
  });

  it("ships with the documented preset list", () => {
    expect(presetFontNames).toContain("PingFang SC");
    expect(presetFontNames).toContain("Source Han Sans SC");
  });
});
