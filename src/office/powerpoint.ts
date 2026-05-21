import { isPowerPointApiSupported } from "./requirements";

export class PowerPointCapabilityError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PowerPointCapabilityError";
  }
}

export function ensurePowerPointApi(version: string, feature: string): void {
  if (!isPowerPointApiSupported(version)) {
    throw new PowerPointCapabilityError(`${feature} 需要 PowerPointApi ${version} 或更高版本。`);
  }
}

export async function runPowerPoint<T>(callback: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T> {
  if (typeof PowerPoint === "undefined") {
    throw new PowerPointCapabilityError("当前页面不在 PowerPoint Office.js 运行环境中。请在 PowerPoint for Mac 中打开插件。");
  }

  return PowerPoint.run(callback);
}
