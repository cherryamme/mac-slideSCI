import type { OfficeReadyInfo, RequirementStatus } from "../shared/types";

const POWERPOINT_REQUIREMENTS = ["1.1", "1.4", "1.6", "1.8", "1.10"];

export function isOfficeRuntimeAvailable(): boolean {
  return typeof Office !== "undefined" && typeof Office.context !== "undefined";
}

export function isPowerPointApiSupported(version: string): boolean {
  if (!isOfficeRuntimeAvailable()) {
    return false;
  }

  return Office.context.requirements.isSetSupported("PowerPointApi", version);
}

export function getRequirementStatuses(): RequirementStatus[] {
  return POWERPOINT_REQUIREMENTS.map((version) => ({
    name: "PowerPointApi",
    version,
    supported: isPowerPointApiSupported(version),
  }));
}

export function getOfficeHostLabel(info?: OfficeReadyInfo): string {
  if (!info?.host) {
    return isOfficeRuntimeAvailable() ? String(Office.context.host) : "Browser preview";
  }

  return String(info.host);
}

export function getOfficePlatformLabel(info?: OfficeReadyInfo): string {
  if (!info?.platform) {
    return isOfficeRuntimeAvailable() ? String(Office.context.platform) : "Local browser";
  }

  return String(info.platform);
}
