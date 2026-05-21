import { addDefaultCaptions, addDefaultPanelLabels, alignSelection, arrangeSelectionGrid, arrangeSelectionHorizontal, arrangeSelectionVertical, copySelectionGeometry, distributeSelection, equalizeSelection, insertArrow, insertCallout, insertLine, pasteSelectionGeometry, saveSelectionAsCombo, unifyFontAllSlides, unifyFontCurrentSlide, unifyFontSelection } from "../shared/actions";
import type { AlignmentMode, DistributionMode, EqualizeMode, GeometryApplyMode, OperationResult } from "../shared/types";

type OfficeEvent = { completed: () => void };
type CommandAction = () => Promise<OperationResult>;

function logCommandError(name: string, error: unknown): void {
  const message = error instanceof Error ? error.message : String(error);
  console.warn(`[mac-slideSCI] ${name} failed: ${message}`);
}

function associateCommand(name: string, action: CommandAction): void {
  Office.actions.associate(name, async (event: OfficeEvent) => {
    try {
      await action();
    } catch (error) {
      logCommandError(name, error);
    } finally {
      event.completed();
    }
  });
}

function align(mode: AlignmentMode): CommandAction {
  return () => alignSelection(mode);
}

function distribute(mode: DistributionMode): CommandAction {
  return () => distributeSelection(mode);
}

function equalize(mode: EqualizeMode): CommandAction {
  return () => equalizeSelection(mode);
}

function paste(mode: GeometryApplyMode): CommandAction {
  return () => pasteSelectionGeometry(mode);
}

Office.onReady()
  .then(() => {
    associateCommand("arrangeHorizontalCommand", arrangeSelectionHorizontal);
    associateCommand("arrangeVerticalCommand", arrangeSelectionVertical);
    associateCommand("arrangeGridCommand", arrangeSelectionGrid);
    associateCommand("alignLeftCommand", align("left"));
    associateCommand("alignCenterCommand", align("center"));
    associateCommand("alignRightCommand", align("right"));
    associateCommand("alignTopCommand", align("top"));
    associateCommand("alignMiddleCommand", align("middle"));
    associateCommand("alignBottomCommand", align("bottom"));
    associateCommand("distributeHorizontalCommand", distribute("horizontal"));
    associateCommand("distributeVerticalCommand", distribute("vertical"));
    associateCommand("equalizeWidthCommand", equalize("width"));
    associateCommand("equalizeHeightCommand", equalize("height"));
    associateCommand("equalizeBothCommand", equalize("both"));
    associateCommand("addPanelLabelsCommand", addDefaultPanelLabels);
    associateCommand("addCaptionsCommand", addDefaultCaptions);
    associateCommand("copyGeometryCommand", copySelectionGeometry);
    associateCommand("pastePositionCommand", paste("position"));
    associateCommand("pasteSizeCommand", paste("size"));
    associateCommand("pasteGeometryCommand", paste("all"));
    associateCommand("insertLineCommand", insertLine);
    associateCommand("insertArrowCommand", insertArrow);
    associateCommand("insertCalloutCommand", insertCallout);
    associateCommand("unifyFontSelectionCommand", unifyFontSelection);
    associateCommand("unifyFontCurrentSlideCommand", unifyFontCurrentSlide);
    associateCommand("unifyFontAllSlidesCommand", unifyFontAllSlides);
    associateCommand("saveComboCommand", () => saveSelectionAsCombo());
  })
  .catch((error) => logCommandError("Office.onReady", error));
