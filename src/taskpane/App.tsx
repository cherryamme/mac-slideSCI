import React from "react";
import { addDefaultCaptions, addDefaultPanelLabels, alignSelection, arrangeSelectionGrid, arrangeSelectionHorizontal, arrangeSelectionVertical, copySelectionGeometry, deleteSavedCombo, distributeSelection, equalizeSelection, insertArrow, insertCallout, insertLine, insertSavedCombo, listSavedCombos, pasteSelectionGeometry, renameSavedCombo, saveSelectionAsCombo, unifyFontAllSlides, unifyFontCurrentSlide, unifyFontSelection } from "../shared/actions";
import { CAPTION_STORAGE_KEY, GEOMETRY_STORAGE_KEY, LABEL_STORAGE_KEY, loadCaptionSettings, loadFontUnifySettings, loadGeometryMemory, loadLabelSettings, loadLayoutSettings, presetFontNames, saveFontUnifySettings, saveLayoutSettings } from "../shared/settings";
import { saveJson } from "../shared/storage";
import type { AlignmentMode, CaptionSettings, DistributionMode, EqualizeMode, FontHorizontalAlignment, FontUnifySettings, GeometryApplyMode, GeometryMemory, LabelSettings, LayoutSettings, OfficeReadyInfo, OperationResult, SavedCombo } from "../shared/types";

interface AppProps {
  officeInfo?: OfficeReadyInfo;
}

type TaskpaneSection = "geometry" | "layout" | "labels" | "font" | "combos";

function getInitialSection(): TaskpaneSection | null {
  const section = new URLSearchParams(window.location.search).get("section");
  return section === "geometry" || section === "layout" || section === "labels" || section === "font" || section === "combos" ? section : null;
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function NumberField(props: { label: string; value: number; min?: number; step?: number; onChange: (value: number) => void }) {
  return (
    <label className="field">
      <span>{props.label}</span>
      <input type="number" min={props.min} step={props.step ?? 1} value={props.value} onChange={(event) => props.onChange(Number(event.target.value))} />
    </label>
  );
}

function TextField(props: { label: string; value: string; onChange: (value: string) => void }) {
  return (
    <label className="field field-wide">
      <span>{props.label}</span>
      <input value={props.value} onChange={(event) => props.onChange(event.target.value)} />
    </label>
  );
}

function ColorField(props: { label: string; value: string; onChange: (value: string) => void; allowTransparent?: boolean }) {
  return (
    <label className="field">
      <span>{props.label}</span>
      <div className="color-field">
        {props.allowTransparent ? <button className="ghost-button" type="button" onClick={() => props.onChange("transparent")}>透明</button> : null}
        <input type="color" value={props.value === "transparent" ? "#ffffff" : props.value} onChange={(event) => props.onChange(event.target.value)} />
      </div>
    </label>
  );
}

function Button(props: { children: React.ReactNode; onClick: () => void; disabled?: boolean }) {
  return <button type="button" onClick={props.onClick} disabled={props.disabled}>{props.children}</button>;
}

function IconButton(props: { icon: string; label: string; onClick: () => void; disabled?: boolean }) {
  return (
    <button className="icon-button" type="button" title={props.label} aria-label={props.label} onClick={props.onClick} disabled={props.disabled}>
      {props.icon}
    </button>
  );
}

const alignButtons: Array<{ mode: AlignmentMode; icon: string; label: string }> = [
  { mode: "left", icon: "⟸", label: "左对齐" },
  { mode: "center", icon: "↔", label: "水平居中" },
  { mode: "right", icon: "⟹", label: "右对齐" },
  { mode: "top", icon: "⇑", label: "顶对齐" },
  { mode: "middle", icon: "↕", label: "垂直居中" },
  { mode: "bottom", icon: "⇓", label: "底对齐" },
];

const distributeButtons: Array<{ mode: DistributionMode; icon: string; label: string }> = [
  { mode: "horizontal", icon: "⇔", label: "水平均分" },
  { mode: "vertical", icon: "⇕", label: "垂直均分" },
];

const equalizeButtons: Array<{ mode: EqualizeMode; icon: string; label: string }> = [
  { mode: "width", icon: "W", label: "等宽" },
  { mode: "height", icon: "H", label: "等高" },
  { mode: "both", icon: "□", label: "等宽高" },
];

export function App(_props: AppProps) {
  const [message, setMessage] = React.useState("常用工具在顶部 mac-slideSCI ribbon；这里用于设置和组合库。");
  const [busy, setBusy] = React.useState(false);
  const [layoutSettings, setLayoutSettings] = React.useState<LayoutSettings>(() => loadLayoutSettings());
  const [labelSettings, setLabelSettings] = React.useState<LabelSettings>(() => loadLabelSettings());
  const [captionSettings, setCaptionSettings] = React.useState<CaptionSettings>(() => loadCaptionSettings());
  const [geometryMemory, setGeometryMemory] = React.useState<GeometryMemory | null>(() => loadGeometryMemory());
  const [fontUnifySettings, setFontUnifySettings] = React.useState<FontUnifySettings>(() => loadFontUnifySettings());
  const [combos, setCombos] = React.useState<SavedCombo[]>([]);
  const [comboName, setComboName] = React.useState("");
  const [comboExpanded, setComboExpanded] = React.useState(false);
  const [comboMenu, setComboMenu] = React.useState<{ id: string; x: number; y: number } | null>(null);
  const [comboRename, setComboRename] = React.useState<{ id: string; name: string } | null>(null);
  const [comboStatus, setComboStatus] = React.useState("选中对象后点击保存，会按名称写入本机组合库。");
  const [fontStatus, setFontStatus] = React.useState("选择字体后再点击范围按钮，结果会在这里显示。");
  const [targetSection] = React.useState<TaskpaneSection | null>(() => getInitialSection());
  const [highlightedSection, setHighlightedSection] = React.useState<TaskpaneSection | null>(targetSection);
  const geometryRef = React.useRef<HTMLElement>(null);
  const layoutRef = React.useRef<HTMLElement>(null);
  const labelsRef = React.useRef<HTMLElement>(null);
  const fontRef = React.useRef<HTMLElement>(null);
  const combosRef = React.useRef<HTMLElement>(null);

  React.useEffect(() => {
    if (!targetSection) {
      return;
    }

    const refs: Record<TaskpaneSection, React.RefObject<HTMLElement>> = {
      geometry: geometryRef,
      layout: layoutRef,
      labels: labelsRef,
      font: fontRef,
      combos: combosRef,
    };
    window.setTimeout(() => refs[targetSection].current?.scrollIntoView({ block: "start", behavior: "smooth" }), 80);
    const timer = window.setTimeout(() => setHighlightedSection(null), 1600);
    return () => window.clearTimeout(timer);
  }, [targetSection]);

  React.useEffect(() => {
    saveLayoutSettings(layoutSettings);
  }, [layoutSettings]);

  React.useEffect(() => {
    saveJson(LABEL_STORAGE_KEY, labelSettings);
  }, [labelSettings]);

  React.useEffect(() => {
    saveJson(CAPTION_STORAGE_KEY, captionSettings);
  }, [captionSettings]);

  React.useEffect(() => {
    saveJson(GEOMETRY_STORAGE_KEY, geometryMemory);
  }, [geometryMemory]);

  React.useEffect(() => {
    saveFontUnifySettings(fontUnifySettings);
  }, [fontUnifySettings]);

  const refreshCombos = React.useCallback(async (): Promise<void> => {
    try {
      const items = await listSavedCombos();
      setCombos(items);
    } catch (error) {
      setMessage(getErrorMessage(error));
    }
  }, []);

  React.useEffect(() => {
    void refreshCombos();
  }, [refreshCombos]);

  React.useEffect(() => {
    if (!comboMenu) {
      return;
    }
    function close(): void {
      setComboMenu(null);
    }
    window.addEventListener("click", close);
    window.addEventListener("scroll", close, true);
    return () => {
      window.removeEventListener("click", close);
      window.removeEventListener("scroll", close, true);
    };
  }, [comboMenu]);

  async function runOperation(action: () => Promise<OperationResult | void>): Promise<void> {
    setBusy(true);
    try {
      const result = await action();
      setGeometryMemory(loadGeometryMemory());
      if (result) {
        setMessage(result.message);
      }
    } catch (error) {
      setMessage(getErrorMessage(error));
    } finally {
      setBusy(false);
    }
  }

  function updateLayoutSettings(partial: Partial<LayoutSettings>): void {
    setLayoutSettings((current) => ({ ...current, ...partial }));
  }

  function updateLabelSettings(partial: Partial<LabelSettings>): void {
    setLabelSettings((current) => ({ ...current, ...partial }));
  }

  function updateCaptionSettings(partial: Partial<CaptionSettings>): void {
    setCaptionSettings((current) => ({ ...current, ...partial }));
  }

  function updateFontUnifySettings(partial: Partial<FontUnifySettings>): void {
    setFontUnifySettings((current) => ({ ...current, ...partial }));
  }

  async function applyFont(scope: "selection" | "slide" | "presentation"): Promise<void> {
    setBusy(true);
    setFontStatus("正在统一字体...");
    try {
      saveFontUnifySettings(fontUnifySettings);
      const action = scope === "selection" ? unifyFontSelection : scope === "slide" ? unifyFontCurrentSlide : unifyFontAllSlides;
      const result = await action();
      setFontStatus(result.message);
      setMessage(result.message);
    } catch (error) {
      const errorMessage = getErrorMessage(error);
      setFontStatus(errorMessage);
      setMessage(errorMessage);
    } finally {
      setBusy(false);
    }
  }

  async function saveCombo(): Promise<void> {
    setBusy(true);
    setComboStatus("正在保存组合...");
    try {
      const result = await saveSelectionAsCombo(comboName);
      setComboName("");
      await refreshCombos();
      setComboStatus(result.message);
      setMessage(result.message);
    } catch (error) {
      const errorMessage = getErrorMessage(error);
      setComboStatus(errorMessage);
      setMessage(errorMessage);
    } finally {
      setBusy(false);
    }
  }

  async function insertCombo(id: string): Promise<void> {
    setBusy(true);
    setComboStatus("正在插入组合...");
    try {
      const result = await insertSavedCombo(id);
      setComboStatus(result.message);
      setMessage(result.message);
    } catch (error) {
      const errorMessage = getErrorMessage(error);
      setComboStatus(errorMessage);
      setMessage(errorMessage);
    } finally {
      setBusy(false);
    }
  }

  async function removeCombo(id: string): Promise<void> {
    setBusy(true);
    try {
      const result = await deleteSavedCombo(id);
      await refreshCombos();
      setComboStatus(result.message);
      setMessage(result.message);
    } catch (error) {
      const errorMessage = getErrorMessage(error);
      setComboStatus(errorMessage);
      setMessage(errorMessage);
    } finally {
      setBusy(false);
    }
  }

  async function applyComboRename(): Promise<void> {
    if (!comboRename) {
      return;
    }

    setBusy(true);
    try {
      const result = await renameSavedCombo(comboRename.id, comboRename.name);
      setComboRename(null);
      await refreshCombos();
      setComboStatus(result.message);
      setMessage(result.message);
    } catch (error) {
      const errorMessage = getErrorMessage(error);
      setComboStatus(errorMessage);
      setMessage(errorMessage);
    } finally {
      setBusy(false);
    }
  }

  function openComboMenu(event: React.MouseEvent, combo: SavedCombo): void {
    event.preventDefault();
    setComboMenu({ id: combo.id, x: event.clientX, y: event.clientY });
  }

  async function copyGeometry(): Promise<void> {
    await runOperation(copySelectionGeometry);
  }

  async function applyGeometry(mode: GeometryApplyMode): Promise<void> {
    await runOperation(() => pasteSelectionGeometry(mode));
  }

  function sectionClass(section: TaskpaneSection, extra = ""): string {
    return `panel ${highlightedSection === section ? "panel-highlight" : ""} ${extra}`.trim();
  }

  const disabled = busy;

  return (
    <main className="app-shell compact-shell">
      <p className="message" role="status">{busy ? "正在处理..." : message}</p>

      <section ref={geometryRef} className={sectionClass("geometry")}>
        <div className="panel-heading"><h2>位置尺寸与标注</h2><p>顶部 ribbon 也可直接使用这些一键工具。</p></div>
        {geometryMemory ? <p className="memory-preview">left {geometryMemory.left.toFixed(1)}, top {geometryMemory.top.toFixed(1)}, width {geometryMemory.width.toFixed(1)}, height {geometryMemory.height.toFixed(1)}</p> : <p className="memory-preview">尚未复制对象几何信息。</p>}
        <div className="button-row">
          <Button disabled={disabled} onClick={() => void copyGeometry()}>复制几何</Button>
          <Button disabled={disabled || !geometryMemory} onClick={() => void applyGeometry("position")}>粘贴位置</Button>
          <Button disabled={disabled || !geometryMemory} onClick={() => void applyGeometry("size")}>粘贴尺寸</Button>
          <Button disabled={disabled || !geometryMemory} onClick={() => void applyGeometry("all")}>粘贴全部</Button>
        </div>
        <div className="icon-toolbar">
          <IconButton icon="/" label="插入线条" disabled={disabled} onClick={() => void runOperation(insertLine)} />
          <IconButton icon="→" label="插入箭头" disabled={disabled} onClick={() => void runOperation(insertArrow)} />
          <IconButton icon="!" label="插入 Callout" disabled={disabled} onClick={() => void runOperation(insertCallout)} />
        </div>
      </section>

      <section ref={layoutRef} className={sectionClass("layout")}>
        <div className="panel-heading"><h2>布局默认设置</h2><p>顶部 ribbon 的横排、竖排和网格会使用这些设置。</p></div>
        <div className="field-grid">
          <NumberField label="网格列数" min={1} value={layoutSettings.columns} onChange={(columns) => updateLayoutSettings({ columns })} />
          <NumberField label="横向间距" min={0} value={layoutSettings.horizontalSpacing} onChange={(horizontalSpacing) => updateLayoutSettings({ horizontalSpacing })} />
          <NumberField label="纵向间距" min={0} value={layoutSettings.verticalSpacing} onChange={(verticalSpacing) => updateLayoutSettings({ verticalSpacing })} />
          <label className="checkbox-field"><input type="checkbox" checked={layoutSettings.equalizeGrid} onChange={(event) => updateLayoutSettings({ equalizeGrid: event.target.checked })} />网格等宽高</label>
        </div>
        <div className="button-row">
          <Button disabled={disabled} onClick={() => void runOperation(arrangeSelectionHorizontal)}>横排</Button>
          <Button disabled={disabled} onClick={() => void runOperation(arrangeSelectionVertical)}>竖排</Button>
          <Button disabled={disabled} onClick={() => void runOperation(arrangeSelectionGrid)}>网格</Button>
        </div>
        <div className="toolbar-block">
          <span>对齐</span>
          <div className="icon-toolbar">{alignButtons.map((button) => <IconButton key={button.mode} icon={button.icon} label={button.label} disabled={disabled} onClick={() => void runOperation(() => alignSelection(button.mode))} />)}</div>
        </div>
        <div className="toolbar-block">
          <span>均分 / 等尺寸</span>
          <div className="icon-toolbar">
            {distributeButtons.map((button) => <IconButton key={button.mode} icon={button.icon} label={button.label} disabled={disabled} onClick={() => void runOperation(() => distributeSelection(button.mode))} />)}
            {equalizeButtons.map((button) => <IconButton key={button.mode} icon={button.icon} label={button.label} disabled={disabled} onClick={() => void runOperation(() => equalizeSelection(button.mode))} />)}
          </div>
        </div>
      </section>

      <section ref={labelsRef} className={sectionClass("labels")}>
        <div className="panel-heading"><h2>Panel label 与 caption 设置</h2><p>顶部 ribbon 的 A/B/C 和 Caption 按钮会使用这些参数。</p></div>
        <div className="field-grid">
          <TextField label="Label 前缀" value={labelSettings.prefix} onChange={(prefix) => updateLabelSettings({ prefix })} />
          <NumberField label="起始序号" min={0} value={labelSettings.startIndex} onChange={(startIndex) => updateLabelSettings({ startIndex })} />
          <NumberField label="字号" min={6} value={labelSettings.fontSize} onChange={(fontSize) => updateLabelSettings({ fontSize })} />
          <NumberField label="X 偏移" value={labelSettings.offsetX} onChange={(offsetX) => updateLabelSettings({ offsetX })} />
          <NumberField label="Y 偏移" value={labelSettings.offsetY} onChange={(offsetY) => updateLabelSettings({ offsetY })} />
          <ColorField label="文字颜色" value={labelSettings.color} onChange={(color) => updateLabelSettings({ color })} />
          <ColorField label="背景" value={labelSettings.backgroundColor} allowTransparent onChange={(backgroundColor) => updateLabelSettings({ backgroundColor })} />
          <label className="checkbox-field"><input type="checkbox" checked={labelSettings.bold} onChange={(event) => updateLabelSettings({ bold: event.target.checked })} />加粗</label>
        </div>
        <div className="field-grid">
          <TextField label="Caption 文本" value={captionSettings.text} onChange={(text) => updateCaptionSettings({ text })} />
          <NumberField label="Caption 字号" min={6} value={captionSettings.fontSize} onChange={(fontSize) => updateCaptionSettings({ fontSize })} />
          <NumberField label="Caption 间距" min={0} value={captionSettings.gap} onChange={(gap) => updateCaptionSettings({ gap })} />
          <ColorField label="Caption 颜色" value={captionSettings.color} onChange={(color) => updateCaptionSettings({ color })} />
        </div>
        <div className="button-row">
          <Button disabled={disabled} onClick={() => void runOperation(addDefaultPanelLabels)}>添加 A/B/C 标签</Button>
          <Button disabled={disabled} onClick={() => void runOperation(addDefaultCaptions)}>添加 Caption</Button>
        </div>
      </section>

      <section ref={fontRef} className={sectionClass("font")}>
        <div className="panel-heading"><h2>字体统一</h2><p>选择字体后对选中对象、本页或全部页统一应用。</p></div>
        <div className="field-grid">
          <label className="field">
            <span>预设字体</span>
            <select value={fontUnifySettings.fontName} onChange={(event) => updateFontUnifySettings({ fontName: event.target.value })}>
              {presetFontNames.map((name) => <option key={name} value={name}>{name}</option>)}
            </select>
          </label>
          <TextField label="自定义字体名" value={fontUnifySettings.customFontName} onChange={(customFontName) => updateFontUnifySettings({ customFontName })} />
          <label className="checkbox-field"><input type="checkbox" checked={fontUnifySettings.unifyAllFormatting} onChange={(event) => updateFontUnifySettings({ unifyAllFormatting: event.target.checked })} />统一所有格式（字体大小、段落等）</label>
        </div>
        {fontUnifySettings.unifyAllFormatting ? (
          <div className="field-grid">
            <NumberField label="字号" min={6} step={1} value={fontUnifySettings.fontSize} onChange={(fontSize) => updateFontUnifySettings({ fontSize })} />
            <label className="checkbox-field"><input type="checkbox" checked={fontUnifySettings.bold} onChange={(event) => updateFontUnifySettings({ bold: event.target.checked })} />加粗</label>
            <label className="checkbox-field"><input type="checkbox" checked={fontUnifySettings.italic} onChange={(event) => updateFontUnifySettings({ italic: event.target.checked })} />斜体</label>
            <label className="field">
              <span>段落对齐</span>
              <select value={fontUnifySettings.horizontalAlignment} onChange={(event) => updateFontUnifySettings({ horizontalAlignment: event.target.value as FontHorizontalAlignment })}>
                <option value="left">左对齐</option>
                <option value="center">居中</option>
                <option value="right">右对齐</option>
                <option value="justify">两端对齐</option>
              </select>
            </label>
          </div>
        ) : null}
        <div className="button-row">
          <Button disabled={disabled} onClick={() => void applyFont("selection")}>统一选中对象</Button>
          <Button disabled={disabled} onClick={() => void applyFont("slide")}>统一本页</Button>
          <Button disabled={disabled} onClick={() => void applyFont("presentation")}>统一所有页</Button>
        </div>
        <p className="export-result" role="status">{fontStatus}</p>
      </section>

      <section ref={combosRef} className={sectionClass("combos")}>
        <div className="panel-heading"><h2>组合库</h2><p>保存选中对象的文本、形状和分组结构，跨 PPT 复用。</p></div>
        <div className="section-note">支持保存文本框、几何形状、线条以及它们的分组；插入时会按原有的层级和相对位置重建结构。</div>
        <div className="field-grid">
          <TextField label="组合名称" value={comboName} onChange={setComboName} />
        </div>
        <div className="button-row">
          <Button disabled={disabled} onClick={() => void saveCombo()}>保存当前选区</Button>
          <Button disabled={disabled} onClick={() => setComboExpanded((value) => !value)}>{comboExpanded ? "收起全部" : "展开全部"}</Button>
        </div>
        <p className="export-result" role="status">{comboStatus}</p>
        {combos.length === 0 ? (
          <p className="memory-preview">尚未保存任何组合。</p>
        ) : (
          <>
            <div className="combo-strip">
              {combos.map((combo) => (
                <button key={combo.id} className="combo-card" type="button" onClick={() => void insertCombo(combo.id)} onContextMenu={(event) => openComboMenu(event, combo)} disabled={disabled}>
                  {combo.thumbnailBase64 ? <img className="combo-thumb" src={combo.thumbnailBase64} alt={combo.name} /> : <div className="combo-thumb combo-thumb-fallback">无预览</div>}
                  <span className="combo-name">{combo.name}</span>
                </button>
              ))}
              <button className="combo-expand" type="button" onClick={() => setComboExpanded((value) => !value)} title="展开全部" aria-label="展开全部">▾</button>
            </div>
            {comboExpanded ? (
              <div className="combo-grid">
                {combos.map((combo) => (
                  <button key={combo.id} className="combo-card" type="button" onClick={() => void insertCombo(combo.id)} onContextMenu={(event) => openComboMenu(event, combo)} disabled={disabled}>
                    {combo.thumbnailBase64 ? <img className="combo-thumb" src={combo.thumbnailBase64} alt={combo.name} /> : <div className="combo-thumb combo-thumb-fallback">无预览</div>}
                    <span className="combo-name">{combo.name}</span>
                  </button>
                ))}
              </div>
            ) : null}
          </>
        )}
        {comboRename ? (
          <div className="field-grid">
            <TextField label="重命名" value={comboRename.name} onChange={(name) => setComboRename(comboRename ? { ...comboRename, name } : null)} />
            <div className="button-row">
              <Button disabled={disabled} onClick={() => void applyComboRename()}>保存</Button>
              <Button disabled={disabled} onClick={() => setComboRename(null)}>取消</Button>
            </div>
          </div>
        ) : null}
        {comboMenu ? (
          <div className="combo-menu" style={{ left: comboMenu.x, top: comboMenu.y }}>
            <button type="button" onClick={() => { const id = comboMenu.id; setComboMenu(null); void insertCombo(id); }}>插入到当前页</button>
            <button type="button" onClick={() => { const combo = combos.find((item) => item.id === comboMenu.id); setComboMenu(null); if (combo) { setComboRename({ id: combo.id, name: combo.name }); } }}>重命名</button>
            <button type="button" onClick={() => { const id = comboMenu.id; setComboMenu(null); void removeCombo(id); }}>删除</button>
          </div>
        ) : null}
      </section>

      <footer className="plugin-about">
        mac-slideSCI 是面向 PowerPoint for Mac 的 SCI PPT 辅助插件，常用工具在顶部 ribbon，右侧面板用于参数设置、字体统一和组合管理。
      </footer>
    </main>
  );
}
