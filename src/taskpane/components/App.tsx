import * as React from "react";
import { useEffect, useRef, useState } from "react";
import Header from "./Header";
import DataTab from "./DataTab";
import FilterTab from "./FilterTab";
import GrafikenTab from "./GrafikenTab";
import ChartEditTab from "./ChartEditTab";
import { Button, Tab, TabList, SelectTabData, SelectTabEvent, ToggleButton, makeStyles } from "@fluentui/react-components";
import { ArrowSync16Regular, Info16Regular, Pin16Regular } from "@fluentui/react-icons";
import { useExcelData } from "../useExcelData";
import { FilterRule } from "../excelImport";
import {
  CHART_CONFIG_TAG_KEY,
  ChartFontSizes,
  ChartTypeId,
  DEFAULT_FONT_SIZES,
  parseChartConfig,
  serializeChartConfig,
} from "../chartRegistry";
import { DEFAULT_BAR_COLOR, generateChartId, insertChart, nextInsertIndex } from "../chartToSlide";
import { ensureUniqueChartId } from "../chartIds";
import { hideInfoOverlay, showInfoOverlay } from "../overlay";
import { loadFilters, saveFilters } from "../filterPersistence";
import { updateAllCharts, updateChartsOnSlide } from "../bulkUpdate";

/* global Office PowerPoint */

interface AppProps {
  title: string;
}

type TabValue = "data" | "filter" | "types" | "edit";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const excelData = useExcelData();
  const [selectedTab, setSelectedTab] = useState<TabValue>("data");
  const [chartType, setChartType] = useState<ChartTypeId>("bar");
  const [selectedColumn, setSelectedColumn] = useState<number | null>(null);
  const [globalFilters, setGlobalFilters] = useState<FilterRule[]>([]);
  const [chartFilters, setChartFilters] = useState<FilterRule[]>([]);
  const [categoryOrder, setCategoryOrder] = useState<string[]>([]);
  const [categoryColors, setCategoryColors] = useState<Record<string, string>>({});
  const [fontSizes, setFontSizes] = useState<ChartFontSizes>(DEFAULT_FONT_SIZES);
  const [color, setColor] = useState(DEFAULT_BAR_COLOR);
  const [name, setName] = useState("");
  const [title, setTitle] = useState("");
  const [editingShapeName, setEditingShapeName] = useState<string | null>(null);
  const [currentSlideId, setCurrentSlideId] = useState<string | null>(null);
  const [overlayVisible, setOverlayVisible] = useState(false);
  const [overlayError, setOverlayError] = useState<string | null>(null);
  const [bulkUpdateStatus, setBulkUpdateStatus] = useState<string | null>(null);
  const [bulkUpdating, setBulkUpdating] = useState(false);
  const [autoOpenEnabled, setAutoOpenEnabled] = useState(false);
  const filtersLoadedRef = useRef(false);
  const bulkUpdatingRef = useRef(false);

  // Load the previously saved filters once on startup (works across
  // close/reopen, since they're stored on the presentation itself).
  useEffect(() => {
    (async () => {
      try {
        const loaded = await loadFilters();
        setGlobalFilters(loaded);
      } catch (e) {
        console.error("Failed to load saved filters", e);
      } finally {
        filtersLoadedRef.current = true;
      }
    })();
  }, []);

  // Persist filter changes. Skipped until the initial load above has run,
  // so it doesn't overwrite the saved filters with the empty initial state.
  useEffect(() => {
    if (!filtersLoadedRef.current) return;
    saveFilters(globalFilters).catch((e) => console.error("Failed to save filters", e));
  }, [globalFilters]);

  // Read the "auto-open task pane with this document" setting, a per-document
  // flag stored in the .pptx file itself (Office.context.document.settings).
  useEffect(() => {
    setAutoOpenEnabled(Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument") === true);
  }, []);

  const handleToggleAutoOpen = () => {
    const next = !autoOpenEnabled;
    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", next);
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to save auto-open setting", result.error);
        return;
      }
      setAutoOpenEnabled(next);
    });
  };

  // Track which slide is currently active, independent of shape selection,
  // so the Filter tab's "this slide" section always follows along. Set once
  // on startup, then kept current by the same selection-changed handler
  // below (which also fires on plain slide navigation).
  useEffect(() => {
    PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items/id");
      await context.sync();
      if (selectedSlides.items.length > 0) {
        setCurrentSlideId(selectedSlides.items[0].id);
      }
    }).catch((e) => console.error("Failed to get the initial active slide", e));
  }, []);

  // When the user selects an existing chart shape on the slide, restore the
  // options that were used to create it so they reappear in the pane. The
  // config is read from a tag on the shape itself, so this also works for
  // charts from a previous session (the tag is saved inside the .pptx file).
  useEffect(() => {
    const handleSelectionChanged = async () => {
      // Each chart updated during a bulk update re-selects itself; ignore
      // those transient selection changes rather than flicker through them.
      if (bulkUpdatingRef.current) return;
      await PowerPoint.run(async (context) => {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items/id");
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/name");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          setCurrentSlideId(selectedSlides.items[0].id);
        }

        if (shapes.items.length !== 1) {
          setEditingShapeName(null);
          return;
        }
        const shape = shapes.items[0];
        const tag = shape.tags.getItemOrNullObject(CHART_CONFIG_TAG_KEY);
        tag.load("value,isNullObject");
        await context.sync();

        const config = tag.isNullObject ? null : parseChartConfig(tag.value);
        if (!config) {
          setEditingShapeName(null);
          return;
        }

        let effectiveConfig = config;

        // A shape copied/duplicated in PowerPoint keeps its tag (so the
        // config survives), but PowerPoint gives the copy a different
        // internal name than the one baked into the tag. Detect that
        // mismatch and assign the copy a fresh ID right away, so two charts
        // never end up sharing a name (which would break our name-based
        // lookups for updates and filters).
        if (config.name !== shape.name) {
          const freshId = await ensureUniqueChartId(generateChartId(), shape.name);
          shape.name = freshId;
          effectiveConfig = { ...config, name: freshId };
          shape.tags.add(CHART_CONFIG_TAG_KEY, serializeChartConfig(effectiveConfig));
          await context.sync();
        }

        setChartType(effectiveConfig.chartType);
        setSelectedColumn(effectiveConfig.columnIndex);
        setCategoryOrder(effectiveConfig.categoryOrder);
        setCategoryColors(effectiveConfig.categoryColors);
        setChartFilters(effectiveConfig.chartFilters);
        setFontSizes(effectiveConfig.fontSizes);
        setColor(effectiveConfig.color);
        setName(effectiveConfig.name);
        setTitle(effectiveConfig.title);
        setEditingShapeName(effectiveConfig.name);
        setSelectedTab("edit");
      });
    };

    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handleSelectionChanged);
    return () => {
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
        handler: handleSelectionChanged,
      });
    };
  }, []);

  const handleCreatePlaceholder = async (chartType: ChartTypeId) => {
    const id = await ensureUniqueChartId(generateChartId());
    const shapeName = await insertChart([], nextInsertIndex(), {
      chartType,
      columnIndex: null,
      color: DEFAULT_BAR_COLOR,
      name: id,
      title: "",
      categoryOrder: [],
      categoryColors: {},
      fontSizes: DEFAULT_FONT_SIZES,
      chartFilters: [],
    });

    setChartType(chartType);
    setSelectedColumn(null);
    setCategoryOrder([]);
    setCategoryColors({});
    setChartFilters([]);
    setFontSizes(DEFAULT_FONT_SIZES);
    setColor(DEFAULT_BAR_COLOR);
    setName(shapeName);
    setTitle("");
    setEditingShapeName(shapeName);
    setSelectedTab("edit");
  };

  const handleUpdateCharts = async (scope: "slide" | "presentation") => {
    setBulkUpdateStatus(null);
    setBulkUpdating(true);
    bulkUpdatingRef.current = true;
    try {
      const count =
        scope === "slide"
          ? currentSlideId
            ? await updateChartsOnSlide(currentSlideId, excelData.grid, excelData.hasHeaderRow, globalFilters)
            : 0
          : await updateAllCharts(excelData.grid, excelData.hasHeaderRow, globalFilters);
      const scopeLabel = scope === "slide" ? "auf dieser Folie" : "in der Präsentation";
      setBulkUpdateStatus(
        count === 0 ? `Keine Grafiken ${scopeLabel} gefunden.` : `${count} Grafik(en) ${scopeLabel} aktualisiert.`
      );
    } catch (e) {
      console.error("Bulk chart update failed", e);
      setBulkUpdateStatus(`Fehler: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      bulkUpdatingRef.current = false;
      setBulkUpdating(false);
    }
  };

  const handleToggleOverlay = async () => {
    try {
      setOverlayError(null);
      if (overlayVisible) {
        await hideInfoOverlay();
        setOverlayVisible(false);
      } else {
        await showInfoOverlay(excelData.grid, excelData.hasHeaderRow, globalFilters);
        setOverlayVisible(true);
      }
    } catch (e) {
      console.error("Toggling info overlay failed", e);
      setOverlayError(e instanceof Error ? e.message : String(e));
    }
  };

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} />
      <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "center", gap: "8px", margin: "8px 0" }}>
        <ToggleButton appearance="subtle" icon={<Info16Regular />} checked={overlayVisible} onClick={handleToggleOverlay}>
          Info-Overlay
        </ToggleButton>
        <ToggleButton appearance="subtle" icon={<Pin16Regular />} checked={autoOpenEnabled} onClick={handleToggleAutoOpen}>
          Beim Öffnen automatisch zeigen
        </ToggleButton>
        <Button
          appearance="subtle"
          icon={<ArrowSync16Regular />}
          disabled={bulkUpdating || !currentSlideId}
          onClick={() => handleUpdateCharts("slide")}
        >
          Alle Grafiken auf Folie aktualisieren
        </Button>
        <Button
          appearance="subtle"
          icon={<ArrowSync16Regular />}
          disabled={bulkUpdating}
          onClick={() => handleUpdateCharts("presentation")}
        >
          Alle Grafiken in Präsentation aktualisieren
        </Button>
      </div>
      {overlayError && <div style={{ color: "#D13438", textAlign: "center" }}>{overlayError}</div>}
      {bulkUpdateStatus && (
        <div
          style={{
            color: bulkUpdateStatus.startsWith("Fehler") ? "#D13438" : "#616161",
            textAlign: "center",
          }}
        >
          {bulkUpdateStatus}
        </div>
      )}
      <TabList
        selectedValue={selectedTab}
        onTabSelect={(_: SelectTabEvent, data: SelectTabData) => setSelectedTab(data.value as TabValue)}
      >
        <Tab value="data">Daten</Tab>
        <Tab value="filter">Filter</Tab>
        <Tab value="types">Grafiken</Tab>
        <Tab value="edit">Diagramm</Tab>
      </TabList>
      {selectedTab === "data" && <DataTab excelData={excelData} />}
      {selectedTab === "filter" && (
        <FilterTab
          excelData={excelData}
          globalFilters={globalFilters}
          onGlobalFiltersChange={setGlobalFilters}
          currentSlideId={currentSlideId}
        />
      )}
      {selectedTab === "types" && <GrafikenTab excelData={excelData} onCreatePlaceholder={handleCreatePlaceholder} />}
      {selectedTab === "edit" && (
        <ChartEditTab
          excelData={excelData}
          globalFilters={globalFilters}
          editingShapeName={editingShapeName}
          onShapeRenamed={setEditingShapeName}
          chartType={chartType}
          selectedColumn={selectedColumn}
          onSelectColumn={setSelectedColumn}
          categoryOrder={categoryOrder}
          onCategoryOrderChange={setCategoryOrder}
          categoryColors={categoryColors}
          onCategoryColorsChange={setCategoryColors}
          chartFilters={chartFilters}
          onChartFiltersChange={setChartFilters}
          fontSizes={fontSizes}
          onFontSizesChange={setFontSizes}
          color={color}
          onColorChange={setColor}
          name={name}
          onNameChange={setName}
          title={title}
          onTitleChange={setTitle}
        />
      )}
    </div>
  );
};

export default App;
