import { state } from "/ui/shared/dataset/dataset_state.js";
import { formatDatasetOriginLabel } from "/ui/shared/dataset/dataset_origin_labels.js";
import {
  formatDatasetChartValue,
  getDisplayDatasetModel,
} from "/ui/shared/tabs/data/dataset_grid_view.js?v=20260907c";
import {
  renderChart as renderChartCanvas,
  setupChartHover,
} from "/ui/dataset_viewer/tabs/dataset_chart_renderer.js?v=20260724a";

export function renderDatasetChart() {
  const canvas = document.getElementById("devChart");
  if (!canvas) return;
  setupChartHover(canvas);
  const legendEl = document.getElementById("devChartLegend");
  const originLen = Number(document.getElementById("originLenSelect")?.value) || 12;

  const titleEl = document.getElementById("chartTitle");
  if (titleEl) {
    titleEl.textContent = state.chartMode === "byCol"
      ? "By Column (Dev Period)"
      : "Development Curves";
  }

  document.querySelectorAll("#chartModeToggle .chartToggleBtn").forEach((button) => {
    button.classList.toggle("active", button.dataset.mode === state.chartMode);
  });

  renderChartCanvas(canvas, getDisplayDatasetModel(), {
    mode: state.chartMode === "byCol" ? "byCol" : "byRow",
    activeCell: state.activeCell,
    formatValue: formatDatasetChartValue,
    legendEl,
    formatOriginLabel: (label) => formatDatasetOriginLabel(label, originLen),
    originLen,
  });
}

export function redrawDatasetChartSafely() {
  const panel = document.getElementById("chartPanel");
  if (!panel) return;
  const rect = panel.getBoundingClientRect();
  if (rect.width < 50 || rect.height < 50) return;
  renderDatasetChart();
}
