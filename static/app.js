const sourceInput = document.querySelector("#sourceFile");
const pivotInput = document.querySelector("#pivotFile");
const monthInput = document.querySelector("#monthInput");
const processButton = document.querySelector("#processButton");
const statusText = document.querySelector("#statusText");
const stats = document.querySelector("#stats");
const downloadLink = document.querySelector("#downloadLink");
const masterStatus = document.querySelector("#masterStatus");
const masterDownload = document.querySelector("#masterDownload");
const searchInput = document.querySelector("#searchInput");
const searchResults = document.querySelector("#searchResults");
const chartTitle = document.querySelector("#chartTitle");
const chartMeta = document.querySelector("#chartMeta");
const chartMount = document.querySelector("#chartMount");

let selectedKey = "";
let searchTimer = 0;

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function formatNumber(value) {
  return new Intl.NumberFormat("ko-KR", { maximumFractionDigits: 2 }).format(value || 0);
}

function nextMonthLabel(monthLabel) {
  const match = String(monthLabel).match(/(\d{4})\D*(\d{1,2})/);
  if (!match) return "다음 달";
  let year = Number(match[1]);
  let month = Number(match[2]) + 1;
  if (month > 12) {
    year += 1;
    month = 1;
  }
  return `${year}년 ${month}월`;
}

function buildForecast(series) {
  const recent = series
    .map((point, index) => ({ index, value: Number(point.value || 0) }))
    .filter((point) => point.value > 0)
    .slice(-3);

  if (recent.length < 3) {
    return null;
  }

  const [oldest, middle, latest] = recent;
  const recentDelta = latest.value - middle.value;
  const previousDelta = middle.value - oldest.value;
  const weightedDelta = recentDelta * 0.66 + previousDelta * 0.33;
  const value = Math.max(0, latest.value + weightedDelta);

  return {
    value,
    latestIndex: latest.index,
    nextIndex: series.length,
  };
}

function setFileName(input, labelId) {
  const label = document.querySelector(labelId);
  if (input.files.length > 1) {
    label.textContent = `${input.files.length}개 파일 선택됨`;
  } else {
    label.textContent = input.files[0]?.name || label.dataset.empty || label.textContent;
  }
}

function setupDropzone(zone) {
  const input = zone.querySelector("input");
  zone.addEventListener("click", () => input.click());
  zone.addEventListener("dragover", (event) => {
    event.preventDefault();
    zone.classList.add("dragging");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragging"));
  zone.addEventListener("drop", (event) => {
    event.preventDefault();
    zone.classList.remove("dragging");
    input.files = event.dataTransfer.files;
    input.dispatchEvent(new Event("change"));
  });
}

document.querySelectorAll(".dropzone").forEach(setupDropzone);
sourceInput.addEventListener("change", () => setFileName(sourceInput, "#sourceName"));
pivotInput.addEventListener("change", () => setFileName(pivotInput, "#pivotName"));

async function refreshStatus() {
  const response = await fetch("/api/status");
  const payload = await response.json();
  if (!payload.ok || !payload.status.exists) {
    masterStatus.textContent = "기준 피벗 없음";
    masterDownload.hidden = true;
    return;
  }
  const { rows, months, updatedAt } = payload.status;
  const firstMonth = months[0] || "";
  const lastMonth = months[months.length - 1] || "";
  masterStatus.textContent = `${rows.toLocaleString("ko-KR")}개 품목 · ${firstMonth} ~ ${lastMonth} · ${updatedAt}`;
  masterDownload.hidden = false;
}

function renderStats(payload) {
  const labels = {
    month: "반영 연월",
    fileCount: "파일",
    sourceRows: "처리 품목",
    updatedRows: "기존 항목",
    newRows: "신약 추가",
    skippedWithoutKd: "제외",
    duplicatedRowsMerged: "중복 병합",
  };
  stats.innerHTML = Object.entries(labels)
    .map(([key, label]) => `<div><dt>${label}</dt><dd>${escapeHtml(payload[key] ?? "")}</dd></div>`)
    .join("");
}

processButton.addEventListener("click", async () => {
  downloadLink.hidden = true;
  stats.innerHTML = "";

  if (!sourceInput.files[0]) {
    statusText.textContent = "월별 원본 엑셀 파일을 선택해주세요.";
    return;
  }

  const formData = new FormData();
  formData.append("month", monthInput.value);
  [...sourceInput.files].forEach((file) => formData.append("source", file));
  if (pivotInput.files[0]) {
    formData.append("pivot", pivotInput.files[0]);
  }

  processButton.disabled = true;
  statusText.textContent = "업데이트 중";

  try {
    const response = await fetch("/update", { method: "POST", body: formData });
    const payload = await response.json();
    if (!payload.ok) throw new Error(payload.error || "처리 중 오류가 발생했습니다.");

    renderStats(payload.stats);
    statusText.textContent = "완료";
    downloadLink.href = payload.file;
    downloadLink.download = payload.filename;
    downloadLink.hidden = false;
    downloadLink.textContent = `${payload.filename}`;
    await refreshStatus();
    if (selectedKey) await loadTrend(selectedKey);
  } catch (error) {
    statusText.textContent = error.message;
  } finally {
    processButton.disabled = false;
  }
});

searchInput.addEventListener("input", () => {
  window.clearTimeout(searchTimer);
  searchTimer = window.setTimeout(searchItems, 120);
});

async function searchItems() {
  const query = searchInput.value.trim();
  if (!query) {
    searchResults.innerHTML = "";
    return;
  }
  const response = await fetch(`/api/search?q=${encodeURIComponent(query)}`);
  const payload = await response.json();
  if (!payload.ok) {
    searchResults.innerHTML = `<div class="empty-state">${escapeHtml(payload.error || "검색 실패")}</div>`;
    return;
  }
  if (!payload.results.length) {
    searchResults.innerHTML = '<div class="empty-state">검색 결과 없음</div>';
    return;
  }
  searchResults.innerHTML = payload.results
    .map(
      (item) => `
        <button class="result-item" type="button" data-key="${escapeHtml(item.key)}">
          <span class="result-name">${escapeHtml(item.처방명)}</span>
          <span class="result-meta">${escapeHtml(item.약품코드)} · KD ${escapeHtml(item.KD코드)} · ${escapeHtml(item.제약사)}</span>
        </button>
      `,
    )
    .join("");

  searchResults.querySelectorAll(".result-item").forEach((button) => {
    button.addEventListener("click", () => {
      selectedKey = button.dataset.key;
      searchResults.querySelectorAll(".result-item").forEach((item) => item.classList.remove("selected"));
      button.classList.add("selected");
      loadTrend(selectedKey);
    });
  });
}

async function loadTrend(key) {
  const response = await fetch(`/api/item?key=${encodeURIComponent(key)}`);
  const payload = await response.json();
  if (!payload.ok) {
    chartMount.innerHTML = `<div class="empty-state">${escapeHtml(payload.error || "그래프 로딩 실패")}</div>`;
    return;
  }
  renderChart(payload.item, payload.series, payload.components || []);
}

function renderChart(item, series, components = []) {
  chartTitle.textContent = item.처방명 || item.약품코드;
  const componentLabel =
    components.length > 1 ? ` · KD ${components.length}개 합산` : ` · KD ${item.KD코드}`;
  chartMeta.textContent = `${item.약품코드}${componentLabel} · ${item.제약사} · ${item.판매사}`;

  const forecast = buildForecast(series);
  const forecastMonth = nextMonthLabel(series.at(-1)?.month || "");
  const forecastValue = forecast ? forecast.value : null;
  const chartSeries = forecast
    ? [...series, { month: forecastMonth, value: forecastValue, changePct: null, forecast: true }]
    : series;
  const width = 920;
  const height = 420;
  const margin = { top: 34, right: 36, bottom: 78, left: 76 };
  const values = chartSeries.map((point) => Number(point.value || 0));
  const maxValue = Math.max(...values, 1);
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const step = chartSeries.length > 1 ? plotWidth / (chartSeries.length - 1) : plotWidth;
  const xFor = (index) => margin.left + index * step;
  const yFor = (value) => margin.top + plotHeight - (value / maxValue) * plotHeight;
  const points = series.map((point, index) => `${xFor(index)},${yFor(point.value)}`).join(" ");
  const componentColors = ["#d06b24", "#5b63b7", "#5b8f3d", "#9a4aa5", "#b84b5f"];
  const componentLines =
    components.length > 1
      ? components
          .map((component, componentIndex) => {
            const componentPoints = component.series
              .map((point, index) => `${xFor(index)},${yFor(point.value)}`)
              .join(" ");
            const color = componentColors[componentIndex % componentColors.length];
            return `
              <polyline class="component-line" style="stroke:${color}" points="${componentPoints}"></polyline>
            `;
          })
          .join("")
      : "";
  const componentLegend =
    components.length > 1
      ? components
          .map((component, componentIndex) => {
            const color = componentColors[componentIndex % componentColors.length];
            const y = margin.top + 16 + componentIndex * 17;
            return `
              <circle cx="${margin.left + 6}" cy="${y - 4}" r="4" fill="${color}"></circle>
              <text class="legend-label" x="${margin.left + 16}" y="${y}" text-anchor="start">KD ${component.item.KD코드}</text>
            `;
          })
          .join("") +
        `<line class="legend-total-line" x1="${margin.left}" y1="${margin.top + 16 + components.length * 17 - 4}" x2="${margin.left + 12}" y2="${margin.top + 16 + components.length * 17 - 4}"></line>
         <text class="legend-label" x="${margin.left + 16}" y="${margin.top + 16 + components.length * 17}" text-anchor="start">합산</text>`
      : "";
  const forecastLine = forecast
    ? `
      <line class="forecast-line" x1="${xFor(forecast.latestIndex)}" y1="${yFor(series[forecast.latestIndex].value)}" x2="${xFor(forecast.nextIndex)}" y2="${yFor(forecastValue)}"></line>
      <circle class="forecast-dot" cx="${xFor(forecast.nextIndex)}" cy="${yFor(forecastValue)}" r="6"></circle>
      <text class="forecast-label" x="${xFor(forecast.nextIndex)}" y="${yFor(forecastValue) - 14}" text-anchor="middle">예측 ${formatNumber(forecastValue)}</text>
    `
    : `<text class="model-label" x="${width - margin.right}" y="${margin.top - 10}" text-anchor="end">예측 계산 데이터 부족</text>`;
  const grid = [0, 0.25, 0.5, 0.75, 1]
    .map((ratio) => {
      const y = margin.top + plotHeight - ratio * plotHeight;
      const value = maxValue * ratio;
      return `
        <line class="grid-line" x1="${margin.left}" y1="${y}" x2="${width - margin.right}" y2="${y}"></line>
        <text class="axis-label" x="${margin.left - 12}" y="${y + 4}" text-anchor="end">${formatNumber(value)}</text>
      `;
    })
    .join("");

  const dots = chartSeries
    .map((point, index) => {
      const x = xFor(index);
      const y = yFor(point.value);
      if (point.forecast) {
        return `
          <g>
            <text class="month-label forecast-month" x="${x}" y="${height - margin.bottom + 36}" text-anchor="end" transform="rotate(-38 ${x} ${height - margin.bottom + 36})">${point.month.replace("년 ", ".").replace("월", "")}</text>
          </g>
        `;
      }
      const change = point.changePct;
      const changeText = change === null || change === undefined ? "" : `${change >= 0 ? "+" : ""}${change.toFixed(1)}%`;
      const changeClass = change > 0 ? "increase" : change < 0 ? "decrease" : "flat";
      return `
        <g>
          <circle class="data-dot" cx="${x}" cy="${y}" r="5"></circle>
          <text class="value-label" x="${x}" y="${y - 13}" text-anchor="middle">${formatNumber(point.value)}</text>
          ${
            changeText
              ? `<text class="change-label ${changeClass}" x="${x}" y="${y + 24}" text-anchor="middle">${changeText}</text>`
              : ""
          }
          <text class="month-label" x="${x}" y="${height - margin.bottom + 36}" text-anchor="end" transform="rotate(-38 ${x} ${height - margin.bottom + 36})">${point.month.replace("년 ", ".").replace("월", "")}</text>
        </g>
      `;
    })
    .join("");

  chartMount.innerHTML = `
    <svg class="trend-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="월별 사용량 그래프">
      ${grid}
      <line class="axis-line" x1="${margin.left}" y1="${margin.top}" x2="${margin.left}" y2="${height - margin.bottom}"></line>
      <line class="axis-line" x1="${margin.left}" y1="${height - margin.bottom}" x2="${width - margin.right}" y2="${height - margin.bottom}"></line>
      <text class="axis-title" x="${width / 2}" y="${height - 12}" text-anchor="middle">연월</text>
      <text class="axis-title" x="18" y="${height / 2}" text-anchor="middle" transform="rotate(-90 18 ${height / 2})">사용량</text>
      ${componentLines}
      <polyline class="trend-line" points="${points}"></polyline>
      ${forecastLine}
      ${componentLegend}
      ${dots}
    </svg>
  `;
}

refreshStatus();
