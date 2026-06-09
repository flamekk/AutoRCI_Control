(() => {
  const chartColors = {
    OK: "#198754",
    MANQUANTE_RCI: "#f59e0b",
    HORS_SCOPE_RCI: "#6c757d",
    RCI_HORS_PERIODE: "#0dcaf0",
    RCI_SEULEMENT: "#6f42c1",
    ANOMALIE_MONTANT: "#dc3545",
    ANOMALIE_DATE: "#dc3545",
    DOUBLON: "#8b1e3f",
    CRITIQUE: "#dc3545",
    ELEVEE: "#f59e0b",
    MOYENNE: "#f2c94c",
    A_VERIFIER: "#6c757d",
    INFORMATION: "#0dcaf0",
  };

  const fallbackPalette = [
    "#0f4c81",
    "#1976d2",
    "#198754",
    "#f59e0b",
    "#6f42c1",
    "#0dcaf0",
    "#6c757d",
    "#dc3545",
    "#8b1e3f",
    "#7c3aed",
  ];

  const endpointCache = new Map();
  const numberFormatter = new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 0 });
  const moneyFormatter = new Intl.NumberFormat("fr-FR", {
    maximumFractionDigits: 2,
    minimumFractionDigits: 2,
  });
  const percentFormatter = new Intl.NumberFormat("fr-FR", {
    maximumFractionDigits: 2,
    minimumFractionDigits: 2,
    style: "percent",
  });

  document.addEventListener("DOMContentLoaded", () => {
    if (window.Chart) {
      Chart.defaults.font.family = "'Inter', system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
      Chart.defaults.color = "#6b7280";
      Chart.defaults.borderColor = "#e5e7eb";
    }

    document.querySelectorAll("canvas[data-chart-endpoint]").forEach((canvas) => {
      renderCanvasChart(canvas);
    });
  });

  async function renderCanvasChart(canvas) {
    const endpoint = canvas.dataset.chartEndpoint;
    const chartKey = canvas.dataset.chartKey;
    if (!endpoint || !chartKey) {
      showEmpty(canvas, "Configuration graphique incomplète.");
      return;
    }

    try {
      const payload = await fetchChartEndpoint(endpoint);
      const chartData = payload[chartKey];
      if (!chartData || chartData.empty || !chartData.labels || chartData.labels.length === 0) {
        showEmpty(canvas);
        return;
      }
      if (window.Chart) {
        renderChart(canvas, chartData);
      } else {
        renderFallbackChart(canvas, chartData);
      }
    } catch (error) {
      showEmpty(canvas, "Impossible de charger les données du graphique.");
    }
  }

  async function fetchChartEndpoint(endpoint) {
    if (!endpointCache.has(endpoint)) {
      endpointCache.set(
        endpoint,
        fetch(endpoint, { headers: { Accept: "application/json" } }).then((response) => {
          if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
          }
          return response.json();
        }),
      );
    }
    return endpointCache.get(endpoint);
  }

  function renderChart(canvas, chartData) {
    const type = canvas.dataset.chartType || "bar";
    const axis = canvas.dataset.chartAxis || "x";
    const valueFormat = canvas.dataset.chartFormat || chartData.format || "number";
    const labels = chartData.labels;
    const values = chartData.values.map((value) => (value === null || value === "" ? null : Number(value)));
    const colors = labels.map((label, index) => chartColors[label] || fallbackPalette[index % fallbackPalette.length]);
    const isLine = type === "line";
    const isDoughnut = type === "doughnut";

    canvas.classList.remove("d-none");

    new Chart(canvas, {
      type,
      data: {
        labels,
        datasets: [
          {
            label: chartData.label || "",
            data: values,
            backgroundColor: isLine ? "rgba(15, 76, 129, 0.12)" : colors,
            borderColor: isLine ? "#0f4c81" : colors,
            borderWidth: isLine ? 2 : 1,
            borderRadius: isDoughnut || isLine ? 0 : 8,
            fill: isLine,
            pointBackgroundColor: "#0f4c81",
            pointBorderColor: "#ffffff",
            pointBorderWidth: 2,
            pointRadius: isLine ? 4 : 0,
            tension: isLine ? 0.32 : 0,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: axis,
        interaction: {
          intersect: false,
          mode: "nearest",
        },
        plugins: {
          legend: {
            display: isDoughnut,
            position: "bottom",
            labels: {
              boxWidth: 12,
              padding: 16,
              usePointStyle: true,
            },
          },
          tooltip: {
            callbacks: {
              label(context) {
                const label = context.dataset.label ? `${context.dataset.label}: ` : "";
                return `${label}${formatValue(context.parsedValue ?? context.raw, valueFormat)}`;
              },
            },
          },
        },
        scales: isDoughnut ? {} : buildScales(axis, valueFormat),
      },
    });
  }

  function buildScales(axis, valueFormat) {
    const valueAxis = axis === "y" ? "x" : "y";
    const categoryAxis = axis === "y" ? "y" : "x";
    return {
      [categoryAxis]: {
        grid: { display: false },
        ticks: {
          color: "#6b7280",
          autoSkip: axis !== "y",
          maxRotation: 0,
          callback(value) {
            const label = this.getLabelForValue(value);
            return truncateLabel(label, axis === "y" ? 28 : 18);
          },
        },
      },
      [valueAxis]: {
        beginAtZero: true,
        grid: { color: "rgba(229, 231, 235, 0.75)" },
        ticks: {
          color: "#6b7280",
          callback(value) {
            return formatValue(value, valueFormat, true);
          },
        },
      },
    };
  }

  function formatValue(rawValue, format, compact = false) {
    const value = typeof rawValue === "object" && rawValue !== null
      ? Number(rawValue.x ?? rawValue.y ?? 0)
      : Number(rawValue);

    if (!Number.isFinite(value)) {
      return "-";
    }
    if (format === "money") {
      if (compact && Math.abs(value) >= 1000000) {
        return `${moneyFormatter.format(value / 1000000)} M MAD`;
      }
      if (compact && Math.abs(value) >= 1000) {
        return `${moneyFormatter.format(value / 1000)} k MAD`;
      }
      return `${moneyFormatter.format(value)} MAD`;
    }
    if (format === "percent") {
      return percentFormatter.format(value);
    }
    return numberFormatter.format(value);
  }

  function truncateLabel(label, maxLength) {
    if (!label || label.length <= maxLength) {
      return label;
    }
    return `${label.slice(0, maxLength - 1)}…`;
  }

  function showEmpty(canvas, message) {
    canvas.classList.add("d-none");
    const panel = canvas.closest(".chart-panel");
    const empty = panel ? panel.querySelector(".chart-empty") : null;
    if (empty) {
      if (message) {
        empty.textContent = message;
      }
      empty.classList.remove("d-none");
    }
  }

  function showAllChartErrors(message) {
    document.querySelectorAll("canvas[data-chart-endpoint]").forEach((canvas) => showEmpty(canvas, message));
  }

  function renderFallbackChart(canvas, chartData) {
    const type = canvas.dataset.chartType || "bar";
    const valueFormat = canvas.dataset.chartFormat || chartData.format || "number";
    const labels = chartData.labels || [];
    const values = (chartData.values || []).map((value) => Number(value) || 0);
    const wrap = canvas.closest(".chart-canvas-wrap");
    if (!wrap) {
      return;
    }

    canvas.classList.add("d-none");
    const existing = wrap.querySelector(".fallback-chart");
    if (existing) {
      existing.remove();
    }

    const fallback = document.createElement("div");
    fallback.className = "fallback-chart";
    fallback.appendChild(
      type === "doughnut"
        ? buildFallbackDoughnut(labels, values, valueFormat)
        : buildFallbackBars(labels, values, valueFormat),
    );

    const note = document.createElement("div");
    note.className = "fallback-chart-note";
    note.textContent = "Rendu local affiché car Chart.js CDN n'est pas chargé.";
    fallback.appendChild(note);
    wrap.appendChild(fallback);
  }

  function buildFallbackDoughnut(labels, values, valueFormat) {
    const container = document.createElement("div");
    container.className = "fallback-doughnut-layout";
    const total = values.reduce((sum, value) => sum + Math.max(value, 0), 0);
    const circle = document.createElement("div");
    circle.className = "fallback-doughnut";

    let current = 0;
    const gradientParts = labels.map((label, index) => {
      const color = chartColors[label] || fallbackPalette[index % fallbackPalette.length];
      const portion = total ? (Math.max(values[index], 0) / total) * 100 : 0;
      const start = current;
      const end = current + portion;
      current = end;
      return `${color} ${start}% ${end}%`;
    });
    circle.style.background = `conic-gradient(${gradientParts.join(", ") || "#e5e7eb 0% 100%"})`;
    circle.innerHTML = `<span>${formatValue(total, valueFormat, true)}</span>`;
    container.appendChild(circle);

    const legend = document.createElement("div");
    legend.className = "fallback-legend";
    labels.forEach((label, index) => {
      const color = chartColors[label] || fallbackPalette[index % fallbackPalette.length];
      const item = document.createElement("div");
      item.className = "fallback-legend-item";
      item.innerHTML = `<span class="fallback-dot" style="background:${color}"></span><span>${escapeHtml(label)}</span><strong>${formatValue(values[index], valueFormat, true)}</strong>`;
      legend.appendChild(item);
    });
    container.appendChild(legend);
    return container;
  }

  function buildFallbackBars(labels, values, valueFormat) {
    const container = document.createElement("div");
    container.className = "fallback-bars";
    const max = Math.max(...values.map((value) => Math.abs(value)), 1);
    labels.forEach((label, index) => {
      const value = values[index] || 0;
      const color = chartColors[label] || fallbackPalette[index % fallbackPalette.length];
      const row = document.createElement("div");
      row.className = "fallback-bar-row";
      row.innerHTML = `
        <div class="fallback-bar-label" title="${escapeHtml(label)}">${escapeHtml(label)}</div>
        <div class="fallback-bar-track">
          <div class="fallback-bar-fill" style="width:${Math.max((Math.abs(value) / max) * 100, 2)}%; background:${color}"></div>
        </div>
        <div class="fallback-bar-value">${formatValue(value, valueFormat, true)}</div>
      `;
      container.appendChild(row);
    });
    return container;
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
})();
