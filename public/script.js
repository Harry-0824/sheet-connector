let statusChart = null;
let trendChart = null;

document.addEventListener("DOMContentLoaded", () => {
  fetchData();

  document.getElementById("refreshBtn").addEventListener("click", fetchData);
  document
    .getElementById("executeBtn")
    .addEventListener("click", handleExecute);
});

async function fetchData() {
  showStatus("正在從試算表抓取資料...", "info");
  try {
    const response = await fetch("/api/preview");
    const result = await response.json();

    if (result.error) {
      showStatus(`錯誤: ${result.error}`, "error");
      return;
    }

    renderTable(result.headers, result.data);
    updateStats(result.headers, result.data);
    showStatus("資料同步完成", "success");
  } catch (err) {
    showStatus("連線伺服器失敗", "error");
  }
}

function showStatus(msg, type) {
  const statusEl = document.getElementById("status");
  statusEl.textContent = msg;
  statusEl.className = "status-msg " + type;
  if (type === "success") {
    setTimeout(() => (statusEl.textContent = ""), 3000);
  }
}

function renderTable(headers, data) {
  const head = document.getElementById("tableHead");
  const body = document.getElementById("tableBody");
  head.innerHTML = "";
  body.innerHTML = "";

  if (!headers || headers.length === 0) {
    body.innerHTML = '<tr><td colspan="100">未找到資料</td></tr>';
    return;
  }

  // Headers
  headers.forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    head.appendChild(th);
  });

  // Body
  data.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((h) => {
      const td = document.createElement("td");
      td.textContent = row[h] || "";
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
}

function updateStats(headers, data) {
  const total = data.length;
  let replied = 0;
  const dateCounts = {};

  // 尋找日期欄位
  const dateKey = headers.find((h) => /日期|Date|Time|時間/.test(h));

  data.forEach((item) => {
    if (item["是否自動回覆"] && item["是否自動回覆"].trim() !== "") {
      replied++;
    }

    if (dateKey && item[dateKey]) {
      let dStr = item[dateKey];
      let dateObj = new Date(dStr);
      let key = "";
      if (!isNaN(dateObj)) {
        key = dateObj.toISOString().split("T")[0];
      } else {
        key = dStr.split(" ")[0];
      }
      dateCounts[key] = (dateCounts[key] || 0) + 1;
    }
  });

  const pending = total - replied;

  animateNumber("totalUsers", total);
  animateNumber("pendingUsers", pending);
  animateNumber("repliedUsers", replied);

  renderCharts(replied, pending, dateCounts);
}

function animateNumber(id, endValue) {
  const el = document.getElementById(id);
  let startValue = 0;
  const duration = 1000;
  const startTimestamp = performance.now();

  function step(timestamp) {
    const progress = Math.min((timestamp - startTimestamp) / duration, 1);
    el.textContent = Math.floor(progress * endValue);
    if (progress < 1) {
      window.requestAnimationFrame(step);
    }
  }
  window.requestAnimationFrame(step);
}

function renderCharts(replied, pending, dateCounts) {
  // 樣式變數
  const successColor = "#00ff9d";
  const warningColor = "#ffb800";
  const primaryColor = "#00f3ff";
  const textColor = "#888";

  // Status Chart
  if (statusChart) statusChart.destroy();
  statusChart = new Chart(document.getElementById("statusChart"), {
    type: "doughnut",
    data: {
      labels: ["已回覆", "待處理"],
      datasets: [
        {
          data: [replied, pending],
          backgroundColor: [successColor, warningColor],
          borderWidth: 0,
          hoverOffset: 10,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "bottom", labels: { color: textColor } },
      },
    },
  });

  // Trend Chart
  if (trendChart) trendChart.destroy();
  const sortedDates = Object.keys(dateCounts).sort();
  const trendData = sortedDates.map((d) => dateCounts[d]);

  trendChart = new Chart(document.getElementById("trendChart"), {
    type: "bar",
    data: {
      labels: sortedDates.length > 0 ? sortedDates : ["No Data"],
      datasets: [
        {
          label: "新增人數",
          data: trendData.length > 0 ? trendData : [0],
          backgroundColor: primaryColor,
          borderRadius: 4,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { color: textColor }, grid: { display: false } },
        y: {
          ticks: { color: textColor, stepSize: 1 },
          grid: { color: "rgba(255,255,255,0.05)" },
        },
      },
      plugins: {
        legend: { display: false },
      },
    },
  });
}

async function handleExecute() {
  if (!confirm("您確定要發送 Email 給所有待處理的使用者嗎？")) return;

  const btn = document.getElementById("executeBtn");
  const originalText = btn.textContent;

  btn.disabled = true;
  btn.textContent = "處理中...";
  showStatus("執行自動任務中，請稍候...", "info");

  try {
    const response = await fetch("/api/execute", { method: "POST" });
    const result = await response.json();

    if (result.error) {
      showStatus(`錯誤: ${result.error}`, "error");
    } else {
      showStatus(
        `成功: ${result.message}。已處理 ${result.processed} 封郵件。`,
        "success",
      );
      fetchData(); // 重新整理資料
    }
  } catch (err) {
    showStatus("執行任務失敗，請檢查後端連線", "error");
  } finally {
    btn.disabled = false;
    btn.textContent = originalText;
  }
}
