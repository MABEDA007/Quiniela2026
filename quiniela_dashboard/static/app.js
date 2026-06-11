let state = null;
let selectedParticipant = null;
let activeFilter = "all";
let config = {};

const formatter = new Intl.DateTimeFormat("es-MX", {
  dateStyle: "medium",
  timeStyle: "short",
});

const $ = (selector) => document.querySelector(selector);

function setLoading(isLoading) {
  const refreshBtn = $("#refreshBtn");
  const uploadBtn = $("#uploadBtn");
  if (refreshBtn) refreshBtn.disabled = isLoading;
  if (uploadBtn) uploadBtn.disabled = isLoading;
}

async function loadState(refresh = false) {
  setLoading(true);
  try {
    const [configResponse, stateResponse] = await Promise.all([
      fetch("/api/config"),
      fetch(`/api/state${refresh ? "?refresh=1" : ""}`),
    ]);
    config = await configResponse.json();
    const response = stateResponse;
    state = await response.json();
    if (!selectedParticipant && state.participants.length) {
      selectedParticipant = state.participants[0].name;
    }
    render();
  } catch (error) {
    console.error("No pude cargar el dashboard:", error);
  } finally {
    setLoading(false);
  }
}

function render() {
  renderStatus();
  renderAdmin();
  renderMetrics();
  renderStandings();
  renderParticipantDetail();
  renderMatches();
}

function renderStatus() {
  return;
}

function renderAdmin() {
  const panel = $("#adminPanel");
  panel.classList.toggle("hidden", !config.uploadEnabled);
}

function renderMetrics() {
  const exact = state.participants.reduce((sum, item) => sum + item.stats.exact, 0);
  const summary = matchSummary();
  $("#metricParticipants").textContent = state.participants.length;
  $("#metricPlayed").textContent = summary.played;
  $("#metricPending").textContent = summary.pending;
  $("#metricExact").textContent = exact;
}

function uniqueMatchCount() {
  const keys = new Set(state.predictions.map((item) => `${item.homeKey}|${item.awayKey}|${item.date}`));
  return keys.size;
}

function renderStandings() {
  const body = $("#standingsBody");
  body.innerHTML = "";

  if (!state.participants.length) {
    body.innerHTML = `<tr><td colspan="6" class="empty-state">No encontre archivos .xlsx de quiniela en la carpeta.</td></tr>`;
    return;
  }

  for (const participant of state.participants) {
    const row = document.createElement("tr");
    row.className = participant.name === selectedParticipant ? "selected" : "";
    row.innerHTML = `
      <td><span class="rank">${participant.rank}</span></td>
      <td>
        <span class="participant-name">${participant.name}</span>
        <span class="file-name">${participant.file}</span>
      </td>
      <td class="numeric"><strong>${participant.stats.points}</strong></td>
      <td class="numeric">${participant.stats.exact}</td>
      <td class="numeric">${participant.stats.trend}</td>
      <td class="numeric">${participant.stats.miss}</td>
    `;
    row.addEventListener("click", () => {
      selectedParticipant = participant.name;
      renderStandings();
      renderParticipantDetail();
    });
    body.appendChild(row);
  }
}

function renderParticipantDetail() {
  const participant = state.participants.find((item) => item.name === selectedParticipant);
  const container = $("#participantDetail");

  if (!participant) {
    $("#detailTitle").textContent = "Selecciona un participante";
    container.className = "detail-empty";
    container.textContent = "Haz clic en una fila de la tabla para ver sus aciertos partido por partido.";
    return;
  }

  $("#detailTitle").textContent = participant.name;
  container.className = "detail-list";
  container.innerHTML = "";

  for (const prediction of participant.predictions) {
    const actual = actualScoreText(prediction.fixture);
    const row = document.createElement("div");
    row.className = "prediction-row";
    row.innerHTML = `
      <div>
        <strong>${prediction.home} ${prediction.predHome}-${prediction.predAway} ${prediction.away}</strong>
        <small>${actual ? actual : ""}</small>
      </div>
      <span class="result-symbol ${prediction.score.status}" title="${prediction.score.label}">${resultSymbol(prediction.score.status)}</span>
    `;
    container.appendChild(row);
  }
}

function actualScoreText(fixture) {
  if (!fixture || fixture.homeGoals === null || fixture.awayGoals === null || fixture.homeGoals === undefined || fixture.awayGoals === undefined) {
    return "";
  }
  return `${fixture.home} ${fixture.homeGoals}-${fixture.awayGoals} ${fixture.away}`;
}

function renderMatches() {
  const grid = $("#matchesGrid");
  grid.innerHTML = "";

  const fixtures = state.fixtures.length ? state.fixtures : fallbackFixturesFromPredictions();
  const filtered = fixtures.filter((fixture) => {
    const status = fixture.statusShort;
    const hasScore = fixture.homeGoals !== null && fixture.homeGoals !== undefined && fixture.awayGoals !== null && fixture.awayGoals !== undefined;
    if (activeFilter === "played") return fixtureIsPlayed(fixture);
    if (activeFilter === "pending") return !fixtureIsPlayed(fixture);
    if (activeFilter === "unmatched") return unmatchedPredictionCount(fixture) > 0;
    return true;
  });

  if (!filtered.length) {
    grid.innerHTML = `<div class="empty-state">No hay partidos para este filtro.</div>`;
    return;
  }

  const template = $("#matchCardTemplate");
  for (const fixture of filtered) {
    const card = template.content.cloneNode(true);
    card.querySelector(".round").textContent = fixture.round || fixture.localDate || "Quiniela";
    card.querySelector(".status").textContent = fixture.statusShort || "Pend.";
    card.querySelector(".home span").textContent = fixture.home;
    card.querySelector(".away span").textContent = fixture.away;
    card.querySelector(".home img").src = fixture.homeLogo || "";
    card.querySelector(".away img").src = fixture.awayLogo || "";
    card.querySelector(".score").textContent = scoreText(fixture);
    card.querySelector(".match-sub").textContent = matchSubText(fixture);
    grid.appendChild(card);
  }
}

function fallbackFixturesFromPredictions() {
  const map = new Map();
  for (const prediction of state.predictions) {
    const key = `${prediction.homeKey}|${prediction.awayKey}|${prediction.date}`;
    if (!map.has(key)) {
      map.set(key, {
        id: key,
        date: prediction.date,
        time: prediction.time,
        localDate: prediction.date,
        round: prediction.sheet,
        statusShort: "Sin API",
        home: prediction.home,
        away: prediction.away,
        homeKey: prediction.homeKey,
        awayKey: prediction.awayKey,
        homeGoals: null,
        awayGoals: null,
        venue: prediction.venue,
      });
    }
  }
  return [...map.values()];
}

function unmatchedPredictionCount(fixture) {
  return state.predictions.filter((prediction) => {
    const sameTeams = prediction.homeKey === fixture.homeKey && prediction.awayKey === fixture.awayKey;
    return sameTeams && prediction.score.status === "unmatched";
  }).length;
}

function allFixtures() {
  return state.fixtures.length ? state.fixtures : fallbackFixturesFromPredictions();
}

function fixtureIsPlayed(fixture) {
  const status = fixture.statusShort;
  const hasScore = fixture.homeGoals !== null && fixture.homeGoals !== undefined && fixture.awayGoals !== null && fixture.awayGoals !== undefined;
  return hasScore && !["NS", "TBD"].includes(status);
}

function matchSummary() {
  const fixtures = allFixtures();
  const played = fixtures.filter(fixtureIsPlayed).length;
  return {
    played,
    pending: Math.max(fixtures.length - played, 0),
  };
}

function resultSymbol(status) {
  if (status === "exact") return "🌟";
  if (status === "trend") return "✅";
  if (status === "miss") return "❌";
  return "—";
}

function scoreText(fixture) {
  if (fixture.homeGoals === null || fixture.homeGoals === undefined || fixture.awayGoals === null || fixture.awayGoals === undefined) {
    return "vs";
  }
  return `${fixture.homeGoals} - ${fixture.awayGoals}`;
}

function matchSubText(fixture) {
  const parts = [];
  const dateText = formatFixtureDate(fixture);
  if (dateText) parts.push(dateText);
  if (fixture.venue) parts.push(fixture.city ? `${fixture.venue}, ${fixture.city}` : fixture.venue);
  return parts.join(" - ");
}

function formatFixtureDate(fixture) {
  const raw = fixture.date || fixture.localDate;
  if (!raw) return "";

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    const [year, month, day] = raw.split("-").map(Number);
    const [hour = 0, minute = 0] = String(fixture.time || "").split(":").map(Number);
    const date = new Date(year, month - 1, day, hour, minute);
    return formatter.format(date);
  }

  const date = new Date(raw);
  return Number.isNaN(date.getTime()) ? "" : formatter.format(date);
}

function formatGeneratedAt(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return formatter.format(date);
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      resolve(result.includes(",") ? result.split(",")[1] : result);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function uploadExcel(event) {
  event.preventDefault();
  const file = $("#excelFile").files[0];
  const token = $("#adminToken").value.trim();
  const message = $("#uploadMessage");

  if (!token || !file) {
    message.textContent = "Indica token y archivo .xlsx.";
    message.className = "upload-message error";
    return;
  }

  setLoading(true);
  message.textContent = "Subiendo...";
  message.className = "upload-message";
  try {
    const contentBase64 = await fileToBase64(file);
    const response = await fetch("/api/upload", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ token, filename: file.name, contentBase64 }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "No pude subir el archivo.");
    }
    message.textContent = `Listo: ${payload.filename}`;
    message.className = "upload-message ok";
    $("#excelFile").value = "";
    selectedParticipant = null;
    await loadState(false);
  } catch (error) {
    message.textContent = error.message;
    message.className = "upload-message error";
  } finally {
    setLoading(false);
  }
}

document.querySelectorAll(".filter").forEach((button) => {
  button.addEventListener("click", () => {
    activeFilter = button.dataset.filter;
    document.querySelectorAll(".filter").forEach((item) => item.classList.toggle("active", item === button));
    renderMatches();
  });
});

$("#refreshBtn").addEventListener("click", () => loadState(true));
$("#uploadForm").addEventListener("submit", uploadExcel);

loadState();
