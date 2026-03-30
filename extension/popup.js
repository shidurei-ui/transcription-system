const $ = (id) => document.getElementById(id);
let pageVideoInfo = null;
const SERVER = "http://127.0.0.1:7832";

// ── Direct polling while popup is open (every 1s) ─────────────────────────────
async function pollAndRender() {
  const data = await new Promise((r) => chrome.storage.local.get("activeJobs", r));
  const jobs = data.activeJobs || [];
  const active = jobs.filter(j => !["done", "error", "cancelled"].includes(j.status));
  if (active.length === 0) return;

  const updated = await Promise.all(jobs.map(async (job) => {
    if (["done", "error", "cancelled"].includes(job.status)) return job;
    try {
      const res = await fetch(`${SERVER}/status/${job.jobId}`);
      const status = await res.json();
      return { ...job, ...status };
    } catch (e) { return job; }
  }));

  await new Promise((r) => chrome.storage.local.set({ activeJobs: updated }, r));
  renderJobs(updated);
}

setInterval(pollAndRender, 1000);

// ── Server status ──────────────────────────────────────────────────────────────
function setChip(ok, msg) {
  const chip = $("server-chip");
  chip.className = ok ? "chip-ok" : "chip-err";
  chip.textContent = ok ? "✅ " + msg : "❌ " + msg;
}

// ── Render all jobs ────────────────────────────────────────────────────────────
function renderJobs(jobs) {
  const container = $("jobs-section");
  if (!jobs || jobs.length === 0) {
    container.innerHTML = "";
    return;
  }

  jobs.forEach((job) => {
    const existing = container.querySelector(`[data-job-id="${job.jobId}"]`);
    if (existing) {
      updateJobCard(existing, job);
    } else {
      const card = buildJobCard(job);
      container.appendChild(card);
    }
  });

  // Remove cards for jobs no longer in list
  container.querySelectorAll("[data-job-id]").forEach((el) => {
    if (!jobs.find(j => j.jobId === el.dataset.jobId)) el.remove();
  });
}

function buildJobCard(job) {
  const card = document.createElement("div");
  card.className = "job-card";
  card.dataset.jobId = job.jobId;
  const isRecording = job.type === "recording";
  const firstStep = isRecording
    ? `<div class="step" id="js-rec-${job.jobId}"><span class="step-icon">⏺️</span><span>הקלטה</span></div>`
    : `<div class="step" id="js-dl-${job.jobId}"><span class="step-icon">⬇️</span><span>הורדה</span></div>`;

  card.innerHTML = `
    <div class="job-card-header">
      <div class="job-title" id="jt-${job.jobId}"></div>
    </div>
    <div class="progress-bar-wrap"><div class="progress-bar-fill" id="jpb-${job.jobId}" style="width:0%"></div></div>
    <div class="progress-row">
      <span id="jmsg-${job.jobId}">ממתין...</span>
      <span id="jpct-${job.jobId}">0%</span>
    </div>
    <div class="steps">
      ${firstStep}
      <div class="step" id="js-cv-${job.jobId}"><span class="step-icon">🔄</span><span>המרה</span></div>
      <div class="step" id="js-tr-${job.jobId}"><span class="step-icon">🤖</span><span>תמלול</span></div>
      <div class="step" id="js-sv-${job.jobId}"><span class="step-icon">💾</span><span>שמירה</span></div>
    </div>
    <div id="jres-${job.jobId}" style="display:none"></div>
    <div class="job-actions" id="jact-${job.jobId}"></div>
  `;

  updateJobCard(card, job);
  return card;
}

function updateJobCard(card, job) {
  const id = job.jobId;
  const { status, message, progress = 0, title, output_file, chunk_current, chunk_total, type, course, lesson } = job;

  const label = type === "download" ? "⬇️" : "🎙️";
  const titleEl = card.querySelector(`#jt-${id}`);
  if (titleEl) titleEl.textContent = `${label} ${title || job.url || id}`;

  const pbEl = card.querySelector(`#jpb-${id}`);
  if (pbEl) pbEl.style.width = progress + "%";

  const pctEl = card.querySelector(`#jpct-${id}`);
  if (pctEl) pctEl.textContent = progress + "%";

  const msgEl = card.querySelector(`#jmsg-${id}`);
  if (msgEl) msgEl.textContent = chunk_total ? `${message} (${chunk_current}/${chunk_total})` : message;

  // Steps
  const isRecording = job.type === "recording";
  const STEPS = isRecording ? [
    { el: `js-rec-${id}`, active: "recording" },
    { el: `js-cv-${id}`, active: "converting" },
    { el: `js-tr-${id}`, active: "transcribing" },
    { el: `js-sv-${id}`, active: "saving" },
  ] : [
    { el: `js-dl-${id}`, active: "downloading" },
    { el: `js-cv-${id}`, active: "converting" },
    { el: `js-tr-${id}`, active: "transcribing" },
    { el: `js-sv-${id}`, active: "saving" },
  ];
  const ORDER = isRecording
    ? ["recording", "converting", "transcribing", "saving"]
    : ["downloading", "converting", "transcribing", "saving"];
  const currentIdx = ORDER.indexOf(status);

  STEPS.forEach(({ el }, i) => {
    const stepEl = card.querySelector(`#${el}`);
    if (!stepEl) return;
    stepEl.className = "step";
    stepEl.querySelector(".spinner")?.remove();
    if (status === "cancelled" || status === "cancelling") {
      stepEl.className = "step cancel";
    } else if (i < currentIdx || (status === "done" && i <= 3)) {
      stepEl.className = "step done";
    } else if (i === currentIdx) {
      stepEl.className = "step active";
      const s = document.createElement("span");
      s.className = "spinner";
      s.style.marginRight = "4px";
      stepEl.appendChild(s);
    }
  });

  // Result / error area
  const resEl = card.querySelector(`#jres-${id}`);
  if (resEl) {
    resEl.className = "";
    resEl.style.display = "none";
    resEl.textContent = "";

    if (status === "done") {
      resEl.style.display = "block";
      resEl.className = "job-result";
      const loc = course && lesson ? `📁 ${course} ← ${lesson}` : (output_file || "");
      resEl.textContent = `✅ הושלם!\n${loc}`;
    } else if (status === "error") {
      resEl.style.display = "block";
      resEl.className = "job-error";
      resEl.textContent = "❌ " + message;
    } else if (status === "cancelled") {
      resEl.style.display = "block";
      resEl.className = "job-cancelled";
      resEl.textContent = "⛔ בוטל";
    }
  }

  // Action buttons
  const actEl = card.querySelector(`#jact-${id}`);
  if (actEl) {
    actEl.innerHTML = "";
    const isActive = !["done", "error", "cancelled", "cancelling"].includes(status);
    if (isActive) {
      const cancelBtn = document.createElement("button");
      cancelBtn.className = "btn-cancel";
      cancelBtn.textContent = "⛔ בטל";
      cancelBtn.onclick = () => {
        cancelBtn.disabled = true;
        cancelBtn.textContent = "מבטל...";
        chrome.runtime.sendMessage({ type: "CANCEL_JOB", jobId: id });
      };
      actEl.appendChild(cancelBtn);
    } else {
      const clearBtn = document.createElement("button");
      clearBtn.className = "btn-clear";
      clearBtn.textContent = "✕ הסר";
      clearBtn.onclick = () => {
        chrome.runtime.sendMessage({ type: "CLEAR_JOB", jobId: id });
        card.remove();
      };
      actEl.appendChild(clearBtn);
    }
  }
}

// ── Video info from page ───────────────────────────────────────────────────────
async function getVideoInfo() {
  return new Promise((resolve) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      const tab = tabs[0];
      if (!tab || !tab.url) return resolve(null);
      const url = tab.url;
      if (url.startsWith("chrome://") || url.startsWith("chrome-extension://") ||
          url.startsWith("about:") || url.startsWith("edge://")) return resolve(null);
      chrome.tabs.sendMessage(tab.id, { type: "GET_VIDEO_INFO" }, (res) => {
        if (chrome.runtime.lastError || !res) {
          chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ["content.js"] }, () => {
            if (chrome.runtime.lastError) return resolve(null);
            chrome.tabs.sendMessage(tab.id, { type: "GET_VIDEO_INFO" }, (r) => resolve(r || null));
          });
        } else resolve(res);
      });
    });
  });
}

// ── Action buttons ─────────────────────────────────────────────────────────────
function startAction(type, format) {
  if (!pageVideoInfo) return;
  const btn = type === "TRANSCRIBE" ? $("btn-transcribe") : (format === "mp3" ? $("btn-dl-mp3") : $("btn-dl-mp4"));
  if (btn) { btn.disabled = true; setTimeout(() => { btn.disabled = false; }, 3000); }

  chrome.runtime.sendMessage(
    { type, url: pageVideoInfo.pageUrl, title: pageVideoInfo.title, directUrl: pageVideoInfo.directUrls[0] || "", format },
    (res) => {
      if (!res?.ok) alert("שגיאה: " + (res?.error || "לא ניתן לשלוח לשרת"));
    }
  );
}

$("btn-restart").addEventListener("click", async () => {
  const btn = $("btn-restart");
  btn.textContent = "...";
  btn.disabled = true;
  try {
    await fetch(`${SERVER}/restart`, { method: "POST" });
    setChip(false, "מאתחל...");
    // wait for server to come back
    setTimeout(async () => {
      for (let i = 0; i < 15; i++) {
        await new Promise(r => setTimeout(r, 1000));
        try {
          await fetch(`${SERVER}/health`);
          setChip(true, "שרת פעיל");
          btn.textContent = "↺";
          btn.disabled = false;
          return;
        } catch {}
      }
      setChip(false, "שגיאה");
      btn.textContent = "↺";
      btn.disabled = false;
    }, 800);
  } catch {
    btn.textContent = "↺";
    btn.disabled = false;
  }
});

$("btn-transcribe").addEventListener("click", () => startAction("TRANSCRIBE"));
$("btn-dl-mp3").addEventListener("click",    () => startAction("DOWNLOAD", "mp3"));
$("btn-dl-mp4").addEventListener("click",    () => startAction("DOWNLOAD", "mp4"));

// ── Zoom recording ──────────────────────────────────────────────────────────
$("btn-record").addEventListener("click", () => {
  const btn = $("btn-record");
  btn.disabled = true;
  setTimeout(() => { btn.disabled = false; }, 3000);
  const keepAudio = $("chk-keep-audio")?.checked || false;
  const saveVideo = $("chk-save-video")?.checked || false;
  chrome.runtime.sendMessage({ type: "RECORD_ZOOM", keepAudio, saveVideo }, (res) => {
    if (!res?.ok) alert("שגיאה: " + (res?.error || "לא ניתן להתחיל הקלטה"));
  });
});

// ── Init ───────────────────────────────────────────────────────────────────────
(async () => {
  // 1. Check server
  chrome.runtime.sendMessage({ type: "HEALTH" }, (res) => {
    if (res?.ok) setChip(true, "שרת פעיל");
    else setChip(false, "שרת לא פועל");
  });

  // 2. Load jobs from storage and render
  const data = await new Promise((r) => chrome.storage.local.get("activeJobs", r));
  renderJobs(data.activeJobs || []);

  // 3. Listen for storage changes
  chrome.storage.onChanged.addListener((changes) => {
    if (changes.activeJobs) renderJobs(changes.activeJobs.newValue || []);
  });

  // 4. Get current page video info
  const info = await getVideoInfo();
  pageVideoInfo = info;
  if (info) {
    $("page-url").textContent = info.pageUrl;
    if (info.hasVideo || info.directUrls.length > 0) {
      $("detected-label").textContent = "✅ סרטון זוהה בדף";
      $("detected-label").style.color = "#4ade80";
    } else {
      $("detected-label").textContent = "⚠️ לא זוהה — yt-dlp ינסה בכל זאת";
      $("detected-label").style.color = "#fbbf24";
    }
    $("btn-transcribe").disabled = false;
    $("btn-dl-mp3").disabled = false;
    $("btn-dl-mp4").disabled = false;
  } else {
    $("detected-label").textContent = "❌ לא ניתן לגשת לדף";
  }
})();
