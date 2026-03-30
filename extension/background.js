const SERVER = "http://127.0.0.1:7832";

async function getJobs() {
  const data = await chrome.storage.local.get("activeJobs");
  return data.activeJobs || [];
}

async function setJobs(jobs) {
  await chrome.storage.local.set({ activeJobs: jobs });
}

async function pollJobs() {
  const jobs = await getJobs();
  if (jobs.length === 0) { await chrome.alarms.clear("pollJob"); return; }

  const updated = await Promise.all(jobs.map(async (job) => {
    if (["done", "error", "cancelled"].includes(job.status)) return job;
    try {
      const res = await fetch(`${SERVER}/status/${job.jobId}`);
      const status = await res.json();
      return { ...job, ...status };
    } catch (e) { return job; }
  }));

  await setJobs(updated);

  const stillActive = updated.some(j => !["done", "error", "cancelled", "not_found"].includes(j.status));
  if (!stillActive) await chrome.alarms.clear("pollJob");
}

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "pollJob") pollJobs();
});

async function addJob(jobData) {
  const jobs = await getJobs();
  jobs.unshift(jobData);
  await setJobs(jobs);
  chrome.alarms.create("pollJob", { periodInMinutes: 0.05 }); // ~3 sec
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.type === "TRANSCRIBE") {
    fetch(`${SERVER}/transcribe`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ url: msg.url, title: msg.title || "", direct_url: msg.directUrl || "" }),
    })
      .then(r => r.json())
      .then(async (data) => {
        await addJob({ jobId: data.job_id, url: msg.url, title: msg.title, status: "pending", progress: 0, message: "ממתין..." });
        sendResponse({ ok: true });
      })
      .catch(e => sendResponse({ ok: false, error: e.message }));
    return true;
  }

  if (msg.type === "DOWNLOAD") {
    fetch(`${SERVER}/download`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ url: msg.url, title: msg.title || "", format: msg.format }),
    })
      .then(r => r.json())
      .then(async (data) => {
        await addJob({ jobId: data.job_id, url: msg.url, title: msg.title, status: "pending", progress: 0, message: "ממתין...", type: "download" });
        sendResponse({ ok: true });
      })
      .catch(e => sendResponse({ ok: false, error: e.message }));
    return true;
  }

  if (msg.type === "RECORD_ZOOM") {
    fetch(`${SERVER}/record/start`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title: "", keep_audio: msg.keepAudio || false, save_video: msg.saveVideo || false }),
    })
      .then(r => r.json())
      .then(async (data) => {
        if (data.error) return sendResponse({ ok: false, error: data.error });
        await addJob({ jobId: data.job_id, title: "הקלטת זום", status: "recording", progress: 5, message: "ממתין לזום...", type: "recording" });
        sendResponse({ ok: true });
      })
      .catch(e => sendResponse({ ok: false, error: e.message }));
    return true;
  }

  if (msg.type === "HEALTH") {
    fetch(`${SERVER}/health`).then(() => sendResponse({ ok: true })).catch(() => sendResponse({ ok: false }));
    return true;
  }

  if (msg.type === "CANCEL_JOB") {
    fetch(`${SERVER}/cancel/${msg.jobId}`, { method: "POST" })
      .then(async () => {
        const jobs = await getJobs();
        await setJobs(jobs.map(j =>
          j.jobId === msg.jobId ? { ...j, status: "cancelling", message: "מבטל..." } : j
        ));
        sendResponse({ ok: true });
      })
      .catch(e => sendResponse({ ok: false, error: e.message }));
    return true;
  }

  if (msg.type === "CLEAR_JOB") {
    getJobs().then(async (jobs) => {
      const updated = jobs.filter(j => j.jobId !== msg.jobId);
      await setJobs(updated);
      const stillActive = updated.some(j => !["done", "error", "cancelled"].includes(j.status));
      if (!stillActive) chrome.alarms.clear("pollJob");
      sendResponse({ ok: true });
    });
    return true;
  }
});
