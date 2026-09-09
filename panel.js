const STORAGE_KEY = "teamsTranscript";
// Marks live in their own key so the content script's per-second rewrite of the
// meeting bucket can never race with them.
const MARKS_KEY = "teamsTranscriptMarks";

const badgeEl = document.getElementById("badge");
const statusEl = document.getElementById("status");
const feedEl = document.getElementById("feed");
const searchEl = document.getElementById("search");
const autoscrollEl = document.getElementById("autoscroll");
const onlyActionsEl = document.getElementById("only-actions");
const chooseVaultEl = document.getElementById("choose-vault");
const vaultStatusEl = document.getElementById("vault-status");
const meetingEl = document.getElementById("meeting");
const meetingSelectEl = document.getElementById("meeting-select");
const peopleEl = document.getElementById("people");
const peopleSummaryEl = document.getElementById("people-summary");
const peopleListEl = document.getElementById("people-list");

const SOURCE_LABELS = {
  roster: "список участников",
  tile: "плитка видео",
  caption: "говорил",
};

function pruneEmptyMeetings(raw) {
  const meetings = {};
  for (const [id, meeting] of Object.entries(raw.meetings || {})) {
    if (meeting?.captions?.length) meetings[id] = meeting;
  }
  return {
    ...raw,
    meetings,
    activeMeetingId: meetings[raw.activeMeetingId] ? raw.activeMeetingId : null,
  };
}

const emptyStore = () => ({
  version: 2,
  activeMeetingId: null,
  meetings: {},
  removedIds: {},
});

let store = emptyStore();
let marks = {};
let selectedId = null;
let pinnedToSelection = false;
let vaultHandle = null;
let vaultFileNames = {};
let syncTimer = null;
let syncing = false;
let syncAgain = false;

function openVaultDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("teamsTranscriptVault", 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore("settings");
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getVaultSetting(key) {
  const db = await openVaultDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction("settings", "readonly");
    const request = transaction.objectStore("settings").get(key);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    transaction.oncomplete = () => db.close();
  });
}

async function setVaultSetting(key, value) {
  const db = await openVaultDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction("settings", "readwrite");
    transaction.objectStore("settings").put(value, key);
    transaction.oncomplete = () => {
      db.close();
      resolve();
    };
    transaction.onerror = () => reject(transaction.error);
  });
}

function setVaultStatus(text, state = "") {
  vaultStatusEl.textContent = text;
  vaultStatusEl.dataset.state = state;
}

function captionKey(item) {
  return item.key || `${item.startedAt}:${item.speaker}`;
}

function marksForSelected() {
  const meeting = selectedMeeting();
  return (meeting && marks[meeting.id]) || {};
}

function isMeetingAction(meeting, item) {
  return Boolean(marks[meeting.id]?.[captionKey(item)]);
}

function isActionItem(item) {
  return Boolean(marksForSelected()[captionKey(item)]);
}

async function toggleActionItem(item) {
  const meeting = selectedMeeting();
  if (!meeting) return;

  const key = captionKey(item);
  const forMeeting = { ...(marks[meeting.id] || {}) };

  if (forMeeting[key]) {
    delete forMeeting[key];
  } else {
    forMeeting[key] = { type: "action", createdAt: new Date().toISOString() };
  }

  marks = { ...marks, [meeting.id]: forMeeting };
  await chrome.storage.local.set({ [MARKS_KEY]: marks });
  render();
}

function timeOf(iso) {
  if (!iso) return "";
  return new Date(iso).toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

function meetingsNewestFirst() {
  return Object.values(store.meetings || {}).sort((a, b) =>
    String(b.lastUpdatedAt || b.startedAt).localeCompare(
      String(a.lastUpdatedAt || a.startedAt)
    )
  );
}

function selectedMeeting() {
  const meetings = store.meetings || {};
  if (pinnedToSelection && selectedId && meetings[selectedId]) {
    return meetings[selectedId];
  }
  if (store.activeMeetingId && meetings[store.activeMeetingId]) {
    return meetings[store.activeMeetingId];
  }
  if (selectedId && meetings[selectedId]) return meetings[selectedId];
  return meetingsNewestFirst()[0] || null;
}

function meetingLabel(meeting) {
  const title = meeting.title?.trim() || "без названия";
  const when = meeting.startedAt
    ? new Date(meeting.startedAt).toLocaleString([], {
        day: "2-digit",
        month: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
      })
    : "";
  return `${title} · ${when} · ${meeting.captions?.length || 0}`;
}

const MERGE_GAP_MS = 15000;

function mergeTurns(captions, isActionFn) {
  const turns = [];

  for (const item of captions || []) {
    const text = String(item.text || "").replace(/\s+/g, " ").trim();
    if (!text) continue;

    const last = turns[turns.length - 1];
    // Teams captions overlap: a caption keeps updating after the next one
    // appears, so the gap is often negative. Overlap means the same turn.
    const rawGap = last
      ? new Date(item.startedAt).getTime() - new Date(last.endedAt).getTime()
      : Number.POSITIVE_INFINITY;
    const gap = Number.isFinite(rawGap) ? Math.max(rawGap, 0) : 0;

    if (last && last.speaker === item.speaker && gap <= MERGE_GAP_MS) {
      last.text = `${last.text} ${text}`.replace(/\s+/g, " ").trim();
      const end = item.updatedAt || item.startedAt;
      if (String(end) > String(last.endedAt)) last.endedAt = end;
      last.actionItem = last.actionItem || Boolean(isActionFn(item));
      last.parts += 1;
      continue;
    }

    turns.push({
      speaker: item.speaker,
      text,
      startedAt: item.startedAt,
      endedAt: item.updatedAt || item.startedAt,
      actionItem: Boolean(isActionFn(item)),
      parts: 1,
    });
  }

  return turns;
}

function visibleCaptions() {
  const query = searchEl.value.trim().toLowerCase();
  let captions = selectedMeeting()?.captions || [];

  if (onlyActionsEl.checked) {
    captions = captions.filter(isActionItem);
  }
  if (query) {
    captions = captions.filter(
      (item) =>
        item.text.toLowerCase().includes(query) ||
        item.speaker.toLowerCase().includes(query)
    );
  }

  return captions;
}

function formatTxt(captions) {
  const meeting = selectedMeeting() || {};
  const people = (meeting.participants || []).map((item) => item.name);
  const turns = mergeTurns(captions, isActionItem);
  const actions = turns.filter((turn) => turn.actionItem);
  const first = captions[0]?.startedAt;
  const last = captions[captions.length - 1]?.updatedAt;

  const header = [
    `Встреча: ${meeting.title?.trim() || "без названия"}`,
    meeting.url ? `Ссылка: ${meeting.url}` : null,
    first ? `Период: ${timeOf(first)} – ${timeOf(last)}` : null,
    people.length ? `Участники (${people.length}): ${people.join(", ")}` : null,
    "",
  ].filter((line) => line !== null);

  const actionBlock = actions.length
    ? [
        `Экшен-айтемы (${actions.length}):`,
        ...actions.map(
          (turn) => `- [${timeOf(turn.startedAt)}] ${turn.speaker}: ${turn.text}`
        ),
        "",
      ]
    : [];

  const body = turns.map((turn) => {
    const prefix = turn.actionItem ? "[ACTION] " : "";
    return `${prefix}[${timeOf(turn.startedAt)}] ${turn.speaker}: ${turn.text}`;
  });

  return [...header, ...actionBlock, ...body].join("\n");
}

function exportPayload(meeting) {
  const captions = (meeting.captions || []).map((item) => ({
    ...item,
    actionItem: isMeetingAction(meeting, item),
  }));
  const turns = mergeTurns(meeting.captions, (item) =>
    isMeetingAction(meeting, item)
  );

  return {
    ...meeting,
    actionItemCount: turns.filter((turn) => turn.actionItem).length,
    captions,
    turns,
  };
}

function yamlString(value) {
  return JSON.stringify(String(value || ""));
}

function markdownForMeeting(meeting) {
  const participants = meeting.participants || [];
  const turns = mergeTurns(meeting.captions, (item) =>
    isMeetingAction(meeting, item)
  );
  const actions = turns.filter((turn) => turn.actionItem);
  const title = meeting.title?.trim() || "Без названия";
  const lines = [
    "---",
    `title: ${yamlString(title)}`,
    `date: ${yamlString(meeting.startedAt || "")}`,
    `teams_url: ${yamlString(meeting.url || "")}`,
    `updated: ${yamlString(meeting.lastUpdatedAt || "")}`,
    "participants:",
    ...participants.map((person) => `  - ${yamlString(person.name)}`),
    "---",
    "",
    `# ${title}`,
    "",
    "## Экшен-айтемы",
    "",
    ...(actions.length
      ? actions.map((turn) => `- [ ] **${turn.speaker}:** ${turn.text}`)
      : ["_Нет отмеченных экшен-айтемов._"]),
    "",
    "## Транскрипт",
    "",
    ...turns.map((turn) => {
      const action = turn.actionItem ? " `#action-item`" : "";
      return `- \`${timeOf(turn.startedAt)}\` **${turn.speaker}:** ${turn.text}${action}`;
    }),
    "",
  ];

  return lines.join("\n");
}

function safeFilePart(value) {
  return String(value || "")
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
}

function shortMeetingId(id) {
  let hash = 2166136261;
  for (const char of String(id || "")) {
    hash ^= char.charCodeAt(0);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(36).slice(0, 6);
}

function meetingFileName(meeting) {
  if (vaultFileNames[meeting.id]) return vaultFileNames[meeting.id];

  const date = new Date(meeting.startedAt || Date.now());
  const safeDate = Number.isNaN(date.getTime()) ? new Date() : date;
  const stamp = [
    safeDate.getFullYear(),
    String(safeDate.getMonth() + 1).padStart(2, "0"),
    String(safeDate.getDate()).padStart(2, "0"),
  ].join("-");
  const time = [
    String(safeDate.getHours()).padStart(2, "0"),
    String(safeDate.getMinutes()).padStart(2, "0"),
  ].join("-");
  const title = safeFilePart(meeting.title) || "без-названия";

  vaultFileNames[meeting.id] =
    `${title} ${stamp} ${time} ${shortMeetingId(meeting.id)}.md`;
  return vaultFileNames[meeting.id];
}

async function syncMeetingToVault(meeting) {
  const filename = meetingFileName(meeting);
  const fileHandle = await vaultHandle.getFileHandle(filename, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(markdownForMeeting(meeting));
  await writable.close();
}

async function syncVault() {
  if (!vaultHandle) return;
  if (syncing) {
    syncAgain = true;
    return;
  }

  const permission = await vaultHandle.queryPermission({ mode: "readwrite" });
  if (permission !== "granted") {
    setVaultStatus("Нужно заново разрешить доступ", "error");
    return;
  }

  syncing = true;
  try {
    const meetings = meetingsNewestFirst().filter(
      (meeting) => meeting.captions?.length
    );
    for (const meeting of meetings) {
      await syncMeetingToVault(meeting);
    }
    await setVaultSetting("fileNames", vaultFileNames);
    setVaultStatus(
      meetings.length
        ? `${vaultHandle.name} · синхронизировано`
        : `${vaultHandle.name} · ожидаю реплики`,
      "ok"
    );
  } catch (error) {
    console.error("[Teams Transcript] Ошибка записи в Obsidian:", error);
    setVaultStatus(`Ошибка записи: ${error.message}`, "error");
  } finally {
    syncing = false;
    if (syncAgain) {
      syncAgain = false;
      scheduleVaultSync();
    }
  }
}

function scheduleVaultSync() {
  if (!vaultHandle) return;
  clearTimeout(syncTimer);
  syncTimer = setTimeout(syncVault, 500);
}

async function restoreVault() {
  try {
    vaultHandle = await getVaultSetting("directory");
    vaultFileNames = (await getVaultSetting("fileNames")) || {};
    if (!vaultHandle) return;

    chooseVaultEl.textContent = "Сменить папку Obsidian";
    const permission = await vaultHandle.queryPermission({ mode: "readwrite" });
    if (permission === "granted") {
      setVaultStatus(`${vaultHandle.name} · подключено`, "ok");
      scheduleVaultSync();
    } else {
      setVaultStatus(`${vaultHandle.name} · нажми для доступа`, "error");
    }
  } catch (error) {
    console.error("[Teams Transcript] Не удалось восстановить папку:", error);
    setVaultStatus("Не удалось открыть сохранённую папку", "error");
  }
}

function renderMeetingSelect() {
  const meetings = meetingsNewestFirst();
  const current = selectedMeeting();

  meetingSelectEl.hidden = meetings.length < 1;
  meetingSelectEl.replaceChildren();

  for (const meeting of meetings) {
    const option = document.createElement("option");
    option.value = meeting.id;
    option.textContent = meetingLabel(meeting);
    option.selected = meeting.id === current?.id;
    meetingSelectEl.append(option);
  }
}

function renderPeople() {
  const people = selectedMeeting()?.participants || [];
  peopleEl.hidden = people.length === 0;
  if (!people.length) return;

  peopleSummaryEl.textContent = `Участники: ${people.length}`;
  peopleListEl.replaceChildren();

  for (const person of people) {
    const li = document.createElement("li");
    li.textContent = person.name;

    const labels = (person.sources || [])
      .map((source) => SOURCE_LABELS[source] || source)
      .join(", ");
    if (labels) {
      const source = document.createElement("span");
      source.className = "source";
      source.textContent = ` — ${labels}`;
      li.append(source);
    }

    peopleListEl.append(li);
  }
}

function renderHeader() {
  const meeting = selectedMeeting();
  const total = meeting?.captions?.length || 0;

  const actions = (meeting?.captions || []).filter(isActionItem).length;

  badgeEl.textContent = total ? String(total) : "—";
  meetingEl.textContent = meeting?.title?.trim() || "";

  const isActive = meeting && meeting.id === store.activeMeetingId;

  if (isActive && meeting.captionsVisible) {
    badgeEl.dataset.state = "live";
    statusEl.textContent = "субтитры видны, идёт сбор";
  } else if (total) {
    badgeEl.dataset.state = "idle";
    statusEl.textContent = isActive
      ? "субтитры скрыты, собранное сохранено"
      : "прошлая встреча, сбор не идёт";
  } else {
    badgeEl.dataset.state = "idle";
    statusEl.textContent = "субтитры не найдены";
  }

  if (actions) {
    statusEl.textContent += ` · экшенов: ${actions}`;
  }
}

function renderFeed() {
  const captions = visibleCaptions();
  const nearBottom =
    feedEl.scrollHeight - feedEl.scrollTop - feedEl.clientHeight < 60;

  feedEl.replaceChildren();

  if (!captions.length) {
    const empty = document.createElement("p");
    empty.className = "empty";
    empty.textContent = searchEl.value.trim()
      ? "Ничего не найдено по фильтру."
      : "Включи Live Captions во встрече — реплики появятся здесь.";
    feedEl.append(empty);
    return;
  }

  for (const item of captions) {
    const marked = isActionItem(item);
    const line = document.createElement("article");
    line.className = [
      "line",
      item.finalized ? "" : "pending",
      marked ? "marked" : "",
    ]
      .filter(Boolean)
      .join(" ");

    const head = document.createElement("div");
    head.className = "line-head";

    const speaker = document.createElement("span");
    speaker.className = "speaker";
    speaker.textContent = item.speaker;

    const time = document.createElement("span");
    time.className = "time";
    time.textContent = timeOf(item.startedAt);

    const mark = document.createElement("button");
    mark.type = "button";
    mark.className = "mark";
    mark.textContent = marked ? "✓ экшен" : "+ экшен";
    mark.title = marked
      ? "Снять отметку экшен-айтема"
      : "Отметить как экшен-айтем";
    mark.addEventListener("click", () => toggleActionItem(item));

    const text = document.createElement("div");
    text.className = "text";
    text.textContent = item.text;

    head.append(speaker, time, mark);
    line.append(head, text);
    feedEl.append(line);
  }

  if (autoscrollEl.checked && nearBottom) {
    feedEl.scrollTop = feedEl.scrollHeight;
  }
}

function render() {
  renderMeetingSelect();
  renderHeader();
  renderPeople();
  renderFeed();
}

function fileDate() {
  const meeting = selectedMeeting();
  const iso = meeting?.startedAt || meeting?.captions?.[0]?.startedAt;
  const date = iso ? new Date(iso) : new Date();
  const safe = Number.isNaN(date.getTime()) ? new Date() : date;

  const y = safe.getFullYear();
  const m = String(safe.getMonth() + 1).padStart(2, "0");
  const d = String(safe.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function fileTitle() {
  const slug = (selectedMeeting()?.title || "")
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
  return slug || "без-названия";
}

function exportName(ext) {
  return `${fileTitle()} ${fileDate()}.${ext}`;
}

function download(filename, text, type) {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

async function loadStore() {
  const stored = await chrome.storage.local.get([STORAGE_KEY, MARKS_KEY]);
  const raw = stored[STORAGE_KEY]?.meetings ? stored[STORAGE_KEY] : emptyStore();
  store = pruneEmptyMeetings(raw);
  marks = stored[MARKS_KEY] || {};
  render();

  const emptyCount = Object.keys(raw.meetings || {}).length;
  const keptCount = Object.keys(store.meetings).length;
  if (emptyCount !== keptCount) {
    await chrome.storage.local.set({ [STORAGE_KEY]: store });
  }
}

chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== "local") return;

  if (changes[STORAGE_KEY]) {
    const next = changes[STORAGE_KEY].newValue;
    store = next?.meetings ? pruneEmptyMeetings(next) : emptyStore();
  }
  if (changes[MARKS_KEY]) {
    marks = changes[MARKS_KEY].newValue || {};
  }
  if (changes[STORAGE_KEY] || changes[MARKS_KEY]) {
    render();
    scheduleVaultSync();
  }
});

chooseVaultEl.addEventListener("click", async () => {
  try {
    if (vaultHandle) {
      const permission = await vaultHandle.queryPermission({ mode: "readwrite" });
      if (permission !== "granted") {
        const result = await vaultHandle.requestPermission({ mode: "readwrite" });
        if (result === "granted") {
          await setVaultSetting("directory", vaultHandle);
          setVaultStatus(`${vaultHandle.name} · подключено`, "ok");
          scheduleVaultSync();
          return;
        }
      }
    }

    vaultHandle = await window.showDirectoryPicker({ mode: "readwrite" });
    vaultFileNames = {};
    await setVaultSetting("directory", vaultHandle);
    await setVaultSetting("fileNames", vaultFileNames);
    setVaultStatus(`${vaultHandle.name} · подключено`, "ok");
    chooseVaultEl.textContent = "Сменить папку Obsidian";
    scheduleVaultSync();
  } catch (error) {
    if (error.name === "AbortError") return;
    console.error("[Teams Transcript] Не удалось выбрать папку:", error);
    setVaultStatus(`Ошибка доступа: ${error.message}`, "error");
  }
});

meetingSelectEl.addEventListener("change", () => {
  selectedId = meetingSelectEl.value;
  pinnedToSelection = selectedId !== store.activeMeetingId;
  render();
});

searchEl.addEventListener("input", renderFeed);
onlyActionsEl.addEventListener("change", renderFeed);

document.getElementById("copy").addEventListener("click", async () => {
  await navigator.clipboard.writeText(formatTxt(visibleCaptions()));
});

document.getElementById("export-json").addEventListener("click", () => {
  const meeting = selectedMeeting();
  if (!meeting) return;
  download(
    exportName("json"),
    JSON.stringify(exportPayload(meeting), null, 2),
    "application/json"
  );
});

document.getElementById("export-txt").addEventListener("click", () => {
  const meeting = selectedMeeting();
  if (!meeting) return;
  download(exportName("txt"), formatTxt(meeting.captions || []), "text/plain");
});

async function removeSelectedMeeting() {
  const meeting = selectedMeeting();
  if (!meeting) return false;

  const meetings = { ...store.meetings };
  delete meetings[meeting.id];

  const nextMarks = { ...marks };
  delete nextMarks[meeting.id];

  const removedIds = { ...(store.removedIds || {}), [meeting.id]: new Date().toISOString() };
  const recentRemoved = Object.entries(removedIds)
    .sort((a, b) => String(b[1]).localeCompare(String(a[1])))
    .slice(0, 50);

  store = {
    ...store,
    meetings,
    removedIds: Object.fromEntries(recentRemoved),
    activeMeetingId:
      store.activeMeetingId === meeting.id ? null : store.activeMeetingId,
  };
  marks = nextMarks;
  selectedId = null;
  pinnedToSelection = false;

  await chrome.storage.local.set({ [STORAGE_KEY]: store, [MARKS_KEY]: marks });
  render();
  return true;
}

document.getElementById("clear").addEventListener("click", async () => {
  await removeSelectedMeeting();
});

document.getElementById("delete").addEventListener("click", async () => {
  const meeting = selectedMeeting();
  if (!meeting) return;

  const title = meeting.title?.trim() || "без названия";
  if (!confirm(`Удалить транскрипт «${title}»?`)) return;
  await removeSelectedMeeting();
});

loadStore();
restoreVault();
