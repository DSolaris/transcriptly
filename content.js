(() => {
  const STORAGE_KEY = "teamsTranscript";
  const SESSION_ID =
    crypto.randomUUID?.() || `s${Date.now()}${Math.random().toString(16).slice(2)}`;

  const SELECTORS = {
    captionLists: [
      "[data-tid='closed-caption-v2-virtual-list-content']",
      "[data-tid='closed-caption-renderer-wrapper']",
      "[data-tid='closed-captions-renderer']",
    ],
    captionItems: [
      ".fui-ChatMessageCompact",
      "[data-tid='closed-captions-v2-items-renderer']",
    ],
    author: "[data-tid='author']",
    text: "[data-tid='closed-caption-text']",
    captionIdHost: "[data-tid='closed-captions-v2-items-renderer']",
    captionIdAttr: "data-lpc-hover-target-id",
    meetingTitle: [
      "[data-tid='call-title']",
      "[data-tid='calling-header-title']",
      "[data-tid='call-header-title']",
      "[data-tid='meeting-title']",
      "[data-tid='call-monitor-title-style-container']",
    ],
    rosterRoots: [
      "[data-tid='roster-section']",
      "[data-tid='calling-roster']",
      "[data-tid='people-pane']",
      "[data-tid='participants-pane']",
    ],
    videoTiles: ["[data-tid='calling-participant-stream']"],
    personName: [
      "[data-tid='roster-participant-name']",
      "[data-tid='participant-name']",
      ".fui-Persona__primaryText",
    ],
    callDuration: ["[data-tid='call-duration']"],
  };

  const FINALIZE_AFTER_MS = 1800;
  const SCAN_EVERY_MS = 700;
  const SAVE_EVERY_MS = 1000;
  const CALL_END_GRACE_MS = 15000;
  const MAX_MEETINGS = 20;

  const captionsById = new Map();
  const order = [];
  const pendingTimers = new Map();
  const suppressedIds = new Set();
  const participants = new Map();

  let meetingTitle = "";
  let meetingId = null;
  let callEpoch = 0;
  let callSeenAt = 0;
  let callWasActive = false;
  let lastClearedAt = null;
  let sessionStartedAt = null;
  let captionsVisible = false;
  let dirty = false;
  let overlay = null;
  let observer = null;
  let observedRoot = null;

  function nowIso() {
    return new Date().toISOString();
  }

  function findFirst(selectors, root = document) {
    for (const selector of selectors) {
      const node = root.querySelector(selector);
      if (node) return node;
    }
    return null;
  }

  function findAll(selectors, root = document) {
    const seen = new Set();
    const nodes = [];
    for (const selector of selectors) {
      root.querySelectorAll(selector).forEach((node) => {
        if (!seen.has(node)) {
          seen.add(node);
          nodes.push(node);
        }
      });
    }
    return nodes;
  }

  function captionIdFrom(item) {
    const host = item.matches?.(SELECTORS.captionIdHost)
      ? item
      : item.querySelector(SELECTORS.captionIdHost) || item;
    const raw = host.getAttribute(SELECTORS.captionIdAttr);
    if (raw) {
      const parts = raw.split("-");
      return parts[parts.length - 1] || raw;
    }
    return null;
  }

  function readCaption(item) {
    const author = item.querySelector(SELECTORS.author)?.textContent?.trim() || "Unknown";
    const text = item.querySelector(SELECTORS.text)?.innerText?.trim() || "";
    if (!text) return null;

    const id = captionIdFrom(item) || `${author}::${text.slice(0, 24)}`;
    return { id, speaker: author, text };
  }

  function ensureSession() {
    if (!sessionStartedAt) {
      sessionStartedAt = nowIso();
    }
  }

  const NAME_NOISE =
    /^(unknown|presenter|attendee|organizer|guest|co-organizer|докладчик|участник|организатор|гость)$/i;

  function cleanName(raw) {
    const name = String(raw || "")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/\s*\((guest|external|гость|внешний)\)$/i, "")
      .trim();

    if (name.length < 2 || name.length > 80) return null;
    if (NAME_NOISE.test(name)) return null;
    return name;
  }

  function noteParticipant(rawName, source) {
    const name = cleanName(rawName);
    if (!name) return;

    const timestamp = nowIso();
    const existing = participants.get(name);

    if (!existing) {
      participants.set(name, {
        name,
        sources: new Set([source]),
        firstSeenAt: timestamp,
        lastSeenAt: timestamp,
      });
      dirty = true;
      return;
    }

    existing.lastSeenAt = timestamp;
    if (!existing.sources.has(source)) {
      existing.sources.add(source);
      dirty = true;
    }
  }

  function collectMeetingTitle() {
    const fromDom = findFirst(SELECTORS.meetingTitle)?.textContent?.trim();
    // document.title carries unread-message counters like "(3) Standup | Microsoft Teams".
    const fromTab = document.title
      .replace(/^\(\d+\)\s*/, "")
      .replace(/\s*\|\s*Microsoft Teams.*$/i, "")
      .trim();

    const title = fromDom || fromTab;
    if (title && title !== meetingTitle) {
      meetingTitle = title;
      dirty = true;
    }
  }

  function collectParticipants() {
    for (const root of findAll(SELECTORS.rosterRoots)) {
      for (const node of findAll(SELECTORS.personName, root)) {
        noteParticipant(node.textContent, "roster");
      }
    }

    for (const tile of findAll(SELECTORS.videoTiles)) {
      const named = findFirst(SELECTORS.personName, tile);
      noteParticipant(named?.textContent || tile.getAttribute("aria-label"), "tile");
    }
  }

  function upsertCaption(parsed) {
    if (suppressedIds.has(parsed.id)) return;
    noteParticipant(parsed.speaker, "caption");
    ensureSession();
    const existing = captionsById.get(parsed.id);
    const timestamp = nowIso();

    if (!existing) {
      const record = {
        id: parsed.id,
        speaker: parsed.speaker,
        text: parsed.text,
        startedAt: timestamp,
        updatedAt: timestamp,
        finalized: false,
      };
      captionsById.set(parsed.id, record);
      order.push(parsed.id);
      dirty = true;
      scheduleFinalize(parsed.id);
      return;
    }

    if (existing.text !== parsed.text || existing.speaker !== parsed.speaker) {
      existing.speaker = parsed.speaker;
      existing.text = parsed.text;
      existing.updatedAt = timestamp;
      existing.finalized = false;
      dirty = true;
      scheduleFinalize(parsed.id);
    } else {
      scheduleFinalize(parsed.id);
    }
  }

  function scheduleFinalize(id) {
    const prev = pendingTimers.get(id);
    if (prev) clearTimeout(prev);
    pendingTimers.set(
      id,
      setTimeout(() => {
        const record = captionsById.get(id);
        if (record && !record.finalized) {
          record.finalized = true;
          record.updatedAt = nowIso();
          dirty = true;
        }
        pendingTimers.delete(id);
      }, FINALIZE_AFTER_MS)
    );
  }

  // The thread id stays stable for the whole call, so it survives reloads and
  // SPA navigation. Without it (call not joined yet) captions belong to a
  // per-tab bucket that is rotated every time a call ends.
  function resolveMeetingId() {
    const href = decodeURIComponent(location.href);
    const thread = href.match(/19:meeting_[A-Za-z0-9_-]+@thread\.v2/i);
    return thread ? `thread:${thread[0].toLowerCase()}` : `tab:${SESSION_ID}:${callEpoch}`;
  }

  function currentMeetingId() {
    if (!meetingId) meetingId = resolveMeetingId();
    return meetingId;
  }

  function resetLocal() {
    captionsById.clear();
    order.length = 0;
    participants.clear();
    suppressedIds.clear();
    for (const timer of pendingTimers.values()) clearTimeout(timer);
    pendingTimers.clear();
    sessionStartedAt = null;
    meetingTitle = "";
    lastClearedAt = null;
    dirty = false;
  }

  // Everything captured so far is already in storage under the previous id,
  // so switching meetings only needs to drop the in-memory copy.
  function switchMeeting(nextId) {
    if (nextId === meetingId) return;
    const hadMeeting = meetingId !== null;
    meetingId = nextId;
    if (hadMeeting) resetLocal();
  }

  function trackCallBoundary() {
    const inCall = Boolean(findFirst(SELECTORS.callDuration)) || captionsVisible;

    if (inCall) {
      callSeenAt = Date.now();
      callWasActive = true;
    } else if (callWasActive && Date.now() - callSeenAt > CALL_END_GRACE_MS) {
      callWasActive = false;
      callEpoch += 1;
      switchMeeting(resolveMeetingId());
      return;
    }

    switchMeeting(resolveMeetingId());
  }

  // Captions still on screen when the user clears the log would otherwise be
  // re-captured by the next scan, so they stay suppressed until Teams recycles them.
  function applyClear(clearedAt) {
    if (!clearedAt || clearedAt === lastClearedAt) return false;
    lastClearedAt = clearedAt;

    if (!order.length && !participants.size) return false;

    for (const id of order) suppressedIds.add(id);
    captionsById.clear();
    order.length = 0;
    participants.clear();
    for (const timer of pendingTimers.values()) clearTimeout(timer);
    pendingTimers.clear();
    sessionStartedAt = null;
    dirty = false;
    updateOverlay();
    return true;
  }

  function snapshot() {
    return order
      .map((id) => {
        const record = captionsById.get(id);
        if (!record) return null;
        return { ...record, key: `${SESSION_ID}:${record.id}`, sessionId: SESSION_ID };
      })
      .filter(Boolean);
  }

  function participantsSnapshot() {
    return [...participants.values()].map((item) => ({
      name: item.name,
      sources: [...item.sources],
      firstSeenAt: item.firstSeenAt,
      lastSeenAt: item.lastSeenAt,
    }));
  }

  function scan() {
    const list = findFirst(SELECTORS.captionLists);
    const items = list
      ? findAll(SELECTORS.captionItems, list)
      : findAll(SELECTORS.captionItems);

    const visible = items.length > 0 || Boolean(findFirst([SELECTORS.text]));
    if (visible !== captionsVisible) {
      captionsVisible = visible;
      dirty = true;
    }

    if (list && observedRoot !== list) {
      attachObserver(list);
    }

    trackCallBoundary();

    for (const item of items) {
      const parsed = readCaption(item);
      if (parsed) upsertCaption(parsed);
    }

    collectMeetingTitle();
    collectParticipants();
    updateOverlay();
  }

  function attachObserver(root) {
    if (observer) observer.disconnect();
    observedRoot = root;
    observer = new MutationObserver(() => scan());
    observer.observe(root, {
      childList: true,
      subtree: true,
      characterData: true,
    });
  }

  function ensureOverlay() {
    if (overlay) return overlay;
    overlay = document.createElement("div");
    overlay.id = "teams-transcript-overlay";
    overlay.innerHTML = `
      <strong>Teams Transcript</strong>
      <span data-role="status">ожидание субтитров</span>
    `;
    document.documentElement.appendChild(overlay);
    return overlay;
  }

  function plural(count) {
    const mod100 = count % 100;
    const mod10 = count % 10;
    if (mod100 >= 11 && mod100 <= 14) return "реплик";
    if (mod10 === 1) return "реплика";
    if (mod10 >= 2 && mod10 <= 4) return "реплики";
    return "реплик";
  }

  function updateOverlay() {
    const el = ensureOverlay();
    const status = el.querySelector("[data-role='status']");
    const count = order.length;

    if (!extensionAlive()) {
      status.textContent = "расширение обновлено — перезагрузи вкладку";
      el.dataset.state = "stale";
      return;
    }
    if (!captionsVisible) {
      status.textContent = "включи Live Captions в меню встречи";
      el.dataset.state = "waiting";
      return;
    }
    status.textContent = `${count} ${plural(count)}`;
    el.dataset.state = "capturing";
  }

  function extensionAlive() {
    try {
      return Boolean(chrome.runtime?.id);
    } catch {
      return false;
    }
  }

  async function persist() {
    if (!dirty || !extensionAlive()) return;
    dirty = false;

    const stored = (await chrome.storage.local.get(STORAGE_KEY))[STORAGE_KEY] || {};
    const meetings = { ...(stored.meetings || {}) };
    const removedIds = { ...(stored.removedIds || {}) };
    const id = currentMeetingId();
    const bucket = meetings[id] || {};

    if (applyClear(bucket.clearedAt || removedIds[id])) return;

    const byKey = new Map((bucket.captions || []).map((item) => [item.key, item]));
    for (const record of snapshot()) {
      byKey.set(record.key, record);
    }

    const captions = [...byKey.values()].sort((a, b) =>
      String(a.startedAt).localeCompare(String(b.startedAt))
    );

    const byName = new Map((bucket.participants || []).map((item) => [item.name, item]));
    for (const person of participantsSnapshot()) {
      const previous = byName.get(person.name);
      byName.set(
        person.name,
        previous
          ? {
              ...person,
              sources: [...new Set([...(previous.sources || []), ...person.sources])],
              firstSeenAt:
                previous.firstSeenAt < person.firstSeenAt
                  ? previous.firstSeenAt
                  : person.firstSeenAt,
            }
          : person
      );
    }

    if (!captions.length) {
      if (!meetings[id] && stored.activeMeetingId !== id) return;
      delete meetings[id];
    } else {
      delete removedIds[id];
      meetings[id] = {
        id,
        title: meetingTitle || bucket.title || "",
        url: location.href,
        startedAt: bucket.startedAt || sessionStartedAt || nowIso(),
        lastUpdatedAt: nowIso(),
        clearedAt: lastClearedAt,
        captionsVisible,
        participants: [...byName.values()].sort((a, b) => a.name.localeCompare(b.name)),
        captions,
      };
    }

    await chrome.storage.local.set({
      [STORAGE_KEY]: {
        version: 2,
        activeMeetingId: captions.length ? id : stored.activeMeetingId === id ? null : stored.activeMeetingId,
        lastUpdatedAt: nowIso(),
        removedIds,
        meetings: keepRecent(meetings),
      },
    });
  }

  // storage.local has a 10 MB quota, so old meetings drop out instead of
  // making writes fail once enough transcripts pile up.
  function keepRecent(meetings) {
    const recent = Object.values(meetings)
      .filter((meeting) => meeting.captions?.length)
      .sort((a, b) =>
        String(b.lastUpdatedAt || b.startedAt).localeCompare(
          String(a.lastUpdatedAt || a.startedAt)
        )
      )
      .slice(0, MAX_MEETINGS);

    return Object.fromEntries(recent.map((meeting) => [meeting.id, meeting]));
  }

  // One-time move of the pre-v2 flat transcript into its own bucket.
  async function migrateLegacy(stored) {
    if (stored.meetings || !stored.captions?.length) return;

    const id = `legacy:${stored.sessionStartedAt || stored.lastUpdatedAt || nowIso()}`;
    await chrome.storage.local.set({
      [STORAGE_KEY]: {
        version: 2,
        activeMeetingId: null,
        lastUpdatedAt: nowIso(),
        meetings: {
          [id]: {
            id,
            title: stored.meeting?.title || "",
            url: stored.meeting?.url || stored.meetingHint || "",
            startedAt: stored.sessionStartedAt || stored.captions[0]?.startedAt || nowIso(),
            lastUpdatedAt: stored.lastUpdatedAt || nowIso(),
            clearedAt: stored.clearedAt || null,
            captionsVisible: false,
            participants: stored.participants || [],
            captions: stored.captions,
          },
        },
      },
    });
  }

  async function start() {
    ensureOverlay();

    if (extensionAlive()) {
      try {
        const stored = (await chrome.storage.local.get(STORAGE_KEY))[STORAGE_KEY] || {};
        await migrateLegacy(stored);
        lastClearedAt = stored.meetings?.[currentMeetingId()]?.clearedAt || null;
      } catch (error) {
        console.warn("[Teams Transcript] не удалось прочитать состояние:", error);
      }

      chrome.storage.onChanged.addListener((changes, area) => {
        if (area !== "local" || !changes[STORAGE_KEY]) return;
        const next = changes[STORAGE_KEY].newValue || {};
        const id = currentMeetingId();
        applyClear(next.meetings?.[id]?.clearedAt || next.removedIds?.[id]);
      });
    }

    scan();
    setInterval(scan, SCAN_EVERY_MS);
    setInterval(() => {
      persist().catch((error) => {
        dirty = true;
        console.warn("[Teams Transcript] не удалось сохранить субтитры:", error);
        updateOverlay();
      });
    }, SAVE_EVERY_MS);
  }

  start();
})();
