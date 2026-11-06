import { app } from "@microsoft/teams-js";
import { config } from "./config";
import { getTokenInteractive, getTokenSilent, initAuth, NeedsInteractionError } from "./auth";
import { getMe, getMembersBatch, listChats, type Chat, type Member } from "./graph";

type Filter = "all" | "unread" | "oneOnOne" | "group" | "meeting";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const listEl = $<HTMLUListElement>("list");
const statusEl = $("status");
const errorEl = $("error");
const emptyEl = $("empty");
const signinEl = $("signin");
const searchEl = $<HTMLInputElement>("search");

const state = {
  inTeams: false,
  me: { id: "", displayName: "" },
  chats: new Map<string, Chat>(),
  members: loadMemberCache(),
  filter: "all" as Filter,
  query: "",
  loading: false,
  lastSync: 0,
};

// ---------- 成員快取（減少 Graph 呼叫） ----------
const MEMBER_CACHE_KEY = "tcs.members.v1";
function loadMemberCache(): Map<string, Member[]> {
  try {
    return new Map(Object.entries(JSON.parse(localStorage.getItem(MEMBER_CACHE_KEY) || "{}")));
  } catch {
    return new Map();
  }
}
function saveMemberCache() {
  try {
    localStorage.setItem(MEMBER_CACHE_KEY, JSON.stringify(Object.fromEntries(state.members)));
  } catch {}
}

// ---------- 工具函式 ----------
function activityTime(c: Chat): number {
  return Date.parse(c.lastMessagePreview?.createdDateTime || c.lastUpdatedDateTime || c.createdDateTime) || 0;
}

function isUnread(c: Chat): boolean {
  const msg = c.lastMessagePreview;
  if (!msg || msg.messageType === "systemEventMessage") return false;
  if (msg.from?.user?.id && msg.from.user.id === state.me.id) return false;
  const read = Date.parse(c.viewpoint?.lastMessageReadDateTime || "") || 0;
  return Date.parse(msg.createdDateTime) > read;
}

function stripHtml(html: string): string {
  const doc = new DOMParser().parseFromString(html, "text/html");
  doc.querySelectorAll("attachment").forEach((a) => a.replaceWith("[附件]"));
  return (doc.body.textContent || "").replace(/\s+/g, " ").trim();
}

function chatTitle(c: Chat): string {
  if (c.topic) return c.topic;
  const members = state.members.get(c.id);
  if (members?.length) {
    const others = members.filter((m) => m.userId !== state.me.id).map((m) => m.displayName || m.email || "未知");
    if (others.length === 0) return `${state.me.displayName}（你自己）`;
    return others.join("、");
  }
  if (c.chatType === "oneOnOne") {
    const from = c.lastMessagePreview?.from?.user;
    if (from?.displayName && from.id !== state.me.id) return from.displayName;
  }
  return "（載入中…）";
}

function previewText(c: Chat): string {
  const m = c.lastMessagePreview;
  if (!m) return "尚無訊息";
  if (m.isDeleted) return "此訊息已刪除";
  if (m.messageType === "systemEventMessage") return "（系統事件）";
  let text = m.body?.content ? (m.body.contentType === "html" ? stripHtml(m.body.content) : m.body.content) : "";
  if (!text) text = "[附件或卡片]";
  const fromId = m.from?.user?.id;
  const sender =
    fromId && fromId === state.me.id
      ? "你"
      : c.chatType !== "oneOnOne"
        ? (m.from?.user?.displayName || m.from?.application?.displayName || "").split(/[ (（]/)[0]
        : "";
  return sender ? `${sender}：${text}` : text;
}

const timeFmt = new Intl.DateTimeFormat("zh-TW", { hour: "2-digit", minute: "2-digit", hour12: false });
const dateFmt = new Intl.DateTimeFormat("zh-TW", { month: "numeric", day: "numeric" });
const fullDateFmt = new Intl.DateTimeFormat("zh-TW", { year: "numeric", month: "numeric", day: "numeric" });
const weekdayFmt = new Intl.DateTimeFormat("zh-TW", { weekday: "short" });
const fullFmt = new Intl.DateTimeFormat("zh-TW", { dateStyle: "full", timeStyle: "short" });

function startOfDay(d: Date) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

function shortTime(t: number): string {
  if (!t) return "";
  const d = new Date(t);
  const today = startOfDay(new Date());
  const day = startOfDay(d);
  if (day === today) return timeFmt.format(d);
  if (day === today - 86_400_000) return "昨天";
  if (today - day < 6 * 86_400_000) return weekdayFmt.format(d);
  if (d.getFullYear() === new Date().getFullYear()) return dateFmt.format(d);
  return fullDateFmt.format(d);
}

function bucket(t: number): string {
  const today = startOfDay(new Date());
  const day = startOfDay(new Date(t));
  if (day >= today) return "今天";
  if (day >= today - 86_400_000) return "昨天";
  if (day >= today - 6 * 86_400_000) return "過去 7 天";
  if (day >= today - 30 * 86_400_000) return "過去 30 天";
  return "更早";
}

function initials(name: string): string {
  const n = name.replace(/[（(].*$/, "").trim();
  if (!n) return "?";
  const cjk = n.match(/[\u3400-\u9fff]/);
  if (cjk) return cjk[0];
  const parts = n.split(/[\s、,]+/).filter(Boolean);
  return ((parts[0]?.[0] || "") + (parts[1]?.[0] || "")).toUpperCase() || "?";
}

function hue(s: string): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = Math.imul(h ^ s.charCodeAt(i), 2654435761) >>> 0;
  return (h >>> 7) % 360;
}

const typeLabel: Record<string, string> = { oneOnOne: "", group: "群組", meeting: "會議" };

// ---------- 繪製 ----------
function render() {
  const q = state.query.toLowerCase();
  const sorted = [...state.chats.values()]
    .filter((c) => !c.viewpoint?.isHidden)
    .filter((c) => {
      if (state.filter === "unread") return isUnread(c);
      if (state.filter !== "all") return c.chatType === state.filter;
      return true;
    })
    .filter((c) => !q || (chatTitle(c) + " " + previewText(c)).toLowerCase().includes(q))
    .sort((a, b) => activityTime(b) - activityTime(a));

  const frag = document.createDocumentFragment();
  let lastBucket = "";
  for (const c of sorted) {
    const t = activityTime(c);
    const b = bucket(t);
    if (b !== lastBucket) {
      const h = document.createElement("li");
      h.className = "bucket";
      h.textContent = b;
      frag.appendChild(h);
      lastBucket = b;
    }
    const title = chatTitle(c);
    const unread = isUnread(c);
    const li = document.createElement("li");
    li.className = "chat" + (unread ? " unread" : "");
    li.tabIndex = 0;
    li.dataset.id = c.id;
    li.title = fullFmt.format(new Date(t));

    const avatar = document.createElement("div");
    avatar.className = `avatar ${c.chatType}`;
    avatar.style.setProperty("--h", String(hue(c.id)));
    avatar.textContent = c.chatType === "meeting" ? "📅" : initials(title);

    const body = document.createElement("div");
    body.className = "body";
    const top = document.createElement("div");
    top.className = "row";
    const name = document.createElement("span");
    name.className = "name";
    name.textContent = title;
    const time = document.createElement("span");
    time.className = "time";
    time.textContent = shortTime(t);
    top.append(name);
    if (typeLabel[c.chatType]) {
      const tag = document.createElement("span");
      tag.className = "tag";
      tag.textContent = typeLabel[c.chatType];
      top.append(tag);
    }
    top.append(time);
    const prev = document.createElement("div");
    prev.className = "preview";
    prev.textContent = previewText(c);
    body.append(top, prev);

    li.append(avatar, body);
    if (unread) {
      const dot = document.createElement("span");
      dot.className = "dot";
      dot.setAttribute("aria-label", "未讀");
      li.append(dot);
    }
    frag.appendChild(li);
  }
  listEl.replaceChildren(frag);
  emptyEl.hidden = sorted.length > 0 || state.loading;
}

function setStatus(text: string) {
  statusEl.textContent = text;
}

function showError(msg: string | null) {
  errorEl.hidden = !msg;
  errorEl.textContent = msg || "";
}

// ---------- 開啟聊天 ----------
function chatLink(c: Chat): string {
  if (c.webUrl) return c.webUrl;
  const tenant = c.tenantId ? `?tenantId=${c.tenantId}` : "";
  return `https://teams.microsoft.com/l/chat/${encodeURIComponent(c.id)}/0${tenant}`;
}

async function openChat(id: string) {
  const c = state.chats.get(id);
  if (!c) return;
  const url = chatLink(c);
  if (state.inTeams) {
    try {
      await app.openLink(url);
      return;
    } catch (e) {
      console.warn("openLink 失敗，改用新視窗", e);
    }
  }
  window.open(url, "_blank", "noopener");
}

listEl.addEventListener("click", (e) => {
  const li = (e.target as HTMLElement).closest<HTMLLIElement>("li.chat");
  if (li?.dataset.id) openChat(li.dataset.id);
});
listEl.addEventListener("keydown", (e) => {
  const li = (e.target as HTMLElement).closest<HTMLLIElement>("li.chat");
  if (!li) return;
  if (e.key === "Enter") openChat(li.dataset.id!);
  if (e.key === "ArrowDown" || e.key === "ArrowUp") {
    e.preventDefault();
    const items = [...listEl.querySelectorAll<HTMLLIElement>("li.chat")];
    const i = items.indexOf(li) + (e.key === "ArrowDown" ? 1 : -1);
    items[Math.max(0, Math.min(items.length - 1, i))]?.focus();
  }
});

// ---------- 資料載入 ----------
async function fillMissingMembers() {
  const need = [...state.chats.values()]
    .filter((c) => !c.topic && c.chatType !== "meeting" && !state.members.has(c.id))
    .map((c) => c.id);
  if (!need.length) return;
  const got = await getMembersBatch(getTokenSilent, need);
  got.forEach((v, k) =>
    state.members.set(
      k,
      v.map((m) => ({ userId: m.userId, displayName: m.displayName, email: m.email })),
    ),
  );
  saveMemberCache();
  render();
}

async function sync(full: boolean) {
  if (state.loading) return;
  state.loading = true;
  showError(null);
  setStatus(full ? "載入中…" : "更新中…");
  try {
    if (!state.me.id) state.me = await getMe(getTokenSilent);
    let count = 0;
    const fresh = await listChats(getTokenSilent, full ? config.maxPages : 1, (page) => {
      for (const c of page) state.chats.set(c.id, c);
      count += page.length;
      if (full) {
        setStatus(`載入中…（${count}）`);
        render();
      }
    });
    if (full) {
      // 完整同步時移除已不存在的聊天
      const ids = new Set(fresh.map((c) => c.id));
      for (const id of [...state.chats.keys()]) if (!ids.has(id)) state.chats.delete(id);
    }
    state.lastSync = Date.now();
    render();
    await fillMissingMembers();
    setStatus(`${state.chats.size} 個聊天 · ${timeFmt.format(new Date())} 更新`);
  } catch (e) {
    if (e instanceof NeedsInteractionError) {
      signinEl.hidden = false;
      setStatus("");
    } else {
      showError((e as Error).message);
      setStatus("");
    }
  } finally {
    state.loading = false;
    emptyEl.hidden = listEl.childElementCount > 0;
  }
}

// ---------- UI 事件 ----------
searchEl.addEventListener("input", () => {
  state.query = searchEl.value.trim();
  render();
});

document.querySelectorAll<HTMLButtonElement>(".chip").forEach((btn) =>
  btn.addEventListener("click", () => {
    document.querySelectorAll(".chip").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    state.filter = btn.dataset.filter as Filter;
    render();
  }),
);

$("refresh").addEventListener("click", () => sync(true));

$("signin-btn").addEventListener("click", async () => {
  try {
    await getTokenInteractive();
    signinEl.hidden = true;
    await sync(true);
  } catch (e) {
    showError(`登入失敗：${(e as Error).message}`);
  }
});

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible" && Date.now() - state.lastSync > 15_000) sync(false);
});

// ---------- 主題 ----------
function applyTheme(theme?: string) {
  const t =
    theme === "dark" || theme === "contrast"
      ? theme
      : theme === "default"
        ? "light"
        : matchMedia("(prefers-color-scheme: dark)").matches
          ? "dark"
          : "light";
  document.documentElement.dataset.theme = t;
}

// ---------- 啟動 ----------
async function start() {
  try {
    await app.initialize();
    state.inTeams = true;
    const ctx = await app.getContext();
    applyTheme(ctx.app.theme);
    app.registerOnThemeChangeHandler(applyTheme);
    app.notifySuccess();
  } catch {
    state.inTeams = false; // 在一般瀏覽器中開啟
    applyTheme();
  }

  try {
    await initAuth(state.inTeams);
  } catch (e) {
    showError(`驗證初始化失敗：${(e as Error).message}`);
    return;
  }

  await sync(true);
  setInterval(() => {
    if (document.visibilityState === "visible") sync(false);
  }, config.refreshIntervalMs);
  // 每 30 分鐘做一次完整同步
  setInterval(() => sync(true), 30 * 60_000);
  // 每分鐘重繪一次以更新「今天/昨天」等相對時間
  setInterval(render, 60_000);
}

start();
