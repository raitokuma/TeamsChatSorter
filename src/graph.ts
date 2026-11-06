const GRAPH = "https://graph.microsoft.com/v1.0";

export interface ChatMessagePreview {
  id: string;
  createdDateTime: string;
  isDeleted?: boolean;
  messageType?: string;
  body?: { content?: string; contentType?: string };
  from?: { user?: { id?: string; displayName?: string } | null; application?: { displayName?: string } | null } | null;
}

export interface Chat {
  id: string;
  topic: string | null;
  chatType: "oneOnOne" | "group" | "meeting" | "unknownFutureValue";
  createdDateTime: string;
  lastUpdatedDateTime: string;
  webUrl?: string;
  tenantId?: string;
  viewpoint?: { isHidden?: boolean; lastMessageReadDateTime?: string } | null;
  lastMessagePreview?: ChatMessagePreview | null;
}

export interface Member {
  userId?: string;
  displayName?: string | null;
  email?: string | null;
}

export type TokenProvider = () => Promise<string>;

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

async function graphFetch(getToken: TokenProvider, url: string, init: RequestInit = {}, attempt = 0): Promise<any> {
  const token = await getToken();
  const res = await fetch(url.startsWith("http") ? url : GRAPH + url, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(init.headers || {}) },
  });
  if ((res.status === 429 || res.status === 503 || res.status === 504) && attempt < 4) {
    const wait = Number(res.headers.get("Retry-After")) || 2 ** attempt * 2;
    await sleep(wait * 1000);
    return graphFetch(getToken, url, init, attempt + 1);
  }
  if (!res.ok) {
    let msg = `${res.status} ${res.statusText}`;
    try {
      const j = await res.json();
      msg = j?.error?.message || msg;
    } catch {}
    throw new Error(`Graph 錯誤：${msg}`);
  }
  return res.json();
}

export async function getMe(getToken: TokenProvider): Promise<{ id: string; displayName: string }> {
  return graphFetch(getToken, "/me?$select=id,displayName");
}

const CHATS_URL =
  "/me/chats?$expand=lastMessagePreview&$orderby=lastMessagePreview/createdDateTime%20desc&$top=50";

/** 讀取聊天（依最後訊息時間由新到舊）。maxPages=1 時只取最新 50 筆，用於增量刷新 */
export async function listChats(
  getToken: TokenProvider,
  maxPages: number,
  onPage?: (chats: Chat[]) => void,
): Promise<Chat[]> {
  const all: Chat[] = [];
  let url: string | undefined = CHATS_URL;
  for (let page = 0; url && page < maxPages; page++) {
    const data: any = await graphFetch(getToken, url);
    const chats: Chat[] = data.value ?? [];
    all.push(...chats);
    onPage?.(chats);
    url = data["@odata.nextLink"];
  }
  return all;
}

/** 以 $batch 一次取得多個聊天的成員（每批 20 個） */
export async function getMembersBatch(getToken: TokenProvider, chatIds: string[]): Promise<Map<string, Member[]>> {
  const result = new Map<string, Member[]>();
  for (let i = 0; i < chatIds.length; i += 20) {
    const slice = chatIds.slice(i, i + 20);
    const body = {
      requests: slice.map((id, idx) => ({
        id: String(idx),
        method: "GET",
        url: `/chats/${encodeURIComponent(id)}/members`,
      })),
    };
    const data: any = await graphFetch(getToken, "/$batch", { method: "POST", body: JSON.stringify(body) });
    for (const r of data.responses ?? []) {
      if (r.status === 200) result.set(slice[Number(r.id)], r.body?.value ?? []);
    }
  }
  return result;
}
