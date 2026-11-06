export const config = {
  clientId: import.meta.env.VITE_CLIENT_ID as string,
  tenantId: (import.meta.env.VITE_TENANT_ID as string) || "organizations",
  /** Chat.Read：讀取聊天清單、成員與最後訊息預覽；User.Read：取得自己的使用者 ID */
  scopes: ["User.Read", "Chat.Read"],
  /** 自動刷新間隔（毫秒） */
  refreshIntervalMs: 60_000,
  /** 最多讀取幾頁（每頁 50 筆） */
  maxPages: 60,
};

if (!config.clientId || config.clientId.startsWith("00000000")) {
  console.warn("VITE_CLIENT_ID 尚未設定，請參考 .env.example");
}
