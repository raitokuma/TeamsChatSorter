import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { config } from "./config";

let pca: IPublicClientApplication;

/**
 * 在 Teams 內使用 Nested App Authentication（NAA），由 Teams 代為取得 Graph 權杖，
 * 不需要後端與 OBO；在一般瀏覽器中（本機測試）改用標準 MSAL 彈出視窗登入。
 */
export async function initAuth(inTeams: boolean): Promise<void> {
  const msalConfig: Configuration = {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
      supportsNestedAppAuth: true,
      redirectUri: new URL(`${import.meta.env.BASE_URL}redirect.html`, location.origin).href,
    },
    cache: { cacheLocation: "localStorage" },
  };

  if (inTeams) {
    pca = await createNestablePublicClientApplication(msalConfig);
  } else {
    const std = new PublicClientApplication(msalConfig);
    await std.initialize();
    await std.handleRedirectPromise().catch(() => undefined);
    pca = std;
  }
}

function currentAccount(): AccountInfo | null {
  return pca.getActiveAccount() ?? pca.getAllAccounts()[0] ?? null;
}

function remember(result: AuthenticationResult): string {
  if (result.account) pca.setActiveAccount(result.account);
  return result.accessToken;
}

export class NeedsInteractionError extends Error {}

/** 靜默取得權杖；若需要使用者互動則丟出 NeedsInteractionError */
export async function getTokenSilent(): Promise<string> {
  const scopes = config.scopes;
  const account = currentAccount();
  try {
    if (account) return remember(await pca.acquireTokenSilent({ scopes, account }));
    return remember(await pca.ssoSilent({ scopes }));
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError || !account) {
      throw new NeedsInteractionError((e as Error)?.message ?? "需要登入");
    }
    // 帳戶存在但靜默失敗，再試一次 ssoSilent
    try {
      return remember(await pca.ssoSilent({ scopes }));
    } catch (e2) {
      throw new NeedsInteractionError((e2 as Error)?.message ?? "需要登入");
    }
  }
}

/** 互動式登入（必須由使用者點擊觸發，避免彈出視窗被封鎖） */
export async function getTokenInteractive(): Promise<string> {
  return remember(await pca.acquireTokenPopup({ scopes: config.scopes }));
}
