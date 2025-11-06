// src/auth.ts
import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  type AuthenticationResult,
} from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./config";

export const msalInstance = new PublicClientApplication(msalConfig);

export async function ensureLogin(): Promise<AuthenticationResult> {
  await msalInstance.initialize();
  const accounts = msalInstance.getAllAccounts();

  if (accounts.length > 0) {
    try {
      return await msalInstance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        return msalInstance.acquireTokenPopup(loginRequest);
      }
      throw e;
    }
  }
  return msalInstance.acquireTokenPopup(loginRequest);
}

/** 统一对外获取 Access Token */
export async function getAccessToken(): Promise<string> {
  const res = await ensureLogin();
  return res.accessToken;
}

/** 兼容别名（如果其他文件用了 getToken 也能工作） */
export const getToken = getAccessToken;
