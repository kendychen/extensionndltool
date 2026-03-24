/**
 * Background Service Worker — nhận lệnh từ web app, mở tab OAuth, trả kết quả.
 */

const CLIENT_ID = "9199bf20-a13f-4107-85dc-02114787ef48";
const REDIRECT_URI = "https://outlook.live.com/mail/oauthRedirect.html";
const SCOPE = "service::outlook.office.com::MBI_SSL openid profile offline_access";

// PKCE helpers
function base64UrlEncode(buffer) {
  const bytes = new Uint8Array(buffer);
  let str = "";
  bytes.forEach(b => str += String.fromCharCode(b));
  return btoa(str).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

async function generatePKCE() {
  const buffer = crypto.getRandomValues(new Uint8Array(32));
  const codeVerifier = base64UrlEncode(buffer);
  const encoded = new TextEncoder().encode(codeVerifier);
  const hash = await crypto.subtle.digest("SHA-256", encoded);
  const codeChallenge = base64UrlEncode(hash);
  return { codeVerifier, codeChallenge };
}

// Active tasks: tabId → { resolve, reject, email, password, codeChallenge }
const activeTasks = new Map();

// Listen for messages from web app
chrome.runtime.onMessageExternal.addListener((message, sender, sendResponse) => {
  if (message.type === "PING") {
    sendResponse({ success: true, version: chrome.runtime.getManifest().version });
    return true;
  }

  if (message.type === "GET_OAUTH_CODE") {
    handleOAuthRequest(message).then(sendResponse).catch(err => {
      sendResponse({ success: false, error: err.message || String(err) });
    });
    return true; // async response
  }

  if (message.type === "CANCEL_OAUTH") {
    cancelTask(message.taskId);
    sendResponse({ success: true });
    return true;
  }
});

// Listen for tab URL changes
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  const task = activeTasks.get(tabId);
  if (!task || !changeInfo.url) return;

  const url = changeInfo.url;

  // Got OAuth code in redirect
  if (url.includes("oauthRedirect.html") && url.includes("code=")) {
    const hash = url.split("#")[1] || "";
    const params = new URLSearchParams(hash);
    const code = params.get("code") || "";

    if (code) {
      // Navigate to OWA to get DefaultAnchorMailbox cookie
      getAnchorMailbox(tabId).then(anchorMailbox => {
        task.resolve({ success: true, code, anchorMailbox, codeVerifier: task.codeVerifier });
        activeTasks.delete(tabId);
        chrome.tabs.remove(tabId).catch(() => {});
      });
      return;
    }
  }

  // Detect identity blocks
  if (url.includes("identity/confirm") || url.includes("Verify?mkt") ||
      url.includes("recover?mkt") || url.includes("Recover?mkt") ||
      url.includes("Abuse?mkt") || url.includes("Update?mkt") ||
      url.includes("RecoverAccount?mkt")) {
    task.resolve({ success: false, identityBlocked: true, error: "Identity blocked" });
    activeTasks.delete(tabId);
    chrome.tabs.remove(tabId).catch(() => {});
    return;
  }

  // Detect proofs/Add
  if (url.includes("proofs/Add") || url.includes("proofs/Manage")) {
    task.resolve({ success: false, error: "proofs/Add" });
    activeTasks.delete(tabId);
    chrome.tabs.remove(tabId).catch(() => {});
    return;
  }
});

// Listen for tab close
chrome.tabs.onRemoved.addListener((tabId) => {
  const task = activeTasks.get(tabId);
  if (task) {
    task.resolve({ success: false, error: "Tab closed by user" });
    activeTasks.delete(tabId);
  }
});

// Listen for content script messages
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === "LOGIN_STEP" && sender.tab?.id) {
    const task = activeTasks.get(sender.tab.id);
    if (task && task.onStep) {
      task.onStep(message.step);
    }
  }

  if (message.type === "NEED_CREDENTIALS" && sender.tab?.id) {
    const task = activeTasks.get(sender.tab.id);
    if (task) {
      sendResponse({ email: task.email, password: task.password });
    }
    return true;
  }

  if (message.type === "OAUTH_RESULT" && sender.tab?.id) {
    const task = activeTasks.get(sender.tab.id);
    if (task) {
      task.resolve({
        success: !!message.code,
        code: message.code || "",
        anchorMailbox: message.anchorMailbox || "",
        codeVerifier: task.codeVerifier,
        error: message.error,
        identityBlocked: message.identityBlocked,
      });
      activeTasks.delete(sender.tab.id);
      if (!message.keepTab) {
        chrome.tabs.remove(sender.tab.id).catch(() => {});
      }
    }
  }
});

async function handleOAuthRequest(message) {
  const { email, password, proxyUrl } = message;
  const { codeVerifier, codeChallenge } = await generatePKCE();

  const authUrl = `https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?client_id=${CLIENT_ID}&scope=${encodeURIComponent(SCOPE)}&redirect_uri=${encodeURIComponent(REDIRECT_URI)}&response_mode=fragment&response_type=code&x-client-SKU=msal.js.browser&x-client-VER=4.4.0&client_info=1&code_challenge=${codeChallenge}&code_challenge_method=S256&login_hint=${encodeURIComponent(email)}`;

  // Open tab (hidden if possible)
  const tab = await chrome.tabs.create({ url: authUrl, active: false });

  return new Promise((resolve) => {
    const timeout = setTimeout(() => {
      activeTasks.delete(tab.id);
      chrome.tabs.remove(tab.id).catch(() => {});
      resolve({ success: false, error: "Timeout (60s)" });
    }, 60000);

    activeTasks.set(tab.id, {
      resolve: (result) => {
        clearTimeout(timeout);
        resolve(result);
      },
      email,
      password,
      codeVerifier,
      codeChallenge,
    });
  });
}

async function getAnchorMailbox(tabId) {
  try {
    await chrome.tabs.update(tabId, { url: "https://outlook.live.com/mail/" });
    await new Promise(r => setTimeout(r, 5000));

    const results = await chrome.scripting.executeScript({
      target: { tabId },
      func: () => {
        const cookies = document.cookie.split(";");
        for (const c of cookies) {
          const [name, ...val] = c.trim().split("=");
          if (name === "DefaultAnchorMailbox") return val.join("=");
        }
        return "";
      },
    });

    return results?.[0]?.result || "";
  } catch {
    return "";
  }
}

function cancelTask(taskId) {
  for (const [tabId, task] of activeTasks) {
    task.resolve({ success: false, error: "Cancelled" });
    activeTasks.delete(tabId);
    chrome.tabs.remove(tabId).catch(() => {});
  }
}
