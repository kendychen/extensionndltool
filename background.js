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

// Đóng tab + window ẩn
function closeTab(tabId) {
  chrome.tabs.get(tabId, (tab) => {
    if (chrome.runtime.lastError) return;
    chrome.windows.remove(tab.windowId).catch(() => {
      chrome.tabs.remove(tabId).catch(() => {});
    });
  });
}

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
        closeTab(tabId);
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
    closeTab(tabId);
    return;
  }

  // Detect proofs/Add
  if (url.includes("proofs/Add") || url.includes("proofs/Manage")) {
    task.resolve({ success: false, error: "proofs/Add" });
    activeTasks.delete(tabId);
    closeTab(tabId);
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
        closeTab(sender.tab.id);
      }
    }
  }
});

async function handleOAuthRequest(message) {
  const { email, password, proxyUrl } = message;
  const { codeVerifier, codeChallenge } = await generatePKCE();

  const authUrl = `https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?client_id=${CLIENT_ID}&scope=${encodeURIComponent(SCOPE)}&redirect_uri=${encodeURIComponent(REDIRECT_URI)}&response_mode=fragment&response_type=code&x-client-SKU=msal.js.browser&x-client-VER=4.4.0&client_info=1&code_challenge=${codeChallenge}&code_challenge_method=S256&login_hint=${encodeURIComponent(email)}`;

  // Tạo cửa sổ ẩn (minimized) để user không thấy
  let tab;
  try {
    const win = await chrome.windows.create({
      url: authUrl,
      type: "popup",
      width: 1,
      height: 1,
      left: -9999,
      top: -9999,
      focused: false,
    });
    tab = win.tabs[0];
  } catch {
    // Fallback: tạo tab ẩn nếu windows.create fail
    tab = await chrome.tabs.create({ url: authUrl, active: false });
  }

  // Track window ID để đóng sau
  const windowId = tab.windowId;

  return new Promise((resolve) => {
    const timeout = setTimeout(() => {
      activeTasks.delete(tab.id);
      chrome.windows.remove(windowId).catch(() => {});
      resolve({ success: false, error: "Timeout (60s)" });
    }, 60000);

    activeTasks.set(tab.id, {
      resolve: (result) => {
        clearTimeout(timeout);
        // Đóng cửa sổ ẩn
        chrome.windows.remove(windowId).catch(() => {});
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
    closeTab(tabId);
  }
}
