/**
 * Content Script — inject vào trang Microsoft login.
 * Tự động điền email/password, click buttons, detect trạng thái.
 */

(async function() {
  // Chỉ chạy 1 lần
  if (window.__hotmailHelperInjected) return;
  window.__hotmailHelperInjected = true;

  const url = window.location.href;

  // === Detect redirect with code ===
  if (url.includes("oauthRedirect.html") && url.includes("code=")) {
    const hash = url.split("#")[1] || "";
    const params = new URLSearchParams(hash);
    const code = params.get("code") || "";
    if (code) {
      chrome.runtime.sendMessage({ type: "OAUTH_RESULT", code, anchorMailbox: "" });
      return;
    }
  }

  // === Detect identity blocks ===
  if (url.includes("identity/confirm") || url.includes("Verify?mkt") ||
      url.includes("recover?mkt") || url.includes("Recover?mkt") ||
      url.includes("Abuse?mkt") || url.includes("Update?mkt") ||
      url.includes("RecoverAccount?mkt") || url.includes("family/child-consent")) {
    chrome.runtime.sendMessage({ type: "OAUTH_RESULT", code: "", identityBlocked: true, error: "Identity blocked" });
    return;
  }

  // === Detect proofs/Add ===
  if (url.includes("proofs/Add") || url.includes("proofs/Manage")) {
    chrome.runtime.sendMessage({ type: "OAUTH_RESULT", code: "", error: "proofs/Add" });
    return;
  }

  // === Detect too many attempts ===
  const bodyText = document.body?.innerText || "";
  if (bodyText.includes("tried to sign in too many times")) {
    chrome.runtime.sendMessage({ type: "OAUTH_RESULT", code: "", error: "Too many attempts" });
    return;
  }

  // === Get credentials from background ===
  const creds = await new Promise(resolve => {
    chrome.runtime.sendMessage({ type: "NEED_CREDENTIALS" }, resolve);
  });

  if (!creds?.email) return;

  const step = (msg) => chrome.runtime.sendMessage({ type: "LOGIN_STEP", step: msg });

  // === Wait helper ===
  const wait = (ms) => new Promise(r => setTimeout(r, ms));
  const waitFor = (selector, timeout = 10000) => new Promise((resolve) => {
    const start = Date.now();
    const check = () => {
      const el = document.querySelector(selector);
      if (el) return resolve(el);
      if (Date.now() - start > timeout) return resolve(null);
      setTimeout(check, 300);
    };
    check();
  });

  // === Fill email ===
  const emailInput = await waitFor("#usernameEntry, #i0116, input[name='loginfmt']", 5000);
  if (emailInput) {
    step("Entering email...");
    emailInput.value = creds.email;
    emailInput.dispatchEvent(new Event("input", { bubbles: true }));
    await wait(500);

    const nextBtn = document.querySelector('button[type="submit"], input[type="submit"], #idSIButton9');
    if (nextBtn) nextBtn.click();
    await wait(4000);
  }

  // === Navigate to password input ===
  for (let attempt = 0; attempt < 4; attempt++) {
    const pwInput = document.querySelector('input[type="password"]');
    if (pwInput) break;

    const text = document.body.innerText;

    // Click "Use your password" / "Use a password"
    if (text.includes("Use your password") || text.includes("Use a password")) {
      step("Clicking 'Use your password'...");
      clickTextButton("Use your password") || clickTextButton("Use a password");
      await wait(3000);
      continue;
    }

    // Click "Other ways to sign in"
    if (text.includes("Other ways to sign in") || text.includes("Sign in another way")) {
      step("Clicking 'Other ways to sign in'...");
      clickTextButton("Other ways to sign in") || clickTextButton("Sign in another way");
      await wait(3000);
      clickTextButton("Use your password") || clickTextButton("Use a password") || clickTextButton("Password");
      await wait(3000);
      continue;
    }

    await wait(2000);
  }

  // === Fill password ===
  const passInput = document.querySelector('input[type="password"]');
  if (passInput) {
    step("Entering password...");
    passInput.value = creds.password;
    passInput.dispatchEvent(new Event("input", { bubbles: true }));
    await wait(500);

    const signInBtn = document.querySelector('button[type="submit"]');
    if (signInBtn) signInBtn.click();
    await wait(6000);

    // KMSI - "Stay signed in?"
    handleKMSI();
    await wait(5000);
  }

  // === Check current URL after login ===
  const afterUrl = window.location.href;

  // ar/cancel — click submit
  if (afterUrl.includes("ar/cancel")) {
    step("Handling ar/cancel...");
    const submitBtn = document.getElementById("pageDialogForm_0_submit");
    if (submitBtn) submitBtn.click();
    await wait(8000);
    handleKMSI();
    await wait(5000);
  }

  // Re-check URL
  const finalUrl = window.location.href;

  if (finalUrl.includes("identity/confirm") || finalUrl.includes("Verify?mkt") ||
      finalUrl.includes("recover?mkt") || finalUrl.includes("Abuse?mkt") ||
      finalUrl.includes("Update?mkt") || finalUrl.includes("RecoverAccount?mkt")) {
    chrome.runtime.sendMessage({ type: "OAUTH_RESULT", code: "", identityBlocked: true, error: "Identity blocked" });
    return;
  }

  // fido — go back
  if (finalUrl.includes("fido")) {
    step("Bypassing fido...");
    history.back();
    await wait(3000);
  }

  // Check for code in URL
  if (finalUrl.includes("code=")) {
    const hash = finalUrl.split("#")[1] || "";
    const params = new URLSearchParams(hash);
    const code = params.get("code") || "";
    if (code) {
      chrome.runtime.sendMessage({ type: "OAUTH_RESULT", code });
      return;
    }
  }

  // Wait for redirect
  step("Waiting for redirect...");
  for (let i = 0; i < 15; i++) {
    await wait(1000);
    const curUrl = window.location.href;
    if (curUrl.includes("code=")) {
      const hash = curUrl.split("#")[1] || "";
      const params = new URLSearchParams(hash);
      const code = params.get("code") || "";
      if (code) {
        chrome.runtime.sendMessage({ type: "OAUTH_RESULT", code });
        return;
      }
    }
    if (curUrl.includes("identity/confirm") || curUrl.includes("recover?mkt") || curUrl.includes("Abuse?mkt")) {
      chrome.runtime.sendMessage({ type: "OAUTH_RESULT", code: "", identityBlocked: true });
      return;
    }
  }

  // Timeout — no code found
  chrome.runtime.sendMessage({ type: "OAUTH_RESULT", code: "", error: `No code. URL: ${window.location.href.substring(0, 80)}` });

  // === Helpers ===
  function clickTextButton(text) {
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const t = walker.currentNode.textContent?.trim();
      if (t === text) {
        let el = walker.currentNode.parentElement;
        while (el && el !== document.body) {
          if (el.getAttribute("role") === "button" || el.getAttribute("role") === "link" ||
              el.tagName === "A" || el.tagName === "BUTTON" || el.tagName === "SPAN") {
            el.click();
            return true;
          }
          el = el.parentElement;
        }
      }
    }
    return false;
  }

  function handleKMSI() {
    try {
      const text = document.body.innerText;
      if (text.includes("Stay signed in") || text.includes("Keep me signed in")) {
        document.querySelectorAll("button, input[type='submit']").forEach(el => {
          const t = el.textContent?.trim() || el.value?.trim();
          if (t === "Yes" || t === "Có") el.click();
        });
        const si = document.querySelector("#idSIButton9");
        if (si) si.click();
      }
    } catch {}
  }
})();
