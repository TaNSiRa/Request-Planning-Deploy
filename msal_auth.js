(function () {
  let msalScriptPromise = null;
  let msalApp = null;

  function getAuthConfig() {
    const config = window.arAuthConfig || {};
    return {
      clientId: String(config.clientId || "").trim(),
      tenantId: String(config.tenantId || "").trim(),
      scopes: String(config.scopes || "openid profile email User.Read")
        .split(/[,\s]+/)
        .map((scope) => scope.trim())
        .filter(Boolean)
    };
  }

  function loadMsalScript() {
    if (window.msal) return Promise.resolve();
    if (msalScriptPromise) return msalScriptPromise;
    msalScriptPromise = loadScriptCandidates([
      "vendor/msal-browser.min.js",
      "https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js"
    ]);
    return msalScriptPromise;
  }

  function loadScriptCandidates(sources) {
    return new Promise((resolve, reject) => {
      const tryNext = (index) => {
        if (window.msal) {
          resolve();
          return;
        }
        if (index >= sources.length) {
          reject(new Error("Unable to load Microsoft authentication script."));
          return;
        }
        const src = sources[index];
        const script = document.createElement("script");
        script.src = src;
        script.async = true;
        script.onload = () => resolve();
        script.onerror = () => {
          script.remove();
          tryNext(index + 1);
        };
        document.head.appendChild(script);
      };
      tryNext(0);
    });
  }

  async function getMsalApp() {
    const config = getAuthConfig();
    if (!config.clientId || !config.tenantId) {
      throw new Error("Microsoft 365 login is not configured.");
    }
    await loadMsalScript();
    if (!msalApp) {
      msalApp = new window.msal.PublicClientApplication({
        auth: {
          clientId: config.clientId,
          authority: "https://login.microsoftonline.com/" + config.tenantId,
          redirectUri: window.location.origin + window.location.pathname
        },
        cache: {
          cacheLocation: "sessionStorage",
          storeAuthStateInCookie: false
        }
      });
      if (typeof msalApp.initialize === "function") await msalApp.initialize();
    }
    return { app: msalApp, scopes: config.scopes };
  }

  async function arMicrosoftSignIn() {
    const { app, scopes } = await getMsalApp();
    const response = await app.loginPopup({ scopes, prompt: "select_account" });
    const account = response.account || {};
    const claims = response.idTokenClaims || account.idTokenClaims || {};
    const email = account.username || claims.preferred_username || claims.email || claims.upn || "";
    if (!response.idToken) throw new Error("Microsoft did not return an ID token.");
    return {
      email,
      name: account.name || claims.name || email,
      tenantId: claims.tid || "",
      idToken: response.idToken
    };
  }

  window.arMicrosoftSignIn = arMicrosoftSignIn;
  window.arMicrosoftSignInWithEvent = async function () {
    try {
      const result = await arMicrosoftSignIn();
      window.dispatchEvent(new CustomEvent("arMicrosoftAuthResult", {
        detail: JSON.stringify({ ok: true, result: JSON.stringify(result) })
      }));
    } catch (error) {
      window.dispatchEvent(new CustomEvent("arMicrosoftAuthResult", {
        detail: JSON.stringify({ ok: false, error: error && error.message ? error.message : String(error) })
      }));
    }
  };

  window.addEventListener("arMicrosoftAuthStart", window.arMicrosoftSignInWithEvent);
})();
