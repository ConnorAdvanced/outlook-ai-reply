/* AI Reply Drafter — Outlook task pane logic.
 * Reads the currently open message, sends it to Claude, and shows a draft reply
 * that the user can edit and open in Outlook's Reply / Reply All window.
 * Nothing is ever sent without the user pressing Send in Outlook.
 */
(function () {
  "use strict";

  const MAX_EMAIL_CHARS = 30000; // keep long threads within a sensible size
  const STORE_KEY = "aiReplyDrafter.settings";
  const $ = (id) => document.getElementById(id);

  let currentEmail = null;   // { subject, fromName, fromEmail, body, itemId }
  let requestCounter = 0;    // ignore stale responses when the user switches emails

  // ---------- settings ----------
  function loadSettings() {
    try { return JSON.parse(localStorage.getItem(STORE_KEY)) || {}; } catch (e) { return {}; }
  }
  function saveSettings(s) {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(s)); } catch (e) { /* storage blocked */ }
  }
  function readSettingsForm() {
    return {
      apiKey: $("apiKey").value.trim(),
      model: $("model").value,
      instructions: $("instructions").value.trim(),
      signature: $("signature").value.trim(),
      autoDraft: $("autoDraft").checked,
      tone: $("tone").value,
      length: $("length").value,
    };
  }
  function fillSettingsForm(s) {
    $("apiKey").value = s.apiKey || "";
    if (s.model) $("model").value = s.model;
    $("instructions").value = s.instructions || "";
    $("signature").value = s.signature || "";
    $("autoDraft").checked = s.autoDraft !== false;
    if (s.tone) $("tone").value = s.tone;
    if (s.length) $("length").value = s.length;
  }

  // ---------- UI helpers ----------
  function setStatus(msg, isError) {
    const el = $("status");
    el.textContent = msg || "";
    el.classList.toggle("error", !!isError);
  }
  function setDraftButtons(enabled) {
    ["reply", "replyAll", "copy"].forEach((id) => ($(id).disabled = !enabled));
  }
  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }
  function textToHtml(text) {
    return text.split(/\n{2,}/)
      .map((p) => "<p>" + escapeHtml(p).replace(/\n/g, "<br>") + "</p>")
      .join("");
  }

  // ---------- reading the email ----------
  function readCurrentEmail() {
    return new Promise((resolve, reject) => {
      const item = Office.context.mailbox.item;
      if (!item) return resolve(null);
      item.body.getAsync(Office.CoercionType.Text, (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(new Error("Couldn't read the email body: " + res.error.message));
        }
        const from = item.from || item.sender || {};
        resolve({
          subject: item.subject || "(no subject)",
          fromName: from.displayName || "",
          fromEmail: from.emailAddress || "",
          to: (item.to || []).map((r) => r.displayName || r.emailAddress).join(", "),
          body: (res.value || "").trim().slice(0, MAX_EMAIL_CHARS),
          itemId: item.itemId,
        });
      });
    });
  }

  async function onItemChanged() {
    $("draft").value = "";
    setDraftButtons(false);
    setStatus("");
    $("guidance").value = "";
    try {
      currentEmail = await readCurrentEmail();
    } catch (e) {
      currentEmail = null;
      setStatus(e.message, true);
      return;
    }
    if (!currentEmail) {
      $("emailMeta").textContent = "Open an email to get started.";
      return;
    }
    $("emailMeta").innerHTML =
      "<strong>" + escapeHtml(currentEmail.subject) + "</strong><br>From: " +
      escapeHtml(currentEmail.fromName || currentEmail.fromEmail);

    const s = loadSettings();
    if (s.autoDraft !== false && s.apiKey) generate();
  }

  // ---------- calling Claude ----------
  function buildPrompt(email, s, guidance) {
    const me = Office.context.mailbox.userProfile;
    const system = [
      "You draft email replies on behalf of " + (me.displayName || "the user") +
        " (" + (me.emailAddress || "") + ").",
      "Write ONLY the body of the reply: no subject line, no preamble like 'Here is a draft', no commentary.",
      "Reply to the most recent message in the thread; earlier quoted messages are context only.",
      "Match the language the sender wrote in.",
      "Never invent facts, commitments, prices, dates or attachments that aren't supported by the email or the user's guidance. " +
        "Where the reply needs information you don't have, put a clear placeholder in square brackets, e.g. [confirm date].",
      "Tone: " + s.tone + ". Length: " + s.length + ".",
      s.instructions ? "About the user and how they write:\n" + s.instructions : "",
      s.signature
        ? "End the reply with exactly this sign-off:\n" + s.signature
        : "End with a short sign-off using the user's first name.",
    ].filter(Boolean).join("\n\n");

    const user =
      "Here is the email I received:\n\n" +
      "<email>\nFrom: " + email.fromName + " <" + email.fromEmail + ">\n" +
      "To: " + email.to + "\nSubject: " + email.subject + "\n\n" + email.body + "\n</email>\n\n" +
      (guidance ? "My guidance for this reply: " + guidance + "\n\n" : "") +
      "Draft my reply.";
    return { system, user };
  }

  async function callClaude(s, system, user) {
    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": s.apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({
        model: s.model || "claude-sonnet-5-5",
        max_tokens: 1500,
        system,
        messages: [{ role: "user", content: user }],
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      const msg = (data.error && data.error.message) || resp.statusText;
      if (resp.status === 401) throw new Error("API key was rejected. Check it in Settings.");
      throw new Error("Claude API error (" + resp.status + "): " + msg);
    }
    return (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("").trim();
  }

  async function generate() {
    const s = Object.assign(loadSettings(), { tone: $("tone").value, length: $("length").value });
    if (!s.apiKey) {
      $("settings").classList.remove("hidden");
      setStatus("Add your Anthropic API key in Settings first.", true);
      return;
    }
    if (!currentEmail) { setStatus("Open an email first.", true); return; }

    const myRequest = ++requestCounter;
    const emailAtStart = currentEmail.itemId;
    $("generate").disabled = true;
    setDraftButtons(false);
    setStatus("Reading the email and drafting a reply…");
    try {
      const { system, user } = buildPrompt(currentEmail, s, $("guidance").value.trim());
      const text = await callClaude(s, system, user);
      if (myRequest !== requestCounter || !currentEmail || currentEmail.itemId !== emailAtStart) return;
      $("draft").value = text;
      setDraftButtons(!!text);
      setStatus(text ? "Draft ready — edit if needed, then click Reply." : "Claude returned an empty draft. Try again.", !text);
    } catch (e) {
      if (myRequest === requestCounter) setStatus(e.message || String(e), true);
    } finally {
      if (myRequest === requestCounter) $("generate").disabled = false;
    }
  }

  // ---------- inserting the reply ----------
  function openReply(all) {
    const text = $("draft").value.trim();
    if (!text) return;
    const item = Office.context.mailbox.item;
    const html = textToHtml(text);
    try {
      if (all) item.displayReplyAllForm(html);
      else item.displayReplyForm(html);
    } catch (e) {
      setStatus("Couldn't open the reply window: " + e.message, true);
    }
  }

  async function copyDraft() {
    const text = $("draft").value;
    try {
      await navigator.clipboard.writeText(text);
    } catch (e) {
      $("draft").select();
      document.execCommand("copy");
    }
    setStatus("Copied to clipboard.");
  }

  // ---------- startup ----------
  Office.onReady(() => {
    const s = loadSettings();
    fillSettingsForm(s);
    if (!s.apiKey) $("settings").classList.remove("hidden");

    $("settingsToggle").onclick = () => $("settings").classList.toggle("hidden");
    $("saveSettings").onclick = () => {
      saveSettings(readSettingsForm());
      $("settings").classList.add("hidden");
      setStatus("Settings saved.");
    };
    $("tone").onchange = $("length").onchange = () => {
      const cur = loadSettings();
      cur.tone = $("tone").value; cur.length = $("length").value;
      saveSettings(cur);
    };
    $("generate").onclick = generate;
    $("guidance").addEventListener("keydown", (e) => { if (e.key === "Enter") generate(); });
    $("reply").onclick = () => openReply(false);
    $("replyAll").onclick = () => openReply(true);
    $("copy").onclick = copyDraft;

    // When the pane is pinned, follow the user as they click through emails.
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
    }
    onItemChanged();
  });
})();
