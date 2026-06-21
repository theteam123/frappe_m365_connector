/*
 * Send to ERP — Outlook add-in task pane.
 *
 * Reads the open email via Office.js, lets the user pick attachments and a
 * target ERP record, then files them through the m365email send_to_erp API.
 *
 * The panel is served same-origin with ERP, so API calls are relative. Auth is
 * isolated in the `auth` object below: it currently uses the login-session +
 * CSRF approach. If browser third-party-cookie blocking prevents the session
 * cookie from surviving inside the Outlook frame, swap `auth` for a token-based
 * implementation without touching the rest of this file.
 */

const API_BASE = "/api/method/m365email.m365email.send_to_erp.";
const SEARCH_DEBOUNCE_MS = 250;
const MIN_SEARCH_CHARS = 1;
const TASK_TYPE = "ToDo";  // the "Task / To-Do" target searches people, not records

// Microsoft 365 sign-in (Nested App Authentication). These are public OAuth
// identifiers for the existing m365email Azure app — not secrets.
const MSAL_CLIENT_ID = "d2a895aa-c239-4605-ac8e-d844341887fc";
const MSAL_TENANT_ID = "85e4e8f4-774e-4a96-987f-bc1a3813d984";
const MSAL_SCOPES = ["User.Read"];
const TOKEN_REFRESH_SKEW_MS = 120000;          // re-acquire a token this long before it expires
const SIGNIN_DIALOG_SIZE = { height: 65, width: 40 };  // % of screen, for the web sign-in dialog

// Environment badge. The panel is served same-origin with its ERP site, so the
// host reliably identifies whether the user is filing to Production or Staging.
const ENV_LABELS = [
	{ match: "ops.sgcloud", label: "Production", className: "env--prod" },
	{ match: "staging.sgcapp", label: "Staging", className: "env--staging" },
];

const state = {
	msToken: null,     // Microsoft 365 ID token (identity sent to ERP)
	attachments: [],   // {id, name, selected}
	target: null,      // {doctype, value, label}
};


// Microsoft 365 sign-on. On desktop Outlook the Office host brokers a token
// silently via Nested App Authentication (NAA). In Outlook on the web the add-in
// runs in an iframe where the Microsoft sign-in page is blocked, so NAA can't
// complete; there we fall back to the Office dialog API (a standalone window that
// can show the sign-in page) and receive the ID token via messageParent. The ID
// token is sent to ERP as a bearer token, which ERP validates and maps to a user.
const auth = {
	pca: null,
	account: null,

	_isWeb() {
		return Office.context.platform === Office.PlatformType.OfficeOnline;
	},

	// Per-request token: reuse the cached token until it nears expiry. Desktop can
	// refresh silently; the web has no silent path, so it must sign in again.
	async ensureToken() {
		if (state.msToken && !isTokenExpired(state.msToken)) return state.msToken;
		if (auth._isWeb()) throw new Error("Your sign-in has expired — please sign in again.");
		return auth._acquireViaNaa();
	},

	// Interactive sign-in, triggered by a user gesture. Web uses the dialog;
	// desktop uses NAA and falls back to the dialog if the broker is unavailable.
	async signIn() {
		if (auth._isWeb()) return auth._acquireViaDialog();
		try {
			return await auth._acquireViaNaa();
		} catch (e) {
			return auth._acquireViaDialog();
		}
	},

	async _getApp() {
		if (auth.pca) return auth.pca;
		auth.pca = await msal.createNestablePublicClientApplication({
			auth: {
				clientId: MSAL_CLIENT_ID,
				authority: `https://login.microsoftonline.com/${MSAL_TENANT_ID}`,
				supportsNestedAppAuth: true,
			},
			cache: { cacheLocation: "localStorage" },
		});
		return auth.pca;
	},

	async _acquireViaNaa() {
		const pca = await auth._getApp();
		const request = { scopes: MSAL_SCOPES };
		if (auth.account) request.account = auth.account;
		let result;
		try {
			result = await pca.acquireTokenSilent(request);
		} catch (e) {
			result = await pca.acquireTokenPopup(request);
		}
		auth.account = result.account;
		state.msToken = result.idToken;
		return result.idToken;
	},

	// Open the sign-in dialog (a standalone window) and resolve with its ID token.
	_acquireViaDialog() {
		return new Promise((resolve, reject) => {
			const dialogUrl = new URL("dialog.html", window.location.href).href;
			Office.context.ui.displayDialogAsync(dialogUrl, SIGNIN_DIALOG_SIZE, (openResult) => {
				if (openResult.status !== Office.AsyncResultStatus.Succeeded) {
					reject(new Error("Could not open the sign-in window."));
					return;
				}
				const dialog = openResult.value;
				dialog.addEventHandler(Office.EventType.DialogMessageReceived,
					(arg) => handleDialogToken(arg, dialog, resolve, reject));
				dialog.addEventHandler(Office.EventType.DialogEventReceived,
					() => reject(new Error("Sign-in was cancelled.")));
			});
		});
	},
};


// True if the JWT is missing, unreadable, or within the refresh skew of expiry.
function isTokenExpired(token) {
	try {
		let base64 = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
		while (base64.length % 4) base64 += "=";
		const claims = JSON.parse(atob(base64));
		return Date.now() >= (claims.exp * 1000) - TOKEN_REFRESH_SKEW_MS;
	} catch (e) {
		return true;
	}
}


// Handle the message posted back by the sign-in dialog: resolve with the ID
// token, or reject with the reported error.
function handleDialogToken(arg, dialog, resolve, reject) {
	let payload = {};
	try {
		payload = JSON.parse(arg.message);
	} catch (e) { /* leave payload empty -> treated as failure */ }
	dialog.close();
	if (payload.status === "ok" && payload.idToken) {
		state.msToken = payload.idToken;
		resolve(payload.idToken);
	} else {
		reject(new Error(payload.error || "Sign-in failed."));
	}
}


async function apiGet(method, params) {
	const token = await auth.ensureToken();
	const qs = params ? "?" + new URLSearchParams(params).toString() : "";
	const resp = await fetch(`${API_BASE}${method}${qs}`, { headers: { "Authorization": `Bearer ${token}` } });
	if (!resp.ok) throw new Error(`${method} failed (${resp.status})`);
	return (await resp.json()).message;
}


async function apiPost(method, params) {
	const token = await auth.ensureToken();
	const resp = await fetch(`${API_BASE}${method}`, {
		method: "POST",
		headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/x-www-form-urlencoded" },
		body: new URLSearchParams(params),
	});
	if (!resp.ok) {
		const text = await resp.text();
		throw new Error(extractError(text) || `${method} failed (${resp.status})`);
	}
	return (await resp.json()).message;
}


function extractError(text) {
	try {
		const messages = JSON.parse(JSON.parse(text)._server_messages || "[]");
		if (messages.length) return JSON.parse(messages[0]).message;
	} catch (e) { /* fall through */ }
	return null;
}


function $(id) { return document.getElementById(id); }
function show(id) { $(id).classList.remove("hidden"); }
function hide(id) { $(id).classList.add("hidden"); }


function applyEnvBadges() {
	const host = window.location.hostname;
	const known = ENV_LABELS.find(envEntry => host.includes(envEntry.match));
	const label = known ? known.label : host;
	const className = known ? known.className : "env--unknown";
	document.querySelectorAll(".env").forEach(badgeElement => {
		badgeElement.textContent = label;
		badgeElement.classList.add(className);
	});
}


/* ---- Office.js item reading ---- */

function readAttachments() {
	const item = Office.context.mailbox.item;
	const files = (item.attachments || []).filter(a => a.attachmentType === "file" && !a.isInline);
	state.attachments = files.map(a => ({ id: a.id, name: a.name, selected: true }));
}


function getEmailBodyHtml() {
	return new Promise((resolve) => {
		Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (r) => {
			resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "");
		});
	});
}


function getAttachmentContent(attachmentId) {
	return new Promise((resolve, reject) => {
		Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (r) => {
			if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
			else reject(new Error("Could not read attachment."));
		});
	});
}


function emailMeta() {
	const item = Office.context.mailbox.item;
	const from = item.from || item.sender || {};
	const to = (item.to || []).map(r => r.emailAddress).join(", ");
	return {
		subject: item.subject || "",
		sender: from.emailAddress || "",
		recipients: to,
		sent_date: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : "",
	};
}


/* ---- UI wiring ---- */

function renderAttachments() {
	const container = $("attachmentList");
	container.innerHTML = "";
	if (!state.attachments.length) { show("noAttachments"); return; }
	hide("noAttachments");
	state.attachments.forEach((a, i) => {
		const label = document.createElement("label");
		label.innerHTML = `<input type="checkbox" data-i="${i}" checked /> ${escapeHtml(a.name)}`;
		label.querySelector("input").addEventListener("change", (e) => {
			state.attachments[i].selected = e.target.checked;
			refreshFileButton();
		});
		container.appendChild(label);
	});
}


function escapeHtml(s) {
	const d = document.createElement("div");
	d.textContent = s;
	return d.innerHTML;
}


async function populateDoctypes() {
	const types = await apiGet("GetPickableDoctypes");
	const select = $("doctype");
	select.innerHTML = "";
	(types || []).forEach(t => {
		const opt = document.createElement("option");
		opt.value = t.doctype;
		opt.textContent = t.label;
		select.appendChild(opt);
	});
	updateModeForType();
}


// Switch the picker between filing to a record and assigning a Task/To-Do.
function updateModeForType() {
	const isTask = $("doctype").value === TASK_TYPE;
	$("findLabel").textContent = isTask ? "Assign to" : "Find record";
	$("search").placeholder = isTask ? "Search people…" : "Start typing…";
	$("fileBtn").textContent = isTask ? "Create task" : "Attach to ERP";
	$("search").value = "";
	$("results").innerHTML = "";
	hide("chosen");
	state.target = null;
	refreshFileButton();
}


let searchTimer = null;
function onSearchInput() {
	clearTimeout(searchTimer);
	const txt = $("search").value.trim();
	if (txt.length < MIN_SEARCH_CHARS) { $("results").innerHTML = ""; return; }
	searchTimer = setTimeout(() => runSearch(txt), SEARCH_DEBOUNCE_MS);
}


function setResultsMessage(message) {
	const list = $("results");
	list.innerHTML = "";
	const li = document.createElement("li");
	li.textContent = message;
	li.className = "result-message";
	list.appendChild(li);
}


async function runSearch(txt) {
	const doctype = $("doctype").value;
	setResultsMessage("Searching…");
	try {
		const rows = await apiGet("SearchTargets", { target_doctype: doctype, txt });
		const list = $("results");
		list.innerHTML = "";
		if (!rows || !rows.length) {
			setResultsMessage("No matches found.");
			return;
		}
		rows.forEach(row => {
			const li = document.createElement("li");
			const main = document.createElement("div");
			main.textContent = row.label;
			li.appendChild(main);
			if (row.sublabel) {
				const sub = document.createElement("div");
				sub.className = "result-sub";
				sub.textContent = row.sublabel;
				li.appendChild(sub);
			}
			li.addEventListener("click", () => chooseTarget(doctype, row));
			list.appendChild(li);
		});
	} catch (e) {
		setResultsMessage("Search error: " + (e.message || e));
	}
}


function chooseTarget(doctype, row) {
	state.target = { doctype, value: row.value, label: row.label };
	$("results").innerHTML = "";
	$("search").value = "";
	$("chosen").textContent = `→ ${row.label}${row.sublabel ? " · " + row.sublabel : ""}`;
	show("chosen");
	refreshFileButton();
}


function refreshFileButton() {
	const hasAttachmentsSelected = state.attachments.some(a => a.selected);
	const wantsEmail = $("saveEmail").checked;
	const isTask = $("doctype").value === TASK_TYPE;
	// For a task, choosing the assignee is enough; for a record, also require the
	// email, an attachment, or a comment.
	const hasContent = isTask || hasAttachmentsSelected || wantsEmail || $("comment").value.trim() !== "";
	$("fileBtn").disabled = !state.target || !hasContent;
}


async function onFile() {
	const btn = $("fileBtn");
	btn.disabled = true;
	setStatus("Filing…", "");
	try {
		const selected = state.attachments.filter(a => a.selected);
		const payload = [];
		for (const a of selected) {
			const content = await getAttachmentContent(a.id);
			payload.push({ file_name: a.name, content_base64: content.content, content_type: content.format });
		}
		const meta = emailMeta();
		const result = await apiPost("FileEmailToRecord", Object.assign({
			target_doctype: state.target.doctype,
			target_name: state.target.value,
			save_email: $("saveEmail").checked ? 1 : 0,
			comment: $("comment").value,
			body_html: await getEmailBodyHtml(),
			attachments: JSON.stringify(payload),
		}, meta));
		setStatus(result.message || "Done.", "ok");
	} catch (e) {
		setStatus(e.message || "Something went wrong.", "err");
	} finally {
		btn.disabled = false;
	}
}


function setStatus(text, kind) {
	const el = $("status");
	el.textContent = text;
	el.className = "status" + (kind ? " " + kind : "");
}


async function enterFilingView(user) {
	hide("loginView");
	show("fileView");
	$("whoami").textContent = user;
	readAttachments();
	renderAttachments();
	await populateDoctypes();
	refreshFileButton();
}


async function doSignIn() {
	$("loginError").textContent = "";
	$("loginStatus").textContent = "Signing in with Microsoft 365…";
	hide("loginBtn");
	try {
		await auth.signIn();
		const session = await apiGet("GetSession");
		if (!session || session.user === "Guest") {
			throw new Error("Your Microsoft account isn't linked to an ERP user. Contact your administrator.");
		}
		await enterFilingView(session.user);
	} catch (e) {
		$("loginStatus").textContent = "Couldn't sign in automatically.";
		$("loginError").textContent = e.message || "Sign-in failed.";
		show("loginBtn");
	}
}


function wireEvents() {
	$("loginBtn").addEventListener("click", doSignIn);
	$("search").addEventListener("input", onSearchInput);
	$("doctype").addEventListener("change", updateModeForType);
	$("saveEmail").addEventListener("change", refreshFileButton);
	$("comment").addEventListener("input", refreshFileButton);
	$("fileBtn").addEventListener("click", onFile);
}


// Render the environment badge immediately — it is pure DOM and must not depend
// on the Office host initializing (so it shows even if Office.js is slow).
applyEnvBadges();


Office.onReady(async () => {
	wireEvents();
	show("loginView");
	// The web sign-in dialog must be opened by a user gesture, so on the web we
	// wait for the button. Desktop can start sign-in silently on load.
	if (auth._isWeb()) {
		$("loginStatus").textContent = "Sign in to continue.";
		show("loginBtn");
	} else {
		await doSignIn();
	}
});
