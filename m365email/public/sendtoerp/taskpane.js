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

// Microsoft 365 sign-in (Nested App Authentication). These are public OAuth
// identifiers for the existing m365email Azure app — not secrets.
const MSAL_CLIENT_ID = "d2a895aa-c239-4605-ac8e-d844341887fc";
const MSAL_TENANT_ID = "85e4e8f4-774e-4a96-987f-bc1a3813d984";
const MSAL_SCOPES = ["User.Read"];

const state = {
	msToken: null,     // Microsoft 365 ID token (identity sent to ERP)
	attachments: [],   // {id, name, selected}
	target: null,      // {doctype, value, label}
};


// Microsoft 365 sign-on via Nested App Authentication. The Office host brokers
// a token silently; the ID token is sent to ERP as a bearer token, which ERP
// validates and maps to a Frappe user. No passwords, no manifest SSO wiring.
const auth = {
	pca: null,

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

	async signIn() {
		const pca = await auth._getApp();
		const request = { scopes: MSAL_SCOPES };
		let result;
		try {
			result = await pca.acquireTokenSilent(request);
		} catch (e) {
			result = await pca.acquireTokenPopup(request);
		}
		state.msToken = result.idToken;
		return state.msToken;
	},

	headers(extra) {
		return Object.assign({ "Authorization": `Bearer ${state.msToken || ""}` }, extra || {});
	},
};


async function apiGet(method, params) {
	const qs = params ? "?" + new URLSearchParams(params).toString() : "";
	const resp = await fetch(`${API_BASE}${method}${qs}`, { headers: auth.headers() });
	if (!resp.ok) throw new Error(`${method} failed (${resp.status})`);
	return (await resp.json()).message;
}


async function apiPost(method, params) {
	const resp = await fetch(`${API_BASE}${method}`, {
		method: "POST",
		headers: auth.headers({ "Content-Type": "application/x-www-form-urlencoded" }),
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
}


let searchTimer = null;
function onSearchInput() {
	clearTimeout(searchTimer);
	const txt = $("search").value.trim();
	if (txt.length < MIN_SEARCH_CHARS) { $("results").innerHTML = ""; return; }
	searchTimer = setTimeout(() => runSearch(txt), SEARCH_DEBOUNCE_MS);
}


async function runSearch(txt) {
	const doctype = $("doctype").value;
	const rows = await apiGet("SearchTargets", { doctype, txt });
	const list = $("results");
	list.innerHTML = "";
	(rows || []).forEach(row => {
		const li = document.createElement("li");
		li.textContent = row.label;
		li.addEventListener("click", () => chooseTarget(doctype, row));
		list.appendChild(li);
	});
}


function chooseTarget(doctype, row) {
	state.target = { doctype, value: row.value, label: row.label };
	$("results").innerHTML = "";
	$("search").value = "";
	$("chosen").textContent = `→ ${row.label}`;
	show("chosen");
	refreshFileButton();
}


function refreshFileButton() {
	const hasAttachmentsSelected = state.attachments.some(a => a.selected);
	const wantsEmail = $("saveEmail").checked;
	$("fileBtn").disabled = !state.target || (!hasAttachmentsSelected && !wantsEmail);
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
	$("doctype").addEventListener("change", () => { $("results").innerHTML = ""; });
	$("saveEmail").addEventListener("change", refreshFileButton);
	$("fileBtn").addEventListener("click", onFile);
}


Office.onReady(async () => {
	wireEvents();
	show("loginView");
	await doSignIn();
});
