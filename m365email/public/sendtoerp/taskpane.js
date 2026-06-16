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

const state = {
	csrfToken: null,
	attachments: [],   // {id, name, selected}
	target: null,      // {doctype, value, label}
};


const auth = {
	async loadSession() {
		const resp = await fetch(`${API_BASE}GetSession`, { credentials: "include" });
		if (!resp.ok) return null;
		const data = (await resp.json()).message;
		if (!data || data.user === "Guest") return null;
		state.csrfToken = data.csrf_token;
		return data.user;
	},

	async login(usr, pwd) {
		const body = new URLSearchParams({ usr, pwd });
		const resp = await fetch("/api/method/login", {
			method: "POST",
			credentials: "include",
			headers: { "Content-Type": "application/x-www-form-urlencoded" },
			body,
		});
		if (!resp.ok) throw new Error("Invalid email or password.");
		return auth.loadSession();
	},

	headers(extra) {
		return Object.assign({ "X-Frappe-CSRF-Token": state.csrfToken || "" }, extra || {});
	},
};


async function apiGet(method, params) {
	const qs = params ? "?" + new URLSearchParams(params).toString() : "";
	const resp = await fetch(`${API_BASE}${method}${qs}`, { credentials: "include" });
	if (!resp.ok) throw new Error(`${method} failed (${resp.status})`);
	return (await resp.json()).message;
}


async function apiPost(method, params) {
	const resp = await fetch(`${API_BASE}${method}`, {
		method: "POST",
		credentials: "include",
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


function wireEvents() {
	$("loginBtn").addEventListener("click", async () => {
		$("loginError").textContent = "";
		try {
			const user = await auth.login($("usr").value.trim(), $("pwd").value);
			if (!user) throw new Error("Sign-in failed.");
			await enterFilingView(user);
		} catch (e) {
			$("loginError").textContent = e.message;
		}
	});
	$("search").addEventListener("input", onSearchInput);
	$("doctype").addEventListener("change", () => { $("results").innerHTML = ""; });
	$("saveEmail").addEventListener("change", refreshFileButton);
	$("fileBtn").addEventListener("click", onFile);
}


Office.onReady(async () => {
	wireEvents();
	const user = await auth.loadSession();
	if (user) await enterFilingView(user);
	else show("loginView");
});
