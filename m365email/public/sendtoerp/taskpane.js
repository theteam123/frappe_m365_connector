/*
 * Send to ERP — Outlook add-in task pane.
 *
 * Reads the open email via Office.js, lets the user pick attachments and a
 * target ERP record, then files them through the m365email send_to_erp API.
 *
 * The panel is served same-origin with ERP, so API calls are relative. Auth is
 * isolated in the `auth` object below and is token based (bearer header), never
 * the ERP session cookie, which browsers block inside the Outlook frame.
 *
 * Staying signed in: a Microsoft identity token lasts about an hour and, in
 * Outlook on the web, can only be renewed through a sign-in window. So after a
 * successful Microsoft sign-in the panel asks ERP for its own remembered
 * session token (see m365email addin_session.py), keeps it in this panel's
 * localStorage, and reuses it on every later open. ERP expires it after a
 * period of no use, and "Sign out" revokes it.
 */

const API_BASE = "/api/method/m365email.m365email.send_to_erp.";
const SESSION_API_BASE = "/api/method/m365email.m365email.addin_session.";
const PROCUREMENT_API_BASE = "/api/method/m365email.m365email.send_to_erp_procurement.";
const PEOPLE_API_BASE = "/api/method/m365email.m365email.send_to_erp_people.";
const SESSION_STORAGE_KEY = "sendtoerp.session"; // {token, user, expires_on} in localStorage
const AUTH_FAILURE_STATUSES = [401, 403]; // ERP rejected the bearer token
const SEARCH_DEBOUNCE_MS = 250;
const MIN_SEARCH_CHARS = 1;
const TASK_TYPE = "ToDo"; // the "Task / To-Do" target searches people, not records
const TICKET_TYPE = "Support Ticket"; // the "Create support ticket" target lists ticket types
const QUOTE_TYPE = "Supplier Quote"; // "create new" raises a quote for a job (Paul, 2026-11-06)
const ORDER_TYPE = "Purchase Order"; // the PO is suggested from the email; can enter an invoice
// Targets that offer "create new" from the email's sender, with the word used in the panel.
const CREATABLE_TYPES = {
	Contact: "contact",
	Candidate: "candidate",
	[QUOTE_TYPE]: "supplier quote",
};
const PROJECT_TYPE = "Project"; // a new Supplier Quote's job is picked from Projects
const CONTACT_TYPE = "Contact";
// Types that ARE a person (or a task), so they get no separate contact picker (Paul,
// 2026-11-07: every other option lets the user say which Contact the email is about).
const NO_CONTACT_PICKER_TYPES = ["Contact", "Candidate", "Employee", "ToDo"];
const DATE_PART_WIDTH = 2; // month and day in YYYY-MM-DD
const PDF_NAME_PATTERN = /\.pdf$/i;
const INVOICE_WORD_PATTERN = /invoice/i;

// Microsoft 365 sign-in (Nested App Authentication). These are public OAuth
// identifiers for the existing m365email Azure app — not secrets.
const MSAL_CLIENT_ID = "d2a895aa-c239-4605-ac8e-d844341887fc";
const MSAL_TENANT_ID = "85e4e8f4-774e-4a96-987f-bc1a3813d984";
const MSAL_SCOPES = ["User.Read"];
const TOKEN_REFRESH_SKEW_MS = 120000; // re-acquire a token this long before it expires
const SIGNIN_DIALOG_SIZE = { height: 65, width: 40 }; // % of screen, for the web sign-in dialog

// Environment badge. The panel is served same-origin with its ERP site, so the
// host reliably identifies whether the user is filing to Production or Staging.
const ENV_LABELS = [
	{ match: "ops.sgcloud", label: "Production", className: "env--prod" },
	{ match: "staging.sgcapp", label: "Staging", className: "env--staging" },
];

const state = {
	msToken: null, // Microsoft 365 ID token (identity sent to ERP at sign-in)
	erpSession: null, // {token, user, expires_on} — ERP's remembered sign-in, reused between opens
	attachments: [], // {id, name, selected}
	target: null, // {doctype, value, label, create}
	project: null, // {value, label} — the job for a new Supplier Quote
	supplier: null, // {value, label} — the supplier for a new Supplier Quote, when picked or known
	suggestionRun: 0, // increments per PO suggestion, so a stale reply is ignored
	senderName: "", // the person the email is from: a forwarded email's original sender
	senderEmail: "", // that person's address
	person: null, // {name, email, is_forwarded, is_internal} — the server's best guess
	contact: null, // {value, label, sublabel} — the Contact this email is about, editable
	guessedSupplier: null, // {value, label} — the supplier the person works for, if known
	peopleRun: 0, // increments per supplier re-guess, so a stale reply is ignored
	isDone: false, // the current choice has been filed: the button stays greyed until it changes
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

	// Per-request token. ERP's remembered session wins when we have one; before
	// that, reuse the Microsoft token until it nears expiry. Desktop can refresh
	// it silently; the web has no silent path, so it must sign in again.
	async ensureToken() {
		if (state.erpSession) return state.erpSession.token;
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
				dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) =>
					handleDialogToken(arg, dialog, resolve, reject)
				);
				dialog.addEventHandler(Office.EventType.DialogEventReceived, () =>
					reject(new Error("Sign-in was cancelled."))
				);
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
		return Date.now() >= claims.exp * 1000 - TOKEN_REFRESH_SKEW_MS;
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
	} catch (e) {
		/* leave payload empty -> treated as failure */
	}
	dialog.close();
	if (payload.status === "ok" && payload.idToken) {
		state.msToken = payload.idToken;
		resolve(payload.idToken);
	} else {
		reject(new Error(payload.error || "Sign-in failed."));
	}
}

// Methods live under send_to_erp unless given as "session:", "procurement:" or "people:".
function apiUrl(method) {
	const prefixes = {
		"session:": SESSION_API_BASE,
		"procurement:": PROCUREMENT_API_BASE,
		"people:": PEOPLE_API_BASE,
	};
	const prefix = Object.keys(prefixes).find((candidate) => method.startsWith(candidate));
	if (!prefix) return API_BASE + method;
	return prefixes[prefix] + method.slice(prefix.length);
}

// A rejected bearer while using a remembered session means ERP no longer
// accepts it (expired, revoked, or user disabled): forget it so the next open
// signs in with Microsoft again instead of failing forever.
function noteAuthFailure(resp) {
	if (!state.erpSession || !AUTH_FAILURE_STATUSES.includes(resp.status)) return;
	clearStoredSession();
	state.erpSession = null;
}

async function apiGet(method, params) {
	const token = await auth.ensureToken();
	const qs = params ? "?" + new URLSearchParams(params).toString() : "";
	const headers = { Authorization: `Bearer ${token}` };
	const resp = await fetch(`${apiUrl(method)}${qs}`, { headers });
	if (!resp.ok) {
		noteAuthFailure(resp);
		throw new Error(`${method} failed (${resp.status})`);
	}
	return (await resp.json()).message;
}

async function apiPost(method, params) {
	const token = await auth.ensureToken();
	const resp = await fetch(apiUrl(method), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/x-www-form-urlencoded",
		},
		body: new URLSearchParams(params || {}),
	});
	if (!resp.ok) {
		noteAuthFailure(resp);
		const text = await resp.text();
		throw new Error(extractError(text) || `${method} failed (${resp.status})`);
	}
	return (await resp.json()).message;
}

/* ---- Remembered ERP session (survives closing and reopening the panel) ---- */

function loadStoredSession() {
	try {
		const raw = window.localStorage.getItem(SESSION_STORAGE_KEY);
		if (!raw) return null;
		const stored = JSON.parse(raw);
		if (!stored || !stored.token) return null;
		if (stored.expires_on && new Date(stored.expires_on) <= new Date()) return null;
		return stored;
	} catch (e) {
		return null; // storage unavailable (private browsing etc.) — just sign in again
	}
}

function saveStoredSession(session) {
	try {
		window.localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(session));
	} catch (e) {
		/* storage unavailable — the session still works for this open */
	}
}

function clearStoredSession() {
	try {
		window.localStorage.removeItem(SESSION_STORAGE_KEY);
	} catch (e) {
		/* nothing to clear */
	}
}

function describePlatform() {
	const diagnostics = (Office.context && Office.context.diagnostics) || {};
	return [diagnostics.host, Office.context.platform, diagnostics.version]
		.filter(Boolean)
		.join(" ");
}

// After a Microsoft sign-in, ask ERP for a remembered session so later opens
// skip the sign-in. Failure is not fatal: the Microsoft token still works for
// this open, it just will not be remembered.
async function establishErpSession() {
	try {
		const session = await apiPost("session:CreateAddinSession", {
			platform: describePlatform(),
		});
		if (!session || !session.token) return;
		state.erpSession = {
			token: session.token,
			user: session.user,
			expires_on: session.expires_on,
		};
		saveStoredSession(state.erpSession);
	} catch (e) {
		console.warn("Send to ERP: could not create a remembered session", e);
	}
}

// Try the remembered session before asking anyone to sign in. Returns the ERP
// user on success, or null (with the stale session forgotten) otherwise.
async function resumeStoredSession() {
	const stored = loadStoredSession();
	if (!stored) return null;
	state.erpSession = stored;
	try {
		const session = await apiGet("GetSession");
		if (session && session.user && session.user !== "Guest") return session.user;
	} catch (e) {
		/* fall through: treat as not signed in */
	}
	clearStoredSession();
	state.erpSession = null;
	return null;
}

async function signOut() {
	try {
		if (state.erpSession) await apiPost("session:RevokeAddinSession");
	} catch (e) {
		/* revoking is best-effort; the local copy is dropped regardless */
	}
	clearStoredSession();
	state.erpSession = null;
	state.msToken = null;
	hide("fileView");
	show("loginView");
	$("loginError").textContent = "";
	$("loginStatus").textContent = "Signed out. Sign in to continue.";
	show("loginBtn");
}

function extractError(text) {
	try {
		const messages = JSON.parse(JSON.parse(text)._server_messages || "[]");
		if (messages.length) return JSON.parse(messages[0]).message;
	} catch (e) {
		/* fall through */
	}
	return null;
}

function $(id) {
	return document.getElementById(id);
}
function show(id) {
	$(id).classList.remove("hidden");
}
function hide(id) {
	$(id).classList.add("hidden");
}

function applyEnvBadges() {
	const host = window.location.hostname;
	const known = ENV_LABELS.find((envEntry) => host.includes(envEntry.match));
	const label = known ? known.label : host;
	const className = known ? known.className : "env--unknown";
	document.querySelectorAll(".env").forEach((badgeElement) => {
		badgeElement.textContent = label;
		badgeElement.classList.add(className);
	});
}

/* ---- Office.js item reading ---- */

function readAttachments() {
	const item = Office.context.mailbox.item;
	const files = (item.attachments || []).filter(
		(a) => a.attachmentType === "file" && !a.isInline
	);
	state.attachments = files.map((a) => ({ id: a.id, name: a.name, selected: true }));
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
	const to = (item.to || []).map((r) => r.emailAddress).join(", ");
	return {
		subject: item.subject || "",
		sender: from.emailAddress || "",
		recipients: to,
		sent_date: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : "",
	};
}

// The email's Internet message id, which ERP uses so a "create" runs once per email.
function emailMessageId() {
	const item = Office.context.mailbox.item;
	return item.internetMessageId || item.itemId || "";
}

// Read the chosen attachments (or only the PDFs) as {file_name, content_base64, content_type}.
async function readAttachmentPayload(onlyPdfs) {
	const wanted = state.attachments.filter((attachment) =>
		onlyPdfs ? PDF_NAME_PATTERN.test(attachment.name) : attachment.selected
	);
	const payload = [];
	for (const attachment of wanted) {
		const content = await getAttachmentContent(attachment.id);
		payload.push({
			file_name: attachment.name,
			content_base64: content.content,
			content_type: content.format,
		});
	}
	return payload;
}

/* ---- UI wiring ---- */

function renderAttachments() {
	const container = $("attachmentList");
	container.innerHTML = "";
	if (!state.attachments.length) {
		show("noAttachments");
		return;
	}
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
	(types || []).forEach((t) => {
		const opt = document.createElement("option");
		opt.value = t.doctype;
		opt.textContent = t.label;
		select.appendChild(opt);
	});
	updateModeForType();
}

// Picker wording per mode: filing to a record, assigning a task, or raising a ticket.
const MODE_TEXT = {
	[TASK_TYPE]: { find: "Assign to", placeholder: "Search people…", button: "Create task" },
	[TICKET_TYPE]: {
		find: "Ticket type",
		placeholder: "Filter ticket types…",
		button: "Create support ticket",
	},
};
const RECORD_MODE_TEXT = {
	find: "Find record",
	placeholder: "Start typing…",
	button: "Attach to ERP",
};
const INVOICE_BUTTON_TEXT = "Enter Supplier Invoice";
const DONE_BUTTON_TEXT = "✓ Done";
const QUOTE_BUTTON_TEXT = "Create supplier quote";

// Switch the picker between filing to a record, assigning a Task/To-Do, and
// raising a support ticket.
function updateModeForType() {
	clearDone();
	const doctype = $("doctype").value;
	const modeText = MODE_TEXT[doctype] || RECORD_MODE_TEXT;
	$("findLabel").textContent = modeText.find;
	$("search").placeholder = modeText.placeholder;
	$("fileBtn").textContent = modeText.button;
	$("search").value = "";
	$("results").innerHTML = "";
	hide("chosen");
	hide("suggestion");
	hide("quoteFields");
	hide("invoiceFields");
	hide("openRecord");
	hide("personEdit");
	state.target = null;
	state.suggestionRun += 1;
	refreshPersonFields(doctype);

	// Contact, Candidate and Supplier Quote offer creating a brand-new record from the email.
	const createWord = CREATABLE_TYPES[doctype];
	if (createWord) {
		const who = state.senderName || state.senderEmail || "this sender";
		const fromWhom = doctype === QUOTE_TYPE ? " from this email" : `: ${who}`;
		$("createNew").textContent = `➕ Create new ${createWord}${fromWhom}`;
		show("createNew");
	} else {
		hide("createNew");
	}

	// There are only a handful of ticket types, so list them straight away.
	if (doctype === TICKET_TYPE) runSearch("");
	// For Contact, look up the person the email is about straight away.
	if (doctype === CONTACT_TYPE && state.senderEmail) {
		$("search").value = state.senderEmail;
		runSearch(state.senderEmail);
	}
	// A supplier's invoice quotes our PO number: find it and preselect the order.
	if (doctype === ORDER_TYPE) {
		fillInvoiceDetails({}, emailMeta().subject);
		suggestPurchaseOrder();
	}
	refreshFileButton();
}

/* ---- Purchase Order: suggest the order, and enter the email as a supplier invoice ---- */

// Read the PO number from the subject and body; only if that finds nothing, read the PDFs
// too (fetching attachments is the slow part). Preselect a single order, list several.
async function suggestPurchaseOrder() {
	const run = state.suggestionRun;
	const meta = emailMeta();
	setSuggestion("Looking for a PO number in the email…");
	try {
		const params = {
			target_doctype: ORDER_TYPE,
			subject: meta.subject,
			body_html: await getEmailBodyHtml(),
		};
		let suggestion = await apiPost("procurement:SuggestTarget", params);
		const hasPdf = state.attachments.some((attachment) =>
			PDF_NAME_PATTERN.test(attachment.name)
		);
		if (!suggestion.match && !suggestion.choices.length && hasPdf) {
			setSuggestion("Reading the PDF for a PO number…");
			params.attachments = JSON.stringify(await readAttachmentPayload(true));
			suggestion = await apiPost("procurement:SuggestTarget", params);
		}
		if (run !== state.suggestionRun) return; // the user switched type meanwhile
		// Never overwrite an invoice number the user has already typed.
		if (!$("invoiceNumber").value.trim())
			fillInvoiceDetails(suggestion.invoice || {}, meta.subject);
		if (state.target) return; // the user already picked an order themselves
		applySuggestion(suggestion);
	} catch {
		if (run !== state.suggestionRun) return;
		setSuggestion("Couldn't read a PO number from this email. Search for the order below.");
	}
}

function applySuggestion(suggestion) {
	if (suggestion.match) {
		chooseTarget(ORDER_TYPE, suggestion.match);
		setSuggestion(`Found purchase order ${suggestion.match.label} in the email.`);
		return;
	}
	if (suggestion.choices.length) {
		setSuggestion("This email mentions more than one purchase order. Pick the right one:");
		renderResultRows(ORDER_TYPE, suggestion.choices);
		return;
	}
	setSuggestion("No PO number found in the email. Search for the order below.");
}

// Pre-fill what was read from the email; the invoice option starts ticked when the email
// looks like an invoice (an invoice number was found, or "invoice" is in the subject).
function fillInvoiceDetails(invoice, subject) {
	$("invoiceNumber").value = invoice.supplier_invoice_number || "";
	$("invoiceDate").value = invoice.invoice_date || emailDateISO();
	$("dueDate").value = invoice.due_date || todayISO();
	$("enterInvoice").checked =
		Boolean(invoice.supplier_invoice_number) || INVOICE_WORD_PATTERN.test(subject || "");
	refreshInvoiceInputs();
}

function setSuggestion(text) {
	$("suggestion").textContent = text;
	show("suggestion");
}

function isEnteringInvoice() {
	return $("doctype").value === ORDER_TYPE && Boolean(state.target) && $("enterInvoice").checked;
}

function refreshInvoiceInputs() {
	$("invoiceInputs").classList.toggle("hidden", !$("enterInvoice").checked);
	refreshFileButton();
}

function emailDateISO() {
	const created = Office.context.mailbox.item.dateTimeCreated;
	return toLocalISODate(created ? new Date(created) : new Date());
}

function todayISO() {
	return toLocalISODate(new Date());
}

function toLocalISODate(date) {
	const pad = (number) => String(number).padStart(DATE_PART_WIDTH, "0");
	return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

/* ---- Create new Supplier Quote: pick the job, and the supplier when it is not known ---- */

// A small type-ahead picker: search as the user types, list results, remember the choice.
function wirePicker(prefix, search, onChoose) {
	let timer = null;
	$(`${prefix}Search`).addEventListener("input", () => {
		clearTimeout(timer);
		const txt = $(`${prefix}Search`).value.trim();
		if (txt.length < MIN_SEARCH_CHARS) {
			$(`${prefix}Results`).innerHTML = "";
			return;
		}
		timer = setTimeout(
			() => runPickerSearch(prefix, search, onChoose, txt),
			SEARCH_DEBOUNCE_MS
		);
	});
}

async function runPickerSearch(prefix, search, onChoose, txt) {
	const list = $(`${prefix}Results`);
	try {
		const rows = await search(txt);
		list.innerHTML = "";
		if (!rows || !rows.length) {
			appendMessageRow(list, "No matches found.");
			return;
		}
		rows.forEach((row) =>
			list.appendChild(
				buildResultRow(row, () => {
					list.innerHTML = "";
					$(`${prefix}Search`).value = "";
					onChoose(row);
				})
			)
		);
	} catch (e) {
		list.innerHTML = "";
		appendMessageRow(list, "Search error: " + (e.message || e));
	}
}

function chooseProject(row) {
	clearDone();
	state.project = { value: row.value, label: row.label };
	$("projectChosen").textContent = `→ ${row.label}`;
	show("projectChosen");
	refreshFileButton();
}

function chooseSupplier(row) {
	clearDone();
	state.supplier = { value: row.value, label: row.label };
	$("supplierResults").innerHTML = "";
	$("supplierSearch").value = "";
	$("supplierChosen").textContent = `→ ${row.label}`;
	show("supplierChosen");
	$("supplierHint").textContent = "Change it by searching below.";
}

// Show the job and supplier pickers, with the supplier filled in when the sender's contact
// already belongs to one. Otherwise the PDF decides, or the user picks.
async function showQuoteFields() {
	state.project = null;
	state.supplier = null;
	hide("projectChosen");
	hide("supplierChosen");
	show("quoteFields");
	applyGuessedSupplier();
}

// Fill the supplier from the contact's company or email domain, unless the user picked one.
function applyGuessedSupplier() {
	if (state.supplier && !state.supplier.guessed) return;
	if (state.guessedSupplier) {
		chooseSupplier(state.guessedSupplier);
		state.supplier.guessed = true;
		$("supplierHint").textContent =
			"From the contact's company or email address. Change it by searching below.";
		return;
	}
	state.supplier = null;
	hide("supplierChosen");
	$("supplierHint").textContent =
		"Not known from the contact. Pick it, or leave it blank to read it from the quote PDF.";
}

// Choose to create a new Contact or Candidate from the open email's sender,
// instead of an existing record.
function chooseCreateNew() {
	clearDone();
	const doctype = $("doctype").value;
	const createWord = CREATABLE_TYPES[doctype];
	if (!createWord) return;
	const who = state.senderName || state.senderEmail || `new ${createWord}`;
	state.target = { doctype, value: "", label: who, create: true };
	$("results").innerHTML = "";
	$("search").value = "";
	hide("createNew");
	$("chosen").textContent =
		doctype === QUOTE_TYPE
			? "→ New supplier quote from this email"
			: `→ New ${createWord}: ${who}`;
	show("chosen");
	if (doctype === QUOTE_TYPE) showQuoteFields();
	if (doctype !== QUOTE_TYPE) showPersonEdit(createWord);
	refreshFileButton();
}

// A new Contact or Candidate is for the person the email is about, which the user can correct.
function showPersonEdit(createWord) {
	$("personName").value = state.senderName;
	$("personEmail").value = state.senderEmail;
	const isForwarded = state.person && state.person.is_forwarded;
	$("personEditHint").textContent = isForwarded
		? `Taken from the forwarded email's original sender. Correct them if needed.`
		: `The new ${createWord}'s details. Correct them if needed.`;
	show("personEdit");
}

let searchTimer = null;
function onSearchInput() {
	clearTimeout(searchTimer);
	const txt = $("search").value.trim();
	const isTicket = $("doctype").value === TICKET_TYPE;
	// Ticket types are a short list: an empty box shows them all again.
	if (!isTicket && txt.length < MIN_SEARCH_CHARS) {
		$("results").innerHTML = "";
		return;
	}
	searchTimer = setTimeout(() => runSearch(txt), SEARCH_DEBOUNCE_MS);
}

function setResultsMessage(message) {
	const list = $("results");
	list.innerHTML = "";
	appendMessageRow(list, message);
}

function appendMessageRow(list, message) {
	const listItem = document.createElement("li");
	listItem.textContent = message;
	listItem.className = "result-message";
	list.appendChild(listItem);
}

// One clickable picker row: the label, with its detail (email, supplier…) beneath it.
function buildResultRow(row, onClick) {
	const listItem = document.createElement("li");
	const main = document.createElement("div");
	main.textContent = row.label;
	listItem.appendChild(main);
	if (row.sublabel) {
		const sub = document.createElement("div");
		sub.className = "result-sub";
		sub.textContent = row.sublabel;
		listItem.appendChild(sub);
	}
	listItem.addEventListener("click", onClick);
	return listItem;
}

function renderResultRows(doctype, rows) {
	const list = $("results");
	list.innerHTML = "";
	rows.forEach((row) => list.appendChild(buildResultRow(row, () => chooseTarget(doctype, row))));
}

async function runSearch(txt) {
	const doctype = $("doctype").value;
	setResultsMessage("Searching…");
	try {
		const rows = await apiGet("SearchTargets", { target_doctype: doctype, txt });
		// The user switched type while this search was in flight: drop the stale results.
		if ($("doctype").value !== doctype) return;
		if (!rows || !rows.length) {
			setResultsMessage("No matches found.");
			return;
		}
		renderResultRows(doctype, rows);
	} catch (e) {
		setResultsMessage("Search error: " + (e.message || e));
	}
}

function chooseTarget(doctype, row) {
	clearDone();
	state.target = { doctype, value: row.value, label: row.label, create: false };
	$("results").innerHTML = "";
	$("search").value = "";
	hide("createNew");
	hide("quoteFields");
	$("chosen").textContent = `→ ${row.label}${row.sublabel ? " · " + row.sublabel : ""}`;
	show("chosen");
	// A chosen order can take the email as its supplier invoice.
	if (doctype === ORDER_TYPE) show("invoiceFields");
	refreshFileButton();
}

function refreshFileButton() {
	const hasAttachmentsSelected = state.attachments.some((a) => a.selected);
	const wantsEmail = $("saveEmail").checked;
	const doctype = $("doctype").value;
	const hasComment = $("comment").value.trim() !== "";
	// For a task, choosing the assignee is enough. A ticket needs something to
	// describe it: the email or a comment. A record needs the email, an
	// attachment, or a comment.
	let hasContent = hasAttachmentsSelected || wantsEmail || hasComment;
	if (doctype === TASK_TYPE) hasContent = true;
	if (doctype === TICKET_TYPE) hasContent = wantsEmail || hasComment;
	// A new quote needs its job; an invoice needs the supplier's invoice number.
	const isNewQuote = doctype === QUOTE_TYPE && state.target && state.target.create;
	if (isNewQuote) hasContent = hasContent && Boolean(state.project);
	if (isEnteringInvoice()) hasContent = $("invoiceNumber").value.trim() !== "";
	// Paul, 2026-11-07: after success the button greys out and says so, so nobody presses it
	// again thinking nothing happened. Choosing something else brings it back.
	if (state.isDone) {
		$("fileBtn").disabled = true;
		$("fileBtn").textContent = DONE_BUTTON_TEXT;
		return;
	}
	$("fileBtn").disabled = !state.target || !hasContent;
	$("fileBtn").textContent = buttonText(doctype, isNewQuote);
}

// Something new was chosen after a success: the button is live again.
function clearDone() {
	if (!state.isDone) return;
	state.isDone = false;
	hide("openRecord");
	setStatus("", "");
}

function buttonText(doctype, isNewQuote) {
	if (isEnteringInvoice()) return INVOICE_BUTTON_TEXT;
	if (isNewQuote) return QUOTE_BUTTON_TEXT;
	return (MODE_TEXT[doctype] || RECORD_MODE_TEXT).button;
}

async function onFile() {
	const btn = $("fileBtn");
	btn.disabled = true;
	hide("openRecord");
	const isInvoice = isEnteringInvoice();
	setStatus(isInvoice ? "Entering the supplier invoice…" : "Filing…", "");
	try {
		const payload = await readAttachmentPayload(false);
		const common = Object.assign(
			{
				save_email: $("saveEmail").checked ? 1 : 0,
				comment: $("comment").value,
				body_html: await getEmailBodyHtml(),
				attachments: JSON.stringify(payload),
				message_id: emailMessageId(),
			},
			emailMeta()
		);
		const result = isInvoice ? await enterSupplierInvoice(common) : await fileToRecord(common);
		setStatus(result.message || "Done.", "ok");
		showOpenRecord(result.record_url);
		state.isDone = true;
	} catch (e) {
		setStatus(e.message || "Something went wrong.", "err");
	} finally {
		btn.disabled = false;
		refreshFileButton();
	}
}

function fileToRecord(common) {
	const isCreate = state.target.create === true;
	return apiPost(
		"FileEmailToRecord",
		Object.assign(
			{
				target_doctype: state.target.doctype,
				target_name: state.target.value,
				create_new: isCreate ? 1 : 0,
				new_name: isCreate ? personValue("personName", state.senderName) : "",
				new_email: isCreate ? personValue("personEmail", state.senderEmail) : "",
				contact: chosenContact(),
				project: isCreate && state.project ? state.project.value : "",
				supplier: isCreate && state.supplier ? state.supplier.value : "",
			},
			common
		)
	);
}

function enterSupplierInvoice(common) {
	return apiPost(
		"procurement:EnterSupplierInvoice",
		Object.assign(
			{
				purchase_order: state.target.value,
				supplier_invoice_number: $("invoiceNumber").value.trim(),
				invoice_date: $("invoiceDate").value,
				due_date: $("dueDate").value,
				contact: chosenContact(),
			},
			common
		)
	);
}

// The name or email typed in the "new person" boxes, or the guess when they are hidden.
function personValue(inputId, fallback) {
	if ($("personEdit").classList.contains("hidden")) return fallback;
	return $(inputId).value.trim();
}

// The Contact picked for this email, when the contact picker is in use.
function chosenContact() {
	if ($("personFields").classList.contains("hidden") || !state.contact) return "";
	return state.contact.value;
}

/* ---- The person an email is about (Paul, 2026-11-07) ---- */

// Ask ERP who the email is really from: a colleague's forward names the original sender.
async function guessPeople() {
	const from = Office.context.mailbox.item.from || Office.context.mailbox.item.sender || {};
	try {
		const guess = await apiPost("people:GuessPeople", {
			sender_name: from.displayName || "",
			sender_email: from.emailAddress || "",
			body_html: await getEmailBodyHtml(),
		});
		state.person = guess.person;
		state.senderName = guess.person.name || guess.person.email || "";
		state.senderEmail = guess.person.email || "";
		state.contact = guess.contact;
		state.guessedSupplier = guess.supplier;
	} catch {
		/* keep the Outlook sender: the user can still pick the contact */
	}
}

// Show the contact picker on the types that are not themselves a person.
function refreshPersonFields(doctype) {
	if (NO_CONTACT_PICKER_TYPES.includes(doctype)) {
		hide("personFields");
		return;
	}
	renderChosenContact();
	show("personFields");
}

function renderChosenContact() {
	$("personHint").textContent = describePerson();
	if (!state.contact) {
		hide("contactChosen");
		return;
	}
	const contact = state.contact;
	$("contactChosen").textContent = `→ ${contact.label}${
		contact.sublabel && contact.sublabel !== contact.label ? " · " + contact.sublabel : ""
	}`;
	show("contactChosen");
}

function describePerson() {
	const person = state.person;
	if (!person || !person.email) return "Search for the contact this email is about.";
	const who = person.name ? `${person.name} <${person.email}>` : person.email;
	if (person.is_internal) return `Sent by a colleague (${who}). Pick the contact it is about.`;
	const source = person.is_forwarded ? "Original sender of this forwarded email" : "Sender";
	const found = state.contact ? "" : " No contact holds this address yet.";
	return `${source}: ${who}.${found}`;
}

// The user picked a different contact. On a new quote, re-guess the supplier from it.
async function chooseContact(row) {
	clearDone();
	state.contact = { value: row.value, label: row.label, sublabel: row.sublabel || "" };
	renderChosenContact();
	if (!state.target || !state.target.create || state.target.doctype !== QUOTE_TYPE) return;
	if (!row.sublabel) return;
	const run = ++state.peopleRun;
	try {
		const guess = await apiPost("people:GuessPeople", { sender_email: row.sublabel });
		if (run !== state.peopleRun) return;
		state.guessedSupplier = guess.supplier;
		applyGuessedSupplier();
	} catch {
		/* keep the current supplier */
	}
}

// Offer to open the record just created or filed to, in the user's browser.
function showOpenRecord(url) {
	if (!url) return;
	$("openRecord").dataset.url = url;
	show("openRecord");
}

function openRecord() {
	const url = $("openRecord").dataset.url;
	if (!url) return;
	if (Office.context.ui && Office.context.ui.openBrowserWindow) {
		Office.context.ui.openBrowserWindow(url);
		return;
	}
	window.open(url, "_blank", "noopener");
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
	const from = Office.context.mailbox.item.from || Office.context.mailbox.item.sender || {};
	state.senderName = from.displayName || "";
	state.senderEmail = from.emailAddress || "";
	readAttachments();
	renderAttachments();
	await guessPeople();
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
			throw new Error(
				"Your Microsoft account isn't linked to an ERP user. Contact your administrator."
			);
		}
		await establishErpSession();
		await enterFilingView(session.user);
	} catch (e) {
		$("loginStatus").textContent = "Couldn't sign in automatically.";
		$("loginError").textContent = e.message || "Sign-in failed.";
		show("loginBtn");
	}
}

function wireEvents() {
	$("loginBtn").addEventListener("click", doSignIn);
	$("signOutBtn").addEventListener("click", signOut);
	$("search").addEventListener("input", onSearchInput);
	$("doctype").addEventListener("change", updateModeForType);
	$("createNew").addEventListener("click", chooseCreateNew);
	$("saveEmail").addEventListener("change", refreshFileButton);
	$("comment").addEventListener("input", refreshFileButton);
	$("fileBtn").addEventListener("click", onFile);
	$("openRecord").addEventListener("click", openRecord);
	$("enterInvoice").addEventListener("change", refreshInvoiceInputs);
	$("invoiceNumber").addEventListener("input", refreshFileButton);
	wirePicker(
		"project",
		(txt) => apiGet("SearchTargets", { target_doctype: PROJECT_TYPE, txt }),
		chooseProject
	);
	wirePicker(
		"supplier",
		(txt) => apiGet("procurement:SearchSuppliers", { txt }),
		chooseSupplier
	);
	wirePicker(
		"contact",
		(txt) => apiGet("SearchTargets", { target_doctype: CONTACT_TYPE, txt }),
		chooseContact
	);
}

// Render the environment badge immediately — it is pure DOM and must not depend
// on the Office host initializing (so it shows even if Office.js is slow).
applyEnvBadges();

Office.onReady(async () => {
	wireEvents();
	show("loginView");

	// A remembered ERP session skips sign-in entirely, on every platform.
	$("loginStatus").textContent = "Checking your sign-in…";
	const rememberedUser = await resumeStoredSession();
	if (rememberedUser) {
		await enterFilingView(rememberedUser);
		return;
	}

	// The web sign-in dialog must be opened by a user gesture, so on the web we
	// wait for the button. Desktop can start sign-in silently on load.
	if (auth._isWeb()) {
		$("loginStatus").textContent = "Sign in to continue.";
		show("loginBtn");
	} else {
		await doSignIn();
	}
});
