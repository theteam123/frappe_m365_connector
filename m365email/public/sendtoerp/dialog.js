/*
 * Send to ERP — sign-in dialog (Office dialog API).
 *
 * Opened by the task pane via Office.context.ui.displayDialogAsync when the
 * in-panel sign-in can't complete (Outlook on the web, where the Microsoft
 * sign-in page is blocked inside the add-in iframe). This runs in a standalone
 * window, so a normal MSAL redirect sign-in works here. On success it returns
 * the ID token to the task pane via messageParent — the panel and dialog do not
 * share storage on the web, so the token must be passed back, not cached.
 */

const MSAL_CLIENT_ID = "d2a895aa-c239-4605-ac8e-d844341887fc";
const MSAL_TENANT_ID = "85e4e8f4-774e-4a96-987f-bc1a3813d984";
const MSAL_SCOPES = ["User.Read"];


function sendToParent(payload) {
	Office.context.ui.messageParent(JSON.stringify(payload));
}


async function runDialogSignIn() {
	const pca = new msal.PublicClientApplication({
		auth: {
			clientId: MSAL_CLIENT_ID,
			authority: `https://login.microsoftonline.com/${MSAL_TENANT_ID}`,
			redirectUri: new URL("dialog.html", window.location.href).href,
		},
		cache: { cacheLocation: "localStorage" },
	});
	await pca.initialize();

	// Returning from the Microsoft sign-in redirect: hand the ID token back.
	const redirectResult = await pca.handleRedirectPromise();
	if (redirectResult && redirectResult.idToken) {
		sendToParent({ status: "ok", idToken: redirectResult.idToken });
		return;
	}

	// First load: start the interactive redirect to the Microsoft sign-in page.
	await pca.loginRedirect({ scopes: MSAL_SCOPES });
}


Office.onReady(() => {
	runDialogSignIn().catch((e) => {
		const message = (e && e.message) ? e.message : "Sign-in failed.";
		document.getElementById("error").textContent = message;
		sendToParent({ status: "error", error: message });
	});
});
