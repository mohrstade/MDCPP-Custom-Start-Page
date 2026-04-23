import {
    AccountInfo,
    PopupRequest,
    PublicClientApplication,
    SilentRequest,
} from "@azure/msal-browser";

const aadClientId = process.env.ENTRA_APPID ?? "";
if (aadClientId === "") {
    throw new Error("Test app Entra client ID is missing. Please ensure it's defined in the .env file.");
}

const redirectUri = window.location.pathname;
const defaultScopes = ["User.Read", "Files.Read"];
const LAST_ACCOUNT_ID_KEY = "team-startpage:lastAccountId";
const LAST_USERNAME_KEY = "team-startpage:lastUsername";

const app = new PublicClientApplication({
    auth: {
        authority: "https://login.microsoftonline.com/common",
        clientId: aadClientId,
        redirectUri,
    },
    cache: {
        cacheLocation: "localStorage",
    },
});

let activeAccount: AccountInfo | null = null;
let initialized = false;

export function combine(...paths: string[]) {
    return paths
        .map(path => path.replace(/^[\\|/]/, "").replace(/[\\|/]$/, ""))
        .join("/")
        .replace(/\\/g, "/");
}

function setActiveAccount(account: AccountInfo | null) {
    activeAccount = account;
    if (account) {
        app.setActiveAccount(account);
        window.localStorage.setItem(LAST_ACCOUNT_ID_KEY, account.homeAccountId);
        if (account.username) {
            window.localStorage.setItem(LAST_USERNAME_KEY, account.username);
        }
    } else {
        app.setActiveAccount(null);
        window.localStorage.removeItem(LAST_ACCOUNT_ID_KEY);
        window.localStorage.removeItem(LAST_USERNAME_KEY);
    }
}

export async function initializeAuth() {
    if (!initialized) {
        await app.initialize();
        const response = await app.handleRedirectPromise();
        if (response?.account) {
            setActiveAccount(response.account);
        } else {
            setActiveAccount(null);
        }
        initialized = true;
    }

    return activeAccount;
}

export function getActiveAccount() {
    return activeAccount;
}

export function isSignedIn() {
    return activeAccount !== null;
}

export async function signIn() {
    const request: PopupRequest = {
        prompt: "login",
        scopes: defaultScopes,
    };

    const response = await app.loginPopup(request);
    setActiveAccount(response.account);
    return response.account;
}

export async function signOut() {
    if (!activeAccount) return;

    await app.logoutPopup({
        account: activeAccount,
        mainWindowRedirectUri: window.location.origin + window.location.pathname,
    });

    setActiveAccount(null);
}

function normalizeScopes(scopes: string[]) {
    const merged = new Set<string>([...defaultScopes, ...scopes]);
    return Array.from(merged);
}

export async function getGraphToken(scopes: string[] = defaultScopes) {
    if (!activeAccount) {
        throw new Error("User is not signed in.");
    }

    const resolvedScopes = normalizeScopes(scopes);
    const silentRequest: SilentRequest = {
        account: activeAccount,
        scopes: resolvedScopes,
    };

    try {
        const response = await app.acquireTokenSilent(silentRequest);
        return response.accessToken;
    } catch (error: any) {
        const message = error?.message || "";
        if (message.toLowerCase().includes("interaction")) {
            throw new Error("Your Microsoft session needs attention or new permissions must be accepted. Please use the sign out button and reconnect when you are ready.");
        }
        throw error;
    }
}

export async function getToken(command: { resource: string; type: string }) {
    if (command.type === "Default") {
        return getGraphToken(defaultScopes);
    }

    if (command.type === "SharePoint" || command.type === "SharePoint_SelfIssued") {
        return getGraphToken([`${combine(command.resource, ".default")}`]);
    }

    return "";
}

export async function graphGet<T>(url: string, scopes: string[] = defaultScopes): Promise<T> {
    return graphRequest<T>(url, {
        method: "GET",
    }, scopes);
}

export async function graphPost<T>(url: string, body: Record<string, unknown> = {}, scopes: string[] = defaultScopes): Promise<T> {
    return graphRequest<T>(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
    }, scopes);
}

async function graphRequest<T>(url: string, init: RequestInit, scopes: string[] = defaultScopes): Promise<T> {
    const token = await getGraphToken(scopes);
    const response = await fetch(url, {
        ...init,
        headers: {
            Authorization: `Bearer ${token}`,
            ...(init.headers || {}),
        },
    });

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
        const errorMessage = (data as any)?.error?.message || response.statusText || "Graph request failed";
        throw new Error(errorMessage);
    }

    return data as T;
}
