const CLIENT_ID = "a18f7ac2-a67a-425e-9abd-140b5c1278b5";
const TENANT_ID = "0208acf3-2bb4-4c3b-80d9-36a683f79fd6";
const BASE_URL  = "https://carsair2026-bi.github.io/Tablero_GrupoCarsa/";

const msalConfig = {
    auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/" + TENANT_ID,
        redirectUri: BASE_URL + "redirect.html",
        postLogoutRedirectUri: BASE_URL + "index.html"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
    }
};

const LOGIN_REQUEST = { scopes: ["User.Read"] };

function cargarScript(src) {
    return new Promise((resolve, reject) => {
        const s = document.createElement("script");
        s.src = src;
        s.onload = resolve;
        s.onerror = () => reject(new Error("No se pudo cargar " + src));
        document.head.appendChild(s);
    });
}

async function cargarMsal() {
    if (typeof msal !== "undefined") return;
    try {
        await cargarScript("https://alcdn.msauth.net/browser/2.37.1/js/msal-browser.min.js");
    } catch (e) {
        await cargarScript("https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.37.1/lib/msal-browser.min.js");
    }
}

let msalInstance = null;

async function getMsal() {
    if (msalInstance) return msalInstance;
    await cargarMsal();
    msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();
    return msalInstance;
}

async function obtenerCuenta() {
    const app = await getMsal();
    const respuesta = await app.handleRedirectPromise();

    let cuenta = (respuesta && respuesta.account) || app.getAllAccounts()[0] || null;

    if (cuenta && cuenta.tenantId && cuenta.tenantId !== TENANT_ID) {
        cuenta = null;
    }

    if (cuenta) app.setActiveAccount(cuenta);
    return cuenta;
}

function tieneAccesoUsuario() {
    return sessionStorage.getItem("accesoValidado") === "usuarioContrasena";
}
