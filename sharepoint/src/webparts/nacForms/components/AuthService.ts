import { UserManager, User } from 'oidc-client-ts';
import { Config } from './Config';

// SharePoint-only authentication: standard OAuth 2.0 PKCE full-page redirect.
export class AuthService {
    private userManager: UserManager;
    private config: Config;

    constructor(config: Config) {
        this.config = config;
        this.userManager = new UserManager(config.getSettings());
    }

    // Full-page redirect to the Nintex auth server.
    public login(): Promise<void> {
        return this.userManager.signinRedirect();
    }

    // Completes the OAuth flow when the page loads with ?code= in the URL.
    public async handleAuthentication(): Promise<void> {
        await this.userManager.signinCallback();
    }

    // True if the current page load is an OAuth callback (?code= present).
    public static isAuthCallback(): boolean {
        return new URL(window.location.href).searchParams.has('code');
    }

    public getSettings(): { authority: string } {
        return this.userManager.settings as unknown as { authority: string };
    }

    // Attempts a silent token renewal via a hidden iframe.  Works only if
    // the user has an active session on the Nintex auth server AND the
    // silent-callback.html page is uploaded to the site's SiteAssets.
    // Returns the access token on success, undefined on failure.
    public async trySilentLogin(): Promise<string | undefined> {
        try {
            const user: User | null = await this.userManager.signinSilent();
            if (user && user.access_token) {
                return user.access_token;
            }
        } catch (error) {
            // Silent login failed — no active Nintex session, iframe blocked,
            // or silent-callback.html not found.  Caller should fall back to
            // showing the login button.
            console.log('[AuthService] Silent login failed:', (error as Error).message);
        }
        return undefined;
    }

    // Retrieves the access token ONLY if a stored session exists and the
    // token has not yet expired.  Used on page load to resume a previous
    // session without forcing the user to click "Log In" again.
    public async getValidAccessToken(): Promise<string | undefined> {
        // Try the canonical UserManager lookup first.
        const user: User | null = await this.userManager.getUser();
        if (user && !user.expired && user.access_token) {
            return user.access_token;
        }

        // Fall back to scanning localStorage (handles the underscore
        // client-ID key mismatch), but also check expiry.
        const authority = this.config.getSettings().authority;
        const prefix = `oidc.user:${authority}:`;
        const keys = Object.keys(localStorage);
        for (const key of keys) {
            if (key.indexOf(prefix) === 0) {
                try {
                    const stored = JSON.parse(localStorage.getItem(key) || '{}');
                    if (stored && stored.access_token && stored.expires_at) {
                        if (stored.expires_at > Date.now() / 1000) {
                            return stored.access_token;
                        }
                    }
                } catch {
                    // ignore parse errors
                }
            }
        }
        return undefined;
    }

    // Retrieves the access token. Tries getUser() first, then falls back to
    // scanning localStorage to handle the underscore in Nintex client IDs
    // which can cause a key mismatch in oidc-client-ts's getUser().
    public async getAccessToken(): Promise<string | undefined> {
        const user: User | null = await this.userManager.getUser();
        if (user && user.access_token) {
            return user.access_token;
        }

        const authority = 'https://auth.nintexcloud.com';
        const prefix = `oidc.user:${authority}:`;
        const keys = Object.keys(localStorage);
        for (const key of keys) {
            if (key.indexOf(prefix) === 0) {
                try {
                    const stored = JSON.parse(localStorage.getItem(key) || '{}');
                    if (stored && stored.access_token) {
                        return stored.access_token;
                    }
                } catch {
                    // ignore parse errors
                }
            }
        }
        return undefined;
    }
}
