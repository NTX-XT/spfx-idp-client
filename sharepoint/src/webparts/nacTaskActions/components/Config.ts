import { UserManagerSettings, WebStorageStateStore } from "oidc-client-ts";


// Config class that holds the configuration settings for the app
export class Config {
    private authority: string;
    private client_id: string;
    private resource: string;
    private redirect_uri: string;
    private response_type: string;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private extraTokenParams: any;
    private region: string;
    private useProd: boolean;

    // Constructor to allow for customisation of the settings
    constructor(client_id: string, redirect_uri: string, region: string, useProd: boolean) {
        this.useProd = useProd;
        this.client_id = client_id;
        this.resource = 'urn:nc';
        this.redirect_uri = redirect_uri;
        this.response_type = 'code';
        this.region = region;
        this.extraTokenParams = {
            resource: 'urn:nc'
        }

        // Set the authority based on the environment
        this.authority = this.useProd ? 'https://auth.nintexcloud.com' : 'https://auth.nintexcloudtest.com';
    }


    // Returns the settings for the app
    public getSettings(): UserManagerSettings {
        // Inline OIDC metadata so oidc-client-ts never needs to fetch the
        // discovery document. This is critical for Teams web — the Teams
        // iframe CSP blocks fetch() calls to auth.nintexcloud.com, which
        // causes "Failed to fetch" errors during signinRedirect() when it
        // tries to resolve the authorization endpoint from the discovery doc.
        //
        // These values come from:
        // https://auth.nintexcloud.com/.well-known/openid-configuration
        //
        // NOTE: If using the test environment (auth.nintexcloudtest.com),
        // update these endpoints to match that environment's discovery doc.
        const authBase = this.authority;
        const metadata = {
            issuer: authBase,
            authorization_endpoint: `${authBase}/connect/authorize`,
            token_endpoint: `${authBase}/connect/token`,
            userinfo_endpoint: `${authBase}/connect/userinfo`,
            end_session_endpoint: `${authBase}/connect/endsession`,
            revocation_endpoint: `${authBase}/connect/revocation`,
            jwks_uri: `${authBase}/.well-known/openid-configuration/jwks`
        };

        return {
            authority: this.authority,
            client_id: this.client_id,
            resource: this.resource,
            redirect_uri: this.redirect_uri,
            popup_redirect_uri: this.redirect_uri,
            silent_redirect_uri: this.getSilentRedirectUri(),
            response_type: this.response_type,
            // Automatically refresh the token in the background before it
            // expires.  Uses a hidden iframe + silent_redirect_uri.  Requires
            // silent-callback.html to be uploaded to the site's SiteAssets.
            automaticSilentRenew: true,
            // How many seconds before expiry to trigger the silent renew
            // (default 60; we use 120 to give a comfortable buffer).
            accessTokenExpiringNotificationTimeInSeconds: 120,
            // Provide metadata inline to avoid discovery document fetch
            metadata: metadata,
            userStore: new WebStorageStateStore({ store: window.localStorage}),
            stateStore: new WebStorageStateStore({ store: window.localStorage }),
            extraTokenParams: this.extraTokenParams
        }
    }

    // Returns the region for the app
    public getRegion(): string {
        return this.region;
    }

    // Returns the useProd setting for the app
    public getUseProd(): boolean {
        return this.useProd;
    }

    // Derives the silent callback URL from the redirect_uri.
    // Assumes silent-callback.html is uploaded to the site's SiteAssets library.
    // e.g. redirect_uri = https://tenant.sharepoint.com/MySite/SitePages/Home.aspx
    //   →  silent_redirect_uri = https://tenant.sharepoint.com/MySite/SiteAssets/silent-callback.html
    private getSilentRedirectUri(): string {
        try {
            const url = new URL(this.redirect_uri);
            const pathParts = url.pathname.split('/');
            // Remove the page segments (e.g. /SitePages/Home.aspx) to get the site path
            const pageIndex = pathParts.findIndex(p =>
                p.toLowerCase() === 'sitepages' || p.toLowerCase() === 'pages'
            );
            const sitePath = pageIndex > 0
                ? pathParts.slice(0, pageIndex).join('/')
                : pathParts.slice(0, -1).join('/');
            return `${url.origin}${sitePath}/SiteAssets/silent-callback.html`;
        } catch {
            // Fallback: use redirect_uri directory + SiteAssets
            return this.redirect_uri.replace(/\/[^/]*$/, '/../SiteAssets/silent-callback.html');
        }
    }
}
