// A small buffer in seconds to ensure we refresh the token before it actually expires.
const TOKEN_EXPIRY_BUFFER = 60; // Buffer in seconds to refresh token before expiry

/**
 * Manages the authentication token for Dynamics 365, including fetching and caching.
 */
export class AuthManager {
    private tokenCache: {
        accessToken: string;
        expiresAt: number;
    } | null = null;

    /**
     * Retrieves a valid access token, refreshing if necessary.  JP
     * @returns {Promise<string>} A valid bearer token.
     */
    public async getToken(): Promise<string> {
        if (this.isTokenValid() && this.tokenCache) {
            console.log('Using cached token.');
            return this.tokenCache.accessToken;
        }

        console.log('Token is invalid or expired. Fetching a new one...');
        return this.fetchToken();
    }

    /**
     * Checks if the cached token is still valid.
     * @returns {boolean} True if the token is valid, false otherwise.
     */
    private isTokenValid(): boolean {
        return this.tokenCache !== null && this.tokenCache.expiresAt > Date.now();
    }

    /**
     * Fetches a new OAuth token from Azure AD.
     * @returns {Promise<string>} The new access token.
     */
    private async fetchToken(): Promise<string> {
        const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, DYNAMICS_RESOURCE_URL } = process.env;

        if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !DYNAMICS_RESOURCE_URL) {
            throw new Error('Missing required environment variables for authentication.');
        }

        const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/token`;

        const params = new URLSearchParams({
            grant_type: 'client_credentials',
            client_id: CLIENT_ID,
            client_secret: CLIENT_SECRET,
            resource: DYNAMICS_RESOURCE_URL,
        });

        try {
            const response = await fetch(tokenUrl, {
                method: 'POST',
                body: params,
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Failed to fetch token: ${response.status} ${errorText}`);
            }

            const data = await response.json();
            const expiresIn = parseInt(data.expires_in, 10);
            const expiresAt = Date.now() + (expiresIn - TOKEN_EXPIRY_BUFFER) * 1000;

            this.tokenCache = { accessToken: data.access_token, expiresAt };

            console.log('Successfully fetched and cached new token.');
            return this.tokenCache.accessToken;
        } catch (error) {
            console.error('Error during token fetch:', error);
            throw error;
        }
    }
}
