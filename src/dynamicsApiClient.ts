import { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { AuthManager } from './authManager.js';

const authManager = new AuthManager();

// Helper function to safely send notifications
async function notifySafely(
    notify: (notification: any) => void | Promise<void>,
    notification: any
): Promise<void> {
    try {
        await notify(notification);
    } catch {
        // Ignore notification errors
    }
}

/**
 * Makes an API call to Dynamics 365.
 */
export async function callDynamicsApi(
    method: 'GET' | 'POST' | 'PATCH',
    url: string,
    payload: Record<string, unknown> | null,
    notify: (notification: any) => void | Promise<void>
): Promise<CallToolResult> {
    try {
        await notifySafely(notify, {
            method: "notifications/message",
            params: { level: "info", data: `Calling ${method} ${url}` },
        });

        const token = await authManager.getToken();

        const response = await fetch(url, {
            method,
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
                'Accept': 'application/json, text/xml',
                'Prefer': 'odata.maxpagesize=100', // Pagination preference
            },
            ...(payload && { body: JSON.stringify(payload) }),
        });

        if (response.status === 204) {
            return { content: [{ type: 'text', text: 'Operation successful (No Content).' }] };
        }

        const responseText = await response.text();
        if (!response.ok) {
            throw new Error(`API call failed: ${responseText}`);
        }

        return { content: [{ type: 'text', text: responseText }] };
    } catch (error) {
        console.error('Error during API call:', error);
        throw error;
    }
}