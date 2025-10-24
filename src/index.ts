import express from 'express';
import { randomUUID } from 'node:crypto';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { isInitializeRequest } from '@modelcontextprotocol/sdk/types.js';
import { getServer } from './mcpserver.js';
import 'dotenv/config';

// --- Express Server Setup ---
const app = express();
app.use(express.json());

// Store active transports by session ID
const transports: Record<string, StreamableHTTPServerTransport> = {};

// --- MCP Endpoint ---
app.all('/mcp', async (req, res) => {
    const sessionId = req.headers['mcp-session-id'] as string | undefined;

    try {
        let transport: StreamableHTTPServerTransport;

        // Reuse existing transport if session ID is valid
        if (sessionId && transports[sessionId]) {
            transport = transports[sessionId];
        } 
        // Initialize a new transport if no session ID and request is valid
        else if (!sessionId && isInitializeRequest(req.body)) {
            transport = new StreamableHTTPServerTransport({
                sessionIdGenerator: randomUUID,
                onsessioninitialized: (newSessionId: string) => {
                    console.log(`Session initialized with ID: ${newSessionId}`);
                    transports[newSessionId] = transport;
                },
            });

            const server = getServer();
            await server.connect(transport);
        } 
        // Handle invalid or missing session ID
        else {
            return res.status(400).json({
                jsonrpc: '2.0',
                error: { code: -32000, message: 'Bad Request: Missing or invalid session ID.' },
                id: null,
            });
        }

        // Process the request using the transport
        await transport.handleRequest(req, res, req.body);
    } catch (error) {
        console.error('Error handling MCP request:', error);

        if (!res.headersSent) {
            res.status(500).json({
                jsonrpc: '2.0',
                error: { code: -32603, message: 'Internal server error.' },
                id: null,
            });
        }
    }
});

// --- Start Server ---
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Dynamics 365 F&O MCP Server listening on port ${PORT}`);
    console.log('Please ensure you have a .env file with your Dynamics 365 credentials.');
});