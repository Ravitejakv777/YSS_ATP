// Polyfill global crypto for older Node versions if needed
if (!global.crypto) {
    global.crypto = require('crypto');
}
let makeWASocket, useMultiFileAuthState, DisconnectReason, fetchLatestWaWebVersion, Browsers;
const express = require('express');
const qrcode = require('qrcode');
const pino = require('pino');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json());

let sock = null;
let qrCodeData = null;
let connectionStatus = 'Disconnected';
let connectedPhone = null;
let recentMessages = []; // Stores last 10 sent messages

let reconnectTimeout = null;

// Helper to fetch WhatsApp version with a timeout to prevent hanging on slow networks
async function fetchLatestVersionWithTimeout(timeoutMs = 4000) {
    return Promise.race([
        fetchLatestWaWebVersion(),
        new Promise((_, reject) => setTimeout(() => reject(new Error('Version fetch timeout')), timeoutMs))
    ]);
}

// Helper to clear the contents of the authentication folder safely (especially if it is a mounted volume)
function clearAuthFolder(folderPath) {
    try {
        if (fs.existsSync(folderPath)) {
            const files = fs.readdirSync(folderPath);
            for (const file of files) {
                fs.rmSync(path.join(folderPath, file), { recursive: true, force: true });
            }
            console.log(`Successfully cleared contents of folder: ${folderPath}`);
        }
    } catch (err) {
        console.error(`Failed to clear folder contents for ${folderPath}:`, err);
    }
}

async function connectToWhatsApp() {
    // Dynamically import Baileys since it is packaged as an ES Module
    if (!makeWASocket) {
        try {
            console.log('Dynamically importing @whiskeysockets/baileys...');
            const baileys = await import('@whiskeysockets/baileys');
            makeWASocket = baileys.default || baileys;
            useMultiFileAuthState = baileys.useMultiFileAuthState;
            DisconnectReason = baileys.DisconnectReason;
            fetchLatestWaWebVersion = baileys.fetchLatestWaWebVersion;
            Browsers = baileys.Browsers;
        } catch (importErr) {
            console.error('Failed to import @whiskeysockets/baileys:', importErr);
            connectionStatus = 'Disconnected';
            qrCodeData = null;
            reconnectTimeout = setTimeout(connectToWhatsApp, 5000);
            return;
        }
    }

    // Clear any pending reconnect timeouts to avoid concurrent connection loops
    if (reconnectTimeout) {
        clearTimeout(reconnectTimeout);
        reconnectTimeout = null;
    }

    // Clean up previous socket if it exists
    if (sock) {
        try {
            console.log('Closing previous WhatsApp socket connection...');
            sock.ev.removeAllListeners('connection.update');
            sock.ev.removeAllListeners('creds.update');
            if (sock.ws) {
                sock.ws.close();
            }
        } catch (e) {
            console.error('Error closing previous socket:', e);
        }
        sock = null;
    }

    const authFolder = path.join(__dirname, 'auth_info_baileys');
    let state, saveCreds;

    try {
        console.log('Loading authentication state from:', authFolder);
        const authState = await useMultiFileAuthState(authFolder);
        state = authState.state;
        saveCreds = authState.saveCreds;
    } catch (err) {
        console.error('Fatal error loading authentication state:', err);
        // Clear folder contents if state files are corrupted
        console.log('Clearing potentially corrupted credentials folder contents...');
        clearAuthFolder(authFolder);
        connectionStatus = 'Disconnected';
        qrCodeData = null;
        console.log('Retrying connection in 5 seconds...');
        reconnectTimeout = setTimeout(connectToWhatsApp, 5000);
        return;
    }

    let version = undefined; // Fallback to library default stable version
    try {
        console.log('Fetching latest WhatsApp Web version...');
        const { version: latestVersion, isLatest } = await fetchLatestVersionWithTimeout(4000);
        console.log(`Fetched latest WA Web version: ${latestVersion.join('.')}. Is latest: ${isLatest}`);
        version = latestVersion;
    } catch (err) {
        console.log('Failed to fetch latest WA version, using library default. Error:', err.message);
    }

    try {
        console.log('Creating WhatsApp Socket...');
        sock = makeWASocket({
            version,
            auth: state,
            logger: pino({ level: 'error' }), // Output only errors to prevent log spam
            browser: Browsers ? Browsers.macOS('Desktop') : ['Chrome (Windows)', 'Chrome', '110.0.5481.177'],
            defaultQueryTimeoutMs: 60000,
            connectTimeoutMs: 60000,
            keepAliveIntervalMs: 30000,   // Ping server every 30 seconds to keep connection alive
            syncFullHistory: false,       // Do NOT sync full history, saving memory, bandwidth, and preventing socket timeouts
            markOnlineOnConnect: false    // Prevents conflicts with active mobile app notifications
        });
    } catch (err) {
        console.error('Fatal error creating WASocket:', err);
        connectionStatus = 'Disconnected';
        qrCodeData = null;
        console.log('Retrying connection in 10 seconds...');
        reconnectTimeout = setTimeout(connectToWhatsApp, 10000);
        return;
    }

    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', (update) => {
        const { connection, lastDisconnect, qr } = update;

        if (connection) {
            console.log(`Connection status updated: ${connection}`);
            if (connection === 'connecting') {
                connectionStatus = 'Connecting';
            }
        }

        if (qr) {
            console.log('New QR code received. Ready to be scanned.');
            qrcode.toDataURL(qr, (err, url) => {
                if (!err) {
                    qrCodeData = url;
                } else {
                    console.error('Failed to convert QR code to Data URL:', err);
                }
            });
            connectionStatus = 'Waiting for scan';
        }

        if (connection === 'close') {
            const statusCode = lastDisconnect?.error?.output?.statusCode;
            const errorReason = lastDisconnect?.error?.message || lastDisconnect?.error;

            const isLoggedOut = statusCode === DisconnectReason.loggedOut;
            const shouldReconnect = !isLoggedOut;

            console.log(`Connection closed (Status ${statusCode}). Reason:`, errorReason);
            if (lastDisconnect?.error) {
                console.error('Full disconnect error stack:', lastDisconnect.error);
            }

            connectionStatus = 'Disconnected';
            qrCodeData = null;
            connectedPhone = null;

            if (isLoggedOut) {
                console.log('Session is logged out. Clearing credentials folder contents...');
                clearAuthFolder(authFolder);
                console.log('Attempting a fresh connection in 2 seconds...');
                reconnectTimeout = setTimeout(connectToWhatsApp, 2000);
            } else if (shouldReconnect) {
                console.log('Attempting reconnection in 10 seconds...');
                reconnectTimeout = setTimeout(connectToWhatsApp, 10000);
            } else {
                console.log('Reconnection aborted.');
            }
        } else if (connection === 'open') {
            console.log('WhatsApp connection opened successfully!');
            connectionStatus = 'Connected';
            qrCodeData = null;
            if (sock && sock.user && sock.user.id) {
                connectedPhone = sock.user.id.split(':')[0].split('@')[0];
                console.log(`Connected phone number: ${connectedPhone}`);
            } else {
                connectedPhone = null;
            }
        }
    });
}


// API to check status and get QR code
app.get('/status', (req, res) => {
    res.json({
        status: connectionStatus,
        qr: qrCodeData,
        phone: connectedPhone
    });
});

// API to manually reset connection
app.post('/reset', async (req, res) => {
    console.log('Manual reset request received. Resetting connection...');
    
    // Clear reconnect timer
    if (reconnectTimeout) {
        clearTimeout(reconnectTimeout);
        reconnectTimeout = null;
    }

    // Reset status
    connectionStatus = 'Disconnected';
    qrCodeData = null;
    connectedPhone = null;

    // Terminate old socket
    if (sock) {
        try {
            console.log('Closing socket during reset...');
            if (sock.ws) {
                sock.ws.close();
            }
        } catch (e) {
            console.error('Error closing socket during reset:', e);
        }
        sock = null;
    }

    // Delete credentials
    const authFolder = path.join(__dirname, 'auth_info_baileys');
    clearAuthFolder(authFolder);

    // Reconnect immediately
    console.log('Initiating fresh WhatsApp connection...');
    connectToWhatsApp();

    res.json({ success: true, message: 'Gateway reset initiated' });
});

// API to send message
app.post('/send', async (req, res) => {
    const { to, message } = req.body;
    if (!to || !message) {
        return res.status(400).json({ error: 'Missing "to" or "message" parameters' });
    }
    
    if (connectionStatus !== 'Connected') {
        return res.status(500).json({ error: 'WhatsApp is not connected yet.' });
    }
    
    try {
        let formattedPhone = to.toString().replace(/\D/g, '');
        if (formattedPhone.length === 10) {
            formattedPhone = '91' + formattedPhone;
        }
        const jid = `${formattedPhone}@s.whatsapp.net`;
        
        await sock.sendMessage(jid, { text: message });
        
        // Log this message in recentMessages (keep last 10)
        recentMessages.unshift({
            to: to.toString(),
            preview: message.substring(0, 80) + (message.length > 80 ? '...' : ''),
            timestamp: new Date().toISOString()
        });
        if (recentMessages.length > 10) recentMessages.pop();
        
        res.json({ success: true });
    } catch (err) {
        console.error('Failed to send message:', err);
        res.status(500).json({ error: err.message });
    }
});

// API to get last 10 sent messages
app.get('/recent-messages', (req, res) => {
    res.json({ messages: recentMessages });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`WhatsApp API Gateway running on port ${PORT}`);
    connectToWhatsApp();
});
