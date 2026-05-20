// Polyfill global crypto for older Node versions if needed
if (!global.crypto) {
    global.crypto = require('crypto');
}
const { default: makeWASocket, useMultiFileAuthState, DisconnectReason, fetchLatestWaWebVersion } = require('@whiskeysockets/baileys');
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
let recentMessages = []; // Stores last 10 sent messages

async function connectToWhatsApp() {
    const authFolder = path.join(__dirname, 'auth_info_baileys');
    const { state, saveCreds } = await useMultiFileAuthState(authFolder);
    
    let version = [2, 3000, 1017531287]; // Fallback stable version
    try {
        const { version: latestVersion, isLatest } = await fetchLatestWaWebVersion();
        console.log(`Fetched latest WA Web version: ${latestVersion.join('.')}. Is latest: ${isLatest}`);
        version = latestVersion;
    } catch (err) {
        console.log('Failed to fetch latest WA version, using fallback stable version. Error: ', err.message);
    }
    
    sock = makeWASocket({
        version,
        auth: state,
        logger: pino({ level: 'silent' }), // Hide noisy debug logs
        browser: ['Chrome (Windows)', 'Chrome', '110.0.5481.177'],
        defaultQueryTimeoutMs: 60000,
        connectTimeoutMs: 60000
    });
    
    sock.ev.on('creds.update', saveCreds);
    
    sock.ev.on('connection.update', (update) => {
        const { connection, lastDisconnect, qr } = update;
        
        if (qr) {
            qrcode.toDataURL(qr, (err, url) => {
                if (!err) {
                    qrCodeData = url;
                }
            });
            connectionStatus = 'Waiting for scan';
        }
        
        if (connection === 'close') {
            const statusCode = lastDisconnect?.error?.output?.statusCode;
            const errorReason = lastDisconnect?.error?.message || lastDisconnect?.error;
            const shouldReconnect = statusCode !== DisconnectReason.loggedOut;
            
            console.log(`Connection closed (Status ${statusCode}). Reason:`, errorReason);
            if (lastDisconnect?.error) {
                console.log('Full disconnect error stack:', lastDisconnect.error);
            }
            
            connectionStatus = 'Disconnected';
            qrCodeData = null;
            if (shouldReconnect) {
                console.log('Attempting reconnection in 10 seconds...');
                setTimeout(connectToWhatsApp, 10000); // 10s throttle
            }
        } else if (connection === 'open') {
            console.log('WhatsApp connection opened successfully!');
            connectionStatus = 'Connected';
            qrCodeData = null;
        }
    });
}

// API to check status and get QR code
app.get('/status', (req, res) => {
    res.json({
        status: connectionStatus,
        qr: qrCodeData
    });
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
