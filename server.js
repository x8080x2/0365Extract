const express = require('express');
const cors = require('cors');
const path = require('path');
const { OutlookLoginAutomation } = require('./src/outlook-login');

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware - Configure CORS for Replit environment
app.use(cors({
    origin: true, // Allow all origins for Replit proxy
    credentials: true
}));
app.use(express.json());
app.use(express.static('public'));

// Store single automation instance for email extraction only
let activeSession = null;
const SESSION_TIMEOUT = 30 * 60 * 1000; // 30 minutes timeout

// Helper function to initialize browser for email extraction
async function initEmailExtractor() {
    if (activeSession && activeSession.automation) {
        return activeSession.automation;
    }

    const automation = new OutlookLoginAutomation();
    await automation.init();
    
    activeSession = {
        sessionId: Date.now().toString(),
        automation: automation,
        createdAt: Date.now()
    };

    console.log(`Email extractor initialized for session ${activeSession.sessionId}`);
    return automation;
}

// Cleanup expired session
setInterval(async () => {
    if (activeSession) {
        const now = Date.now();
        if (now - activeSession.createdAt > SESSION_TIMEOUT) {
            console.log(`🧹 Cleaning up expired session: ${activeSession.sessionId}`);
            
            try {
                if (activeSession.automation) {
                    await activeSession.automation.close();
                }
                activeSession = null;
            } catch (error) {
                console.error(`Error closing expired session:`, error);
            }
        }
    }
}, 5 * 60 * 1000); // Check every 5 minutes

// Routes

// Health check
app.get('/api/health', (req, res) => {
    res.json({ status: 'OK', message: 'Email Extraction Service is running' });
});

// Extract emails from Outlook (main functionality)
app.get('/api/extract-emails', async (req, res) => {
    try {
        console.log('Starting email extraction...');
        const automation = await initEmailExtractor();
        
        // Navigate to Outlook
        const navigated = await automation.navigateToOutlook();
        if (!navigated) {
            return res.status(500).json({ 
                error: 'Failed to navigate to Outlook' 
            });
        }

        // Scan all emails
        const allEmails = await automation.scanAllEmails();

        const response = {
            sessionId: activeSession.sessionId,
            data: allEmails,
            summary: {
                inboxCount: allEmails.inbox?.length || 0,
                sentCount: allEmails.sent?.length || 0,
                totalCount: (allEmails.inbox?.length || 0) + (allEmails.sent?.length || 0)
            },
            extractionTimestamp: new Date().toISOString()
        };

        if (allEmails.error) {
            response.warning = allEmails.error;
        }

        res.json(response);

    } catch (error) {
        console.error('Error during email extraction:', error);
        res.status(500).json({ 
            error: 'Failed to extract emails',
            details: error.message 
        });
    }
});

// Close current session
app.delete('/api/session', async (req, res) => {
    try {
        if (!activeSession) {
            return res.status(400).json({ 
                error: 'No active session to close' 
            });
        }

        if (activeSession.automation) {
            await activeSession.automation.close();
        }

        activeSession = null;

        res.json({
            message: 'Session closed successfully'
        });

    } catch (error) {
        console.error('Error closing session:', error);
        res.status(500).json({ 
            error: 'Failed to close session',
            details: error.message 
        });
    }
});

// Get current session status
app.get('/api/status', (req, res) => {
    res.json({
        hasActiveSession: activeSession !== null,
        sessionId: activeSession?.sessionId || null,
        sessionAge: activeSession ? Date.now() - activeSession.createdAt : null,
        service: 'Email Extraction Only'
    });
});

// Start server
app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Email Extraction Service running on port ${PORT}`);
    console.log(`📧 Main endpoint: GET /api/extract-emails`);
    console.log(`🔍 Service focuses on email extraction only`);
});

// Graceful shutdown
process.on('SIGINT', async () => {
    console.log('\n🛑 Received SIGINT, shutting down gracefully...');
    
    if (activeSession && activeSession.automation) {
        try {
            await activeSession.automation.close();
        } catch (error) {
            console.error('Error closing session during shutdown:', error);
        }
    }
    
    process.exit(0);
});

process.on('SIGTERM', async () => {
    console.log('\n🛑 Received SIGTERM, shutting down gracefully...');
    
    if (activeSession && activeSession.automation) {
        try {
            await activeSession.automation.close();
        } catch (error) {
            console.error('Error closing session during shutdown:', error);
        }
    }
    
    process.exit(0);
});