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

// Store single automation instance - only one session allowed
let activeSession = null; // { sessionId, automation, isPreloaded, createdAt, email }
const SESSION_TIMEOUT = 30 * 60 * 1000; // 30 minutes timeout
const OPERATION_TIMEOUT = 60 * 1000; // 1 minute for individual operations
const HEALTH_CHECK_INTERVAL = 2 * 60 * 1000; // 2 minutes
let sessionMutex = null; // Prevents race conditions in session management
let initializingSession = false; // Prevents concurrent browser initialization

// Helper function to initialize browser directly - Prevents concurrent initialization
async function initBrowser(session) {
    // Prevent concurrent browser initialization
    if (initializingSession) {
        throw new Error('Browser initialization already in progress');
    }

    initializingSession = true;

    try {
        // Close any existing automation with proper timeout
        if (session.automation) {
            try {
                console.log(`Gracefully closing existing automation for session ${session.sessionId}...`);
                await Promise.race([
                    session.automation.close().catch(err => console.error('Session close error:', err)),
                    new Promise((_, reject) => setTimeout(() => reject(new Error('Close timeout')), 5000))
                ]);
                await new Promise(resolve => setTimeout(resolve, 1000)); // Wait 1 second after close
            } catch (error) {
                console.error('Error closing existing session:', error);
            }
            session.automation = null;
        }

        session.automation = new OutlookLoginAutomation();
        await session.automation.init();

        console.log(`Browser initialized successfully for session ${session.sessionId}`);
        return session.automation;
    } catch (error) {
        console.error(`Failed to initialize browser for session ${session.sessionId}:`, error);
        session.automation = null;
        throw error;
    } finally {
        initializingSession = false;
    }
}

// Cleanup expired session - Avoid race conditions with active operations
setInterval(async () => {
    if (activeSession && !sessionMutex && !initializingSession) {
        const now = Date.now();
        // Only cleanup if session is expired AND not currently in use
        if (now - activeSession.createdAt > SESSION_TIMEOUT && !activeSession.inUse) {
            console.log(`🧹 Cleaning up expired session: ${activeSession.sessionId}`);
            
            // Set mutex to prevent other operations during cleanup
            sessionMutex = Promise.resolve();
            
            try {
                if (activeSession.automation) {
                    await activeSession.automation.close();
                }
                activeSession = null;
            } catch (error) {
                console.error(`Error closing expired session ${activeSession.sessionId}:`, error);
            } finally {
                sessionMutex = null;
            }
        }
    }
}, 5 * 60 * 1000); // Check every 5 minutes

// Health check for browser sessions
setInterval(async () => {
    if (activeSession && activeSession.automation && !sessionMutex && !initializingSession) {
        try {
            // Check if browser is still connected
            const automation = activeSession.automation;
            if (automation.browser && !automation.browser.isConnected()) {
                console.log(`🔧 Detected disconnected browser for session ${activeSession.sessionId}`);
                
                // Clean up disconnected session
                sessionMutex = Promise.resolve();
                try {
                    await automation.close();
                    activeSession = null;
                    console.log('Cleaned up disconnected browser session');
                } catch (error) {
                    console.error('Error cleaning up disconnected session:', error);
                } finally {
                    sessionMutex = null;
                }
            }
        } catch (error) {
            console.error('Health check error:', error);
        }
    }
}, HEALTH_CHECK_INTERVAL);

// Helper function to get or create session (only one allowed) - Thread-safe with mutex
async function getOrCreateSession(sessionId = null) {
    // Wait for any ongoing session operations to complete
    while (sessionMutex) {
        await new Promise(resolve => setTimeout(resolve, 100));
    }

    // Acquire mutex
    sessionMutex = Promise.resolve();

    try {
        // If there's an active session and it matches the requested one, return it
        if (activeSession && sessionId && activeSession.sessionId === sessionId) {
            return { sessionId: activeSession.sessionId, session: activeSession, isNew: false };
        }

        // Close any existing session before creating a new one
        if (activeSession) {
            console.log(`🔄 Closing existing session: ${activeSession.sessionId}`);
            try {
                if (activeSession.automation) {
                    await Promise.race([
                        activeSession.automation.close().catch(err => console.error('Session close error:', err)),
                        new Promise((_, reject) => setTimeout(() => reject(new Error('Close timeout')), 3000))
                    ]);
                    await new Promise(resolve => setTimeout(resolve, 500)); // Brief wait after close
                }
            } catch (error) {
                console.error(`Error closing existing session:`, error);
            }
            activeSession = null;
        }

        // Create new single session
        const newSessionId = Date.now().toString();
        activeSession = {
            sessionId: newSessionId,
            automation: null,
            isPreloaded: false,
            createdAt: Date.now(),
            email: null,
            inUse: false // Track if session is being actively used
        };

        console.log(`📝 Created new session: ${newSessionId} (Single session mode)`);
        return { sessionId: newSessionId, session: activeSession, isNew: true };

    } finally {
        // Release mutex
        sessionMutex = null;
    }
}

// Routes

// Health check
app.get('/api/health', (req, res) => {
    res.json({ status: 'OK', message: 'Outlook Automation Backend is running' });
});

// Preload Outlook page
app.post('/api/preload', async (req, res) => {
    try {
        const requestedSessionId = req.body.sessionId;
        const { sessionId, session, isNew } = await getOrCreateSession(requestedSessionId);

        // Mark session as in use to prevent cleanup during operation
        session.inUse = true;

        try {
            // If already preloaded for this session, return status
            if (session.isPreloaded && session.automation) {
                return res.json({
                    status: 'already-loaded',
                    message: 'Outlook page is already loaded and ready',
                    sessionId: sessionId
                });
            }

        // Close any existing automation for this session
        if (session.automation) {
            console.log(`Closing existing automation for session ${sessionId}...`);
            try {
                await session.automation.close();
            } catch (error) {
                console.error('Error closing existing session:', error);
            }
        }

        // Start new automation session for preloading
        console.log(`Preloading Outlook page for session ${sessionId}...`);

        // Initialize browser directly with error handling and timeout
        try {
            await Promise.race([
                initBrowser(session),
                new Promise((_, reject) => setTimeout(() => reject(new Error('Browser initialization timeout')), OPERATION_TIMEOUT))
            ]);
        } catch (error) {
            console.error('Failed to initialize browser:', error);
            
            // Clean up on timeout or error
            if (session.automation) {
                try {
                    await session.automation.close();
                } catch (closeError) {
                    console.error('Error cleaning up failed initialization:', closeError);
                }
                session.automation = null;
            }
            
            return res.status(500).json({ 
                error: 'Failed to initialize browser',
                details: error.message,
                retryable: true
            });
        }

        // Navigate to Outlook with timeout
        try {
            const navigated = await Promise.race([
                session.automation.navigateToOutlook(),
                new Promise((_, reject) => setTimeout(() => reject(new Error('Navigation timeout')), OPERATION_TIMEOUT))
            ]);
            
            if (!navigated) {
                throw new Error('Navigation failed');
            }
        } catch (error) {
            console.error('Failed to navigate to Outlook:', error);
            
            // Clean up on navigation failure
            try {
                await session.automation.close();
            } catch (closeError) {
                console.error('Error cleaning up after navigation failure:', closeError);
            }
            session.automation = null;
            session.isPreloaded = false;
            
            return res.status(500).json({ 
                error: 'Failed to preload Outlook page',
                details: error.message
            });
        }

        session.isPreloaded = true;
        console.log(`Outlook page preloaded successfully for session ${sessionId}`);

            res.json({
                status: 'preloaded',
                message: 'Outlook page loaded and ready for email input',
                sessionId: sessionId
            });

        } finally {
            // Mark session as no longer in use
            if (session) {
                session.inUse = false;
            }
        }

    } catch (error) {
        console.error('Error preloading Outlook:', error);

        res.status(500).json({ 
            error: 'Failed to preload Outlook',
            details: error.message 
        });
    }
});

// Process email login (uses preloaded page if available)
app.post('/api/login', async (req, res) => {
    try {
        const { email, password, sessionId: requestedSessionId } = req.body;

        if (!email) {
            return res.status(400).json({ 
                error: 'Email is required' 
            });
        }

        const { sessionId, session, isNew } = await getOrCreateSession(requestedSessionId);
        session.email = email; // Store email in session
        
        // Mark session as in use to prevent cleanup during login
        session.inUse = true;

        try {
            // If not preloaded, start fresh session
        if (!session.isPreloaded || !session.automation) {
            console.log(`Starting fresh Puppeteer session for email: ${email} (Session: ${sessionId})`);

            // Close any existing automation
            if (session.automation) {
                try {
                    await session.automation.close();
                } catch (error) {
                    console.error('Error closing existing session:', error);
                }
            }

            // Initialize browser directly with error handling
            try {
                await initBrowser(session);
            } catch (error) {
                console.error('Failed to initialize browser for login:', error);
                return res.status(500).json({ 
                    error: 'Failed to initialize browser',
                    details: error.message,
                    retryable: true
                });
            }

            // Navigate to Outlook
            const navigated = await session.automation.navigateToOutlook();
            if (!navigated) {
                await session.automation.close();
                session.automation = null;
                session.isPreloaded = false;
                return res.status(500).json({ 
                    error: 'Failed to navigate to Outlook' 
                });
            }

            session.isPreloaded = true;
        } else {
            console.log(`Using preloaded Outlook page for email: ${email} (Session: ${sessionId})`);
        }

        // Force fresh login - no checking for existing sessions

        // If password is provided, navigate to email interface for scanning
        if (password) {
            console.log('🔐 Proceeding to email interface...');

            const navigated = await session.automation.navigateToEmailInterface();
            
            // Take screenshot after navigation
            await session.automation.takeScreenshot(`screenshots/session-${sessionId}-email-interface.png`);

            res.json({
                sessionId: sessionId,
                email: email,
                loginComplete: true,
                loginSuccess: navigated,
                authMethod: 'email-scanning',
                message: navigated ? 'Ready for email scanning' : 'Unable to access email interface',
                screenshots: [
                    `screenshots/session-${sessionId}-email-interface.png`
                ]
            });
        } else {
            // Fill email and click Next to get site response
            console.log('Filling email and clicking Next to get site response...');

            let siteReport = {
                emailFilled: false,
                nextClicked: false,
                siteResponse: '',
                errorMessages: [],
                pageUrl: '',
                pageTitle: '',
                needsPassword: false,
                needsMFA: false,
                accountExists: false
            };

            try {
                // Wait for email input and fill it
                await session.automation.page.waitForSelector('input[type="email"]', { timeout: 10000 });
                await session.automation.page.type('input[type="email"]', email);
                console.log('Email entered successfully');
                siteReport.emailFilled = true;

                // Click Next button
                await session.automation.page.click('input[type="submit"]');
                console.log('Clicked Next button');
                siteReport.nextClicked = true;

                // Wait for page response (up to 10 seconds)
                await new Promise(resolve => setTimeout(resolve, 3000));

                // Get current page info
                siteReport.pageUrl = session.automation.page.url();
                siteReport.pageTitle = await session.automation.page.title();

                // Check for different scenarios

                // Navigate to email interface for scanning
                console.log('Navigating to email interface for email scanning...');
                const navigated = await session.automation.navigateToEmailInterface();
                
                if (navigated) {
                    siteReport.siteResponse = 'Ready for email scanning';
                    siteReport.accountExists = true;
                    console.log('Email interface accessible');
                } else {
                    siteReport.siteResponse = 'Unable to access email interface';
                    console.log('Email interface not accessible');
                }

                // If no specific response detected, get general page content
                if (!siteReport.siteResponse) {
                    try {
                        // Look for main content or messages
                        const mainContent = await session.automation.page.$eval('body', el => {
                            // Remove scripts and styles
                            const scripts = el.querySelectorAll('script, style, noscript');
                            scripts.forEach(s => s.remove());

                            // Get visible text content
                            const text = el.textContent || '';
                            return text.replace(/\s+/g, ' ').trim().substring(0, 500);
                        });

                        siteReport.siteResponse = mainContent || 'Page loaded but no specific response detected';
                    } catch (e) {
                        siteReport.siteResponse = 'Page loaded successfully';
                    }
                }

                // Take screenshot after clicking Next
                await session.automation.takeScreenshot(`screenshots/session-${sessionId}-after-next.png`);

                console.log('Site report:', JSON.stringify(siteReport, null, 2));

            } catch (error) {
                console.error('Error during email/next process:', error);
                siteReport.errorMessages.push(`Automation error: ${error.message}`);
                siteReport.siteResponse = `Error occurred: ${error.message}`;
            }

            res.json({
                sessionId: sessionId,
                email: email,
                loginComplete: false,
                siteReport: siteReport,
                message: siteReport.accountExists ? 
                    'Email interface accessible - ready for scanning.' : 
                    'Email interface not accessible - see site report for details.',
                screenshots: [
                    `screenshots/session-${sessionId}-after-next.png`
                ]
            });
        }

        } finally {
            // Mark session as no longer in use
            if (activeSession) {
                activeSession.inUse = false;
            }
        }

    } catch (error) {
        console.error('Error during login process:', error);

        // Clean up on error
        if (activeSession && activeSession.automation) {
            try {
                await activeSession.automation.close();
            } catch (closeError) {
                console.error('Error closing automation on error:', closeError);
            }
            activeSession.automation = null;
            activeSession.isPreloaded = false;
            activeSession.inUse = false;
        }

        res.status(500).json({ 
            error: 'Login process failed',
            details: error.message 
        });
    }
});

// Continue endpoint removed - no longer needed for email scanning only

// Take screenshot of current state
app.post('/api/screenshot', async (req, res) => {
    try {
        const { sessionId: requestedSessionId } = req.body;
        const { sessionId, session } = await getOrCreateSession(requestedSessionId);

        if (!session.automation) {
            return res.status(400).json({ 
                error: 'No active automation session' 
            });
        }

        const filename = `screenshots/session-${sessionId}-${Date.now()}.png`;
        await session.automation.takeScreenshot(filename);

        res.json({
            sessionId: sessionId,
            screenshot: filename,
            message: 'Screenshot taken successfully'
        });

    } catch (error) {
        console.error('Error taking screenshot:', error);
        res.status(500).json({ 
            error: 'Failed to take screenshot',
            details: error.message 
        });
    }
});

// Check emails (if logged in)
app.get('/api/emails', async (req, res) => {
    try {
        const { sessionId: requestedSessionId } = req.query;
        const { sessionId, session } = await getOrCreateSession(requestedSessionId);

        if (!session.automation) {
            return res.status(400).json({ 
                error: 'No active automation session' 
            });
        }

        const emails = await session.automation.checkEmails();

        res.json({
            sessionId: sessionId,
            emails,
            count: emails.length
        });

    } catch (error) {
        console.error('Error checking emails:', error);
        res.status(500).json({ 
            error: 'Failed to check emails',
            details: error.message 
        });
    }
});

// New comprehensive email scanning endpoint
app.get('/api/emails/scan-all', async (req, res) => {
    try {
        const { sessionId: requestedSessionId } = req.query;
        const { sessionId, session } = await getOrCreateSession(requestedSessionId);

        if (!session.automation) {
            return res.status(400).json({ 
                error: 'No active automation session' 
            });
        }

        console.log('Starting comprehensive email scan...');
        const allEmails = await session.automation.scanAllEmails();

        const response = {
            sessionId: sessionId,
            data: allEmails,
            summary: {
                inboxCount: allEmails.inbox?.length || 0,
                sentCount: allEmails.sent?.length || 0,
                totalCount: (allEmails.inbox?.length || 0) + (allEmails.sent?.length || 0)
            },
            scanTimestamp: new Date().toISOString()
        };

        if (allEmails.error) {
            response.warning = allEmails.error;
        }

        res.json(response);

    } catch (error) {
        console.error('Error during comprehensive email scan:', error);
        res.status(500).json({ 
            error: 'Failed to scan all emails',
            details: error.message 
        });
    }
});

// Close current session
app.delete('/api/session', async (req, res) => {
    try {
        const { sessionId: requestedSessionId } = req.body;

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
            sessionId: requestedSessionId,
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

// Handle back navigation - reload Outlook page
app.post('/api/back', async (req, res) => {
    try {
        const { sessionId: requestedSessionId } = req.body;
        const { sessionId, session } = await getOrCreateSession(requestedSessionId);

        if (!session.automation) {
            return res.status(400).json({ 
                error: 'No active session' 
            });
        }

        console.log('Back button clicked - auto reloading Outlook page...');

        // Navigate back to Outlook to reset the form
        const reloaded = await session.automation.navigateToOutlook();
        if (!reloaded) {
            return res.status(500).json({ 
                error: 'Failed to reload Outlook page' 
            });
        }

        console.log('Page successfully reloaded after back navigation');

        res.json({
            sessionId: sessionId,
            message: 'Outlook page reloaded successfully',
            status: 'reloaded'
        });

    } catch (error) {
        console.error('Error during back navigation reload:', error);
        res.status(500).json({ 
            error: 'Failed to reload page on back navigation',
            details: error.message 
        });
    }
});

// Get current session status
app.get('/api/status', (req, res) => {
    res.json({
        hasActiveSession: activeSession !== null,
        sessionCount: activeSession ? 1 : 0,
        sessionId: activeSession ? activeSession.sessionId : null
    });
});

// Extend session timeout by making a request
app.post('/api/extend-session', async (req, res) => {
    try {
        const { sessionId: requestedSessionId } = req.body;
        
        if (!activeSession) {
            return res.status(400).json({ error: 'No active session found.' });
        }

        if (requestedSessionId && activeSession.sessionId !== requestedSessionId) {
            return res.status(400).json({ error: 'Provided sessionId does not match active session.' });
        }

        activeSession.createdAt = Date.now();
        console.log(`Session ${activeSession.sessionId} timeout extended.`);
        res.json({ message: 'Session timeout extended.', sessionId: activeSession.sessionId });

    } catch (error) {
        console.error('Error extending session:', error);
        res.status(500).json({ error: 'Failed to extend session.', details: error.message });
    }
});

// Session endpoints removed - no session saving

// Serve frontend
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error('Unhandled error:', err);
    res.status(500).json({ 
        error: 'Internal server error',
        details: err.message 
    });
});

// Graceful shutdown
// Graceful shutdown handling for multiple signals
const gracefulShutdown = async (signal) => {
    console.log(`\n🔄 Received ${signal}. Shutting down server...`);

    // Close active automation session
    if (activeSession) {
        try {
            console.log(`Closing session ${activeSession.sessionId}...`);
            if (activeSession.automation) {
                await Promise.race([
                    activeSession.automation.close(),
                    new Promise((_, reject) => setTimeout(() => reject(new Error('Shutdown timeout')), 5000))
                ]);
            }
        } catch (error) {
            console.error(`Error closing session ${activeSession.sessionId}:`, error);
        }
    }

    activeSession = null;
    console.log('✅ All sessions closed. Server shutdown complete.');
    process.exit(0);
};

process.on('SIGINT', gracefulShutdown);
process.on('SIGTERM', gracefulShutdown);
process.on('SIGQUIT', gracefulShutdown);

// Start server
app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Outlook Automation Backend running on port ${PORT}`);
    console.log(`📧 API endpoints available at http://localhost:${PORT}/api/`);
    console.log(`🌐 Frontend available at http://localhost:${PORT}/`);
});