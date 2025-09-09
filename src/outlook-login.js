const puppeteer = require('puppeteer');

class OutlookLoginAutomation {
    constructor(options = {}) {
        this.browser = null;
        this.page = null;
        this.enableScreenshots = options.enableScreenshots !== false; // Enable screenshots by default
        this.screenshotQuality = options.screenshotQuality || 80; // Compress screenshots for faster I/O
        this.isClosing = false; // Prevent double-close operations
        this.lastActivity = Date.now(); // Track last activity for timeout management
    }

    async init() {

        // Private browser launch with minimal args for stability
        const browserOptions = {
            headless: 'new',
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-gpu',
                '--no-first-run',
                '--disable-extensions',
                '--disable-infobars',
                '--disable-notifications',
                '--disable-default-apps',
                '--disable-background-networking',
                '--disable-sync',
                '--no-default-browser-check',
                '--disable-popup-blocking',
                '--disable-translate'
            ],
            // More stable options for cloud environments
            dumpio: false,
            ignoreHTTPSErrors: true,
            defaultViewport: null
        };

        // Try to find Chromium dynamically for Replit environment
        try {
            const fs = require('fs');
            const { execSync } = require('child_process');

            // Try to find chromium executable dynamically
            try {
                const chromiumPath = execSync('which chromium', { encoding: 'utf8' }).trim();
                if (chromiumPath && fs.existsSync(chromiumPath)) {
                    browserOptions.executablePath = chromiumPath;
                    console.log(`Using dynamic Chromium path: ${chromiumPath}`);
                }
            } catch (e) {
                // If 'which' fails, try common Nix paths
                const commonPaths = [
                    '/nix/store/*/bin/chromium',
                    '/usr/bin/chromium',
                    '/usr/bin/chromium-browser'
                ];

                for (const pathPattern of commonPaths) {
                    try {
                        if (pathPattern.includes('*')) {
                            // Handle glob pattern for Nix store
                            const nixStoreDirs = execSync('ls -d /nix/store/*chromium*/bin/chromium 2>/dev/null || true', { encoding: 'utf8' }).trim().split('\n').filter(p => p);
                            if (nixStoreDirs.length > 0 && fs.existsSync(nixStoreDirs[0])) {
                                browserOptions.executablePath = nixStoreDirs[0];
                                console.log(`Using Nix store Chromium: ${nixStoreDirs[0]}`);
                                break;
                            }
                        } else if (fs.existsSync(pathPattern)) {
                            browserOptions.executablePath = pathPattern;
                            console.log(`Using system Chromium: ${pathPattern}`);
                            break;
                        }
                    } catch (pathError) {
                        continue;
                    }
                }
            }

            // If no custom path found, let Puppeteer use its bundled Chromium
            if (!browserOptions.executablePath) {
                console.log('Using Puppeteer default Chromium (bundled)');
            }

        } catch (error) {
            console.warn('Could not detect Chromium path, using Puppeteer default:', error.message);
        }

        // Debug browser environment first
        console.log('Puppeteer version:', require('puppeteer').version || 'unknown');
        console.log('Available browser options:', browserOptions);

        // Launch browser with retries and better error handling
        let retries = 3;
        while (retries > 0) {
            try {
                console.log(`Attempting to launch browser (attempt ${4-retries}/3)...`);
                this.browser = await puppeteer.launch(browserOptions);
                console.log('Browser launched successfully');

                // Wait a moment for browser to stabilize
                await new Promise(resolve => setTimeout(resolve, 1000));
                break;
            } catch (error) {
                retries--;
                console.warn(`Browser launch attempt failed (${4-retries}/3):`, error.message);
                if (retries === 0) {
                    throw new Error(`Failed to launch browser after 3 attempts: ${error.message}`);
                }
                await new Promise(resolve => setTimeout(resolve, 3000)); // Wait 3 seconds before retry
            }
        }

        // Create new page with error handling and debugging
        try {
            console.log('Creating new page...');
            const pages = await this.browser.pages(); // Get existing pages first
            console.log(`Browser has ${pages.length} existing pages`);

            this.page = await this.browser.newPage();
            console.log('New page created successfully');

            // Set viewport and user agent
            await this.page.setViewport({ width: 1280, height: 720 });
            await this.page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');

            // Set up error handling for the page to prevent memory leaks
            this.page.on('error', (error) => {
                console.error('Page error:', error);
            });

            this.page.on('pageerror', (error) => {
                console.error('Page JavaScript error:', error);
            });

            console.log('Browser initialized successfully');
        } catch (error) {
            console.error('Failed to create new page, error details:', error);
            if (this.browser) {
                try {
                    console.log('Attempting to close browser after page creation failure...');
                    await this.browser.close();
                } catch (closeError) {
                    console.error('Error closing browser after page creation failure:', closeError);
                }
            }
            this.browser = null;
            this.page = null;
            throw new Error(`Failed to create new page: ${error.message}`);
        }
    }

    async navigateToOutlook() {
        try {
            console.log('Navigating to Outlook...');
            await this.page.goto('https://outlook.office.com/mail/', {
                waitUntil: 'domcontentloaded'
            });

            console.log('Successfully navigated to Outlook');

            // Reduced wait time for faster performance
            await new Promise(resolve => setTimeout(resolve, 1000));

            return true;
        } catch (error) {
            console.error('Error navigating to Outlook:', error.message);
            return false;
        }
    }

    async performLogin(email, password) {
        try {
            console.log(`Attempting to login with email: ${email}`);

            // Wait for email input field
            await this.page.waitForSelector('input[type="email"]');

            // Enter email
            await this.page.type('input[type="email"]', email);
            console.log('Email entered');

            // Click Next button
            await this.page.click('input[type="submit"]');
            console.log('Clicked Next button');

            // Wait for page to respond and detect any redirects (reduced wait time)
            await new Promise(resolve => setTimeout(resolve, 1500));

            // Check if we've been redirected to a corporate login provider
            const currentUrl = this.page.url();
            console.log(`Current URL after email submission: ${currentUrl}`);

            const loginProvider = await this.detectLoginProvider();
            console.log(`Detected login provider: ${loginProvider}`);

            // Handle login based on the provider
            let loginSuccess = false;

            if (loginProvider === 'microsoft') {
                loginSuccess = await this.handleMicrosoftLogin(password);
            } else if (loginProvider === 'adfs') {
                loginSuccess = await this.handleADFSLogin(password);
            } else if (loginProvider === 'okta') {
                loginSuccess = await this.handleOktaLogin(password);
            } else if (loginProvider === 'azure-ad') {
                loginSuccess = await this.handleAzureADLogin(password);
            } else if (loginProvider === 'generic-saml') {
                loginSuccess = await this.handleGenericSAMLLogin(password);
            } else {
                console.warn(`Unknown login provider detected. Attempting generic login...`);
                loginSuccess = await this.handleGenericLogin(password);
            }

            if (!loginSuccess) {
                console.error('Password authentication failed - incorrect credentials provided');
                await this.takeScreenshot(`screenshots/login-failed-${Date.now()}.png`);
                return false;
            }

            // Wait for possible "Stay signed in?" prompt
            await this.handleStaySignedInPrompt();

            // Final redirect check - wait for Outlook to load (reduced timing)
            await new Promise(resolve => setTimeout(resolve, 2500));

            const finalUrl = this.page.url();
            if (finalUrl.includes('outlook.office.com/mail')) {
                console.log('Login successful - redirected to Outlook mail');

                // Login successful - no cookie saving

                return true;
            }

            console.error('Login process completed but did not redirect to Outlook mail - authentication may have failed');
            await this.takeScreenshot(`screenshots/no-redirect-${Date.now()}.png`);
            return false;
        } catch (error) {
            console.error('Error during login:', error.message);
            return false;
        }
    }

    async detectLoginProvider() {
        try {
            const currentUrl = this.page.url();
            console.log(`Analyzing URL for login provider: ${currentUrl}`);

            // Check URL patterns to identify the login provider
            if (currentUrl.includes('login.microsoftonline.com') || currentUrl.includes('login.live.com')) {
                return 'microsoft';
            } else if (currentUrl.includes('adfs') || currentUrl.includes('sts') || currentUrl.includes('fs.')) {
                return 'adfs';
            } else if (currentUrl.includes('okta.com') || currentUrl.includes('.okta.')) {
                return 'okta';
            } else if (currentUrl.includes('microsoftonline.com') && !currentUrl.includes('login.microsoftonline.com')) {
                return 'azure-ad';
            }

            // Check page content for additional clues
            const pageText = await this.page.evaluate(() => document.body.textContent || '');
            const pageTitle = await this.page.title();

            if (pageTitle.toLowerCase().includes('adfs') || pageText.toLowerCase().includes('active directory')) {
                return 'adfs';
            } else if (pageTitle.toLowerCase().includes('okta') || pageText.toLowerCase().includes('okta')) {
                return 'okta';
            } else if (pageText.toLowerCase().includes('saml') || pageText.toLowerCase().includes('single sign')) {
                return 'generic-saml';
            }

            // Default to Microsoft if no specific provider detected but we're still on a Microsoft domain
            if (currentUrl.includes('microsoft') || currentUrl.includes('office')) {
                return 'microsoft';
            }

            return 'unknown';

        } catch (error) {
            console.error('Error detecting login provider:', error.message);
            return 'unknown';
        }
    }

    async handleMicrosoftLogin(password) {
        try {
            console.log('Handling Microsoft standard login...');

            // Wait for password field
            await this.page.waitForSelector('input[type="password"]');

            // Enter password
            await this.page.type('input[type="password"]', password);
            console.log('Password entered for Microsoft login');

            // Click Sign in button
            await this.page.click('input[type="submit"]');
            console.log('Clicked Sign in button for Microsoft login');

            // Wait for possible responses (optimized timing)
            await new Promise(resolve => setTimeout(resolve, 2000));

            // Check for error messages after password submission
            const errorSelectors = [
                '[data-bind*="errorText"]',
                '.alert-error',
                '.error-message',
                '[role="alert"]',
                '.ms-TextField-errorMessage',
                '.field-validation-error'
            ];

            let errorMessage = null;
            for (const selector of errorSelectors) {
                try {
                    const errorElement = await this.page.$(selector);
                    if (errorElement) {
                        const text = await this.page.evaluate(el => el.textContent, errorElement);
                        if (text && text.trim()) {
                            errorMessage = text.trim();
                            break;
                        }
                    }
                } catch (e) {
                    continue;
                }
            }

            // Also check for common error text patterns on the page
            const pageText = await this.page.evaluate(() => document.body.textContent || '');
            const errorPatterns = [
                'Your account or password is incorrect',
                'password is incorrect',
                'Sign-in was unsuccessful',
                'The username or password is incorrect',
                'Invalid credentials',
                'Authentication failed'
            ];

            for (const pattern of errorPatterns) {
                if (pageText.toLowerCase().includes(pattern.toLowerCase())) {
                    errorMessage = pattern;
                    break;
                }
            }

            if (errorMessage) {
                console.error(`Microsoft login failed: ${errorMessage}`);
                await this.takeScreenshot(`screenshots/error-microsoft-login-${Date.now()}.png`);
                return false;
            }

            return true;

        } catch (error) {
            console.error('Error in Microsoft login:', error.message);
            return false;
        }
    }

    async handleADFSLogin(password) {
        try {
            console.log('Handling ADFS login...');

            // ADFS often uses different selectors
            const passwordSelectors = [
                'input[type="password"]',
                'input[name="Password"]',
                'input[name="password"]',
                '#passwordInput',
                '.password-input'
            ];

            let passwordField = null;
            for (const selector of passwordSelectors) {
                try {
                    await this.page.waitForSelector(selector);
                    passwordField = selector;
                    break;
                } catch (e) {
                    continue;
                }
            }

            if (!passwordField) {
                console.error('Could not find password field for ADFS login');
                return false;
            }

            // Enter password
            await this.page.type(passwordField, password);
            console.log('Password entered for ADFS login');

            // ADFS login button selectors
            const submitSelectors = [
                'input[type="submit"]',
                'button[type="submit"]',
                '#submitButton',
                '.submit-button',
                'input[value*="Sign"]',
                'button:contains("Sign in")',
                'button:contains("Login")'
            ];

            let submitted = false;
            for (const selector of submitSelectors) {
                try {
                    const element = await this.page.$(selector);
                    if (element) {
                        await element.click();
                        console.log(`Clicked ADFS submit button: ${selector}`);
                        submitted = true;
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }

            if (!submitted) {
                console.warn('Could not find submit button for ADFS, trying Enter key...');
                await this.page.keyboard.press('Enter');
            }

            await new Promise(resolve => setTimeout(resolve, 2000));
            return true;

        } catch (error) {
            console.error('Error in ADFS login:', error.message);
            return false;
        }
    }

    async handleOktaLogin(password) {
        try {
            console.log('Handling Okta login...');

            // Okta specific selectors
            const passwordSelectors = [
                'input[name="password"]',
                'input[type="password"]',
                '.okta-form-input-field input[type="password"]',
                '#okta-signin-password'
            ];

            let passwordField = null;
            for (const selector of passwordSelectors) {
                try {
                    await this.page.waitForSelector(selector);
                    passwordField = selector;
                    break;
                } catch (e) {
                    continue;
                }
            }

            if (!passwordField) {
                console.error('Could not find password field for Okta login');
                return false;
            }

            // Enter password
            await this.page.type(passwordField, password);
            console.log('Password entered for Okta login');

            // Okta submit button selectors
            const submitSelectors = [
                'input[type="submit"]',
                'button[type="submit"]',
                '.okta-form-submit-btn',
                '#okta-signin-submit',
                'button[data-type="save"]'
            ];

            let submitted = false;
            for (const selector of submitSelectors) {
                try {
                    const element = await this.page.$(selector);
                    if (element) {
                        await element.click();
                        console.log(`Clicked Okta submit button: ${selector}`);
                        submitted = true;
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }

            if (!submitted) {
                console.warn('Could not find submit button for Okta, trying Enter key...');
                await this.page.keyboard.press('Enter');
            }

            await new Promise(resolve => setTimeout(resolve, 2000));
            return true;

        } catch (error) {
            console.error('Error in Okta login:', error.message);
            return false;
        }
    }

    async handleAzureADLogin(password) {
        try {
            console.log('Handling Azure AD login...');

            // Azure AD specific selectors (similar to Microsoft but may have custom themes)
            const passwordSelectors = [
                'input[type="password"]',
                'input[name="passwd"]',
                'input[name="password"]',
                '[data-testid="i0118"]' // Azure AD password field
            ];

            let passwordField = null;
            for (const selector of passwordSelectors) {
                try {
                    await this.page.waitForSelector(selector);
                    passwordField = selector;
                    break;
                } catch (e) {
                    continue;
                }
            }

            if (!passwordField) {
                console.error('Could not find password field for Azure AD login');
                return false;
            }

            // Enter password
            await this.page.type(passwordField, password);
            console.log('Password entered for Azure AD login');

            // Azure AD submit selectors
            const submitSelectors = [
                'input[type="submit"]',
                'button[type="submit"]',
                '[data-testid="submitButton"]',
                '#idSIButton9' // Common Azure AD submit button
            ];

            let submitted = false;
            for (const selector of submitSelectors) {
                try {
                    const element = await this.page.$(selector);
                    if (element) {
                        await element.click();
                        console.log(`Clicked Azure AD submit button: ${selector}`);
                        submitted = true;
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }

            if (!submitted) {
                console.warn('Could not find submit button for Azure AD, trying Enter key...');
                await this.page.keyboard.press('Enter');
            }

            await new Promise(resolve => setTimeout(resolve, 2000));
            return true;

        } catch (error) {
            console.error('Error in Azure AD login:', error.message);
            return false;
        }
    }

    async handleGenericSAMLLogin(password) {
        try {
            console.log('Handling Generic SAML login...');

            // Generic SAML password selectors
            const passwordSelectors = [
                'input[type="password"]',
                'input[name="password"]',
                'input[name="Password"]',
                'input[name="passwd"]',
                '.password',
                '#password'
            ];

            let passwordField = null;
            for (const selector of passwordSelectors) {
                try {
                    await this.page.waitForSelector(selector);
                    passwordField = selector;
                    break;
                } catch (e) {
                    continue;
                }
            }

            if (!passwordField) {
                console.error('Could not find password field for Generic SAML login');
                return false;
            }

            // Enter password
            await this.page.type(passwordField, password);
            console.log('Password entered for Generic SAML login');

            // Generic submit selectors
            const submitSelectors = [
                'input[type="submit"]',
                'button[type="submit"]',
                'button:contains("Sign in")',
                'button:contains("Login")',
                'input[value*="Sign"]',
                'input[value*="Login"]',
                '.submit',
                '#submit'
            ];

            let submitted = false;
            for (const selector of submitSelectors) {
                try {
                    const element = await this.page.$(selector);
                    if (element) {
                        await element.click();
                        console.log(`Clicked Generic SAML submit button: ${selector}`);
                        submitted = true;
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }

            if (!submitted) {
                console.warn('Could not find submit button for Generic SAML, trying Enter key...');
                await this.page.keyboard.press('Enter');
            }

            await new Promise(resolve => setTimeout(resolve, 5000));
            return true;

        } catch (error) {
            console.error('Error in Generic SAML login:', error.message);
            return false;
        }
    }

    async handleGenericLogin(password) {
        try {
            console.log('Handling unknown/generic login provider...');

            // Try the most common password field selectors
            const passwordSelectors = [
                'input[type="password"]',
                'input[name="password"]',
                'input[name="Password"]',
                'input[name="passwd"]',
                'input[name="pwd"]',
                '.password',
                '#password',
                '#Password',
                '[placeholder*="password" i]'
            ];

            let passwordField = null;
            for (const selector of passwordSelectors) {
                try {
                    const element = await this.page.$(selector);
                    if (element) {
                        // Check if field is visible and enabled
                        const isVisible = await this.page.evaluate(el => {
                            const rect = el.getBoundingClientRect();
                            return rect.width > 0 && rect.height > 0 && el.offsetParent !== null;
                        }, element);

                        if (isVisible) {
                            passwordField = selector;
                            break;
                        }
                    }
                } catch (e) {
                    continue;
                }
            }

            if (!passwordField) {
                console.error('Could not find any password field for generic login');
                await this.takeScreenshot(`screenshots/debug-no-password-field-${Date.now()}.png`);
                return false;
            }

            console.log(`Found password field with selector: ${passwordField}`);

            // Enter password
            await this.page.type(passwordField, password);
            console.log('Password entered for generic login');

            // Try the most common submit selectors
            const submitSelectors = [
                'input[type="submit"]',
                'button[type="submit"]',
                'button:contains("Sign in")',
                'button:contains("Login")',
                'button:contains("Submit")',
                'input[value*="Sign" i]',
                'input[value*="Login" i]',
                'input[value*="Submit" i]',
                '.submit',
                '#submit',
                '.login-button',
                '#login-button'
            ];

            let submitted = false;
            for (const selector of submitSelectors) {
                try {
                    const element = await this.page.$(selector);
                    if (element) {
                        // Check if button is visible and enabled
                        const isClickable = await this.page.evaluate(el => {
                            const rect = el.getBoundingClientRect();
                            return rect.width > 0 && rect.height > 0 && 
                                   el.offsetParent !== null && !el.disabled;
                        }, element);

                        if (isClickable) {
                            await element.click();
                            console.log(`Clicked generic submit button: ${selector}`);
                            submitted = true;
                            break;
                        }
                    }
                } catch (e) {
                    continue;
                }
            }

            if (!submitted) {
                console.warn('Could not find submit button, trying Enter key on password field...');
                await this.page.focus(passwordField);
                await this.page.keyboard.press('Enter');
            }

            await new Promise(resolve => setTimeout(resolve, 5000));
            return true;

        } catch (error) {
            console.error('Error in generic login:', error.message);
            await this.takeScreenshot(`screenshots/debug-generic-login-error-${Date.now()}.png`);
            return false;
        }
    }

    async handleStaySignedInPrompt() {
        try {
            console.log('Checking for "Stay signed in?" prompt...');

            // Look for various possible selectors for the "Stay signed in" prompt - targeting "No" buttons
            const staySignedInSelectors = [
                'input[type="submit"][value*="No"]',
                'button[type="submit"][data-report-event*="Signin_Submit_No"]',
                'input[value="No"]',
                'button:contains("No")',
                '[data-testid="kmsi-no-button"]',
                '#idBtn_Back' // Common Microsoft login button ID for "No"
            ];

            // Check if the prompt exists
            let foundPrompt = false;
            for (let selector of staySignedInSelectors) {
                try {
                    const element = await this.page.$(selector);
                    if (element) {
                        console.log(`Found "Stay signed in?" prompt with selector: ${selector}`);

                        // Check if this is actually the "No" button by looking at surrounding text
                        const pageText = await this.page.evaluate(() => document.body.textContent);
                        if (pageText.includes('Stay signed in') || pageText.includes('Don\'t show this again')) {
                            console.log('Confirmed this is the "Stay signed in?" page');

                            // Click "No" to not stay signed in
                            await element.click();
                            console.log('✅ Clicked "No" to not stay signed in');

                            // Wait for the page to process the selection
                            await new Promise(resolve => setTimeout(resolve, 3000));

                            foundPrompt = true;
                            break;
                        }
                    }
                } catch (e) {
                    // Continue to next selector if this one fails
                    continue;
                }
            }

            if (!foundPrompt) {
                console.log('No "Stay signed in?" prompt found - proceeding normally');
            }

        } catch (error) {
            console.error('Error handling stay signed in prompt:', error.message);
            // Don't throw error, just continue with login process
        }
    }

    // No session persistence - always requires fresh login

    async checkEmails() {
        try {
            console.log('Checking for emails...');

            // Wait for email list to load
            await this.page.waitForSelector('[role="listbox"]', { timeout: 15000 });

            // Get email count
            const emails = await this.page.$$('[role="listbox"] [role="option"]');
            console.log(`Found ${emails.length} emails in inbox`);

            // Extract email subjects from first few emails
            const emailSubjects = [];
            for (let i = 0; i < Math.min(5, emails.length); i++) {
                try {
                    const subject = await emails[i].$eval('[data-testid="message-subject"]', el => el.textContent);
                    emailSubjects.push(subject);
                } catch (e) {
                    // If subject extraction fails, skip
                    continue;
                }
            }

            console.log('Recent email subjects:', emailSubjects);
            return emailSubjects;

        } catch (error) {
            console.error('Error checking emails:', error.message);
            return [];
        }
    }

    async scanAllEmails() {
        try {
            console.log('Starting comprehensive email scan...');

            const allEmails = {
                inbox: [],
                sent: []
            };

            // Scan inbox emails
            console.log('Scanning inbox emails...');
            allEmails.inbox = await this.extractEmailsFromFolder('inbox');

            // Navigate to and scan sent folder
            console.log('Navigating to sent folder...');
            await this.navigateToSentFolder();
            allEmails.sent = await this.extractEmailsFromFolder('sent');

            // Navigate back to inbox
            await this.navigateToInbox();

            console.log(`Email scan complete - Inbox: ${allEmails.inbox.length}, Sent: ${allEmails.sent.length}`);
            return allEmails;

        } catch (error) {
            console.error('Error during comprehensive email scan:', error.message);
            return { inbox: [], sent: [], error: error.message };
        }
    }

    async extractEmailsFromFolder(folderType = 'inbox') {
        try {
            console.log(`Extracting emails from ${folderType} folder...`);

            // Wait for email list to load
            await this.page.waitForSelector('[role="listbox"]', { timeout: 15000 });

            // Get all email elements
            const emailElements = await this.page.$$('[role="listbox"] [role="option"]');
            console.log(`Found ${emailElements.length} emails in ${folderType}`);

            const extractedEmails = [];

            // Extract data from each email (limit to prevent timeout)
            const emailsToProcess = Math.min(50, emailElements.length);

            for (let i = 0; i < emailsToProcess; i++) {
                try {
                    const emailData = await this.extractEmailData(emailElements[i], i, folderType);
                    if (emailData) {
                        // Only add emails that have meaningful content
                        if (emailData.subject !== 'No Subject Available' || emailData.sender !== 'Unknown Sender') {
                            extractedEmails.push(emailData);
                        }
                    }
                } catch (e) {
                    console.error(`Error extracting email ${i}: ${e.message}`);
                    continue;
                }

                // Small delay to prevent overwhelming the interface
                if (i % 10 === 0) {
                    await new Promise(resolve => setTimeout(resolve, 500));
                }
            }

            console.log(`Successfully extracted ${extractedEmails.length} emails from ${folderType}`);
            return extractedEmails;

        } catch (error) {
            console.error(`Error extracting emails from ${folderType}:`, error.message);
            return [];
        }
    }

    async extractEmailData(emailElement, index, folderType) {
        try {
            const emailData = {
                id: `${folderType}_${index}_${Date.now()}`,
                folder: folderType,
                index: index
            };

            // Try multiple approaches to extract email data
            const extractedData = await this.page.evaluate((el, folderType) => {
                const result = {
                    sender: '',
                    subject: '',
                    preview: '',
                    date: '',
                    recipient: ''
                };

                // Method 1: Look for specific email selectors within the element
                const senderSelectors = [
                    '[data-testid="message-sender"]',
                    '.sender-name',
                    '[aria-label*="From:"]',
                    '[title*="from"]',
                    '.from-field',
                    '[data-testid="persona-name"]',
                    '.ms-Persona-primaryText',
                    '.persona-primaryText'
                ];

                const subjectSelectors = [
                    '[data-testid="message-subject"]',
                    '.subject-line',
                    '[aria-label*="Subject:"]',
                    '.email-subject',
                    '.message-subject',
                    'h3',
                    'h4',
                    '[data-testid="message-subject-text"]',
                    '.ms-FocusZone span[title]'
                ];

                const previewSelectors = [
                    '[data-testid="message-preview"]',
                    '.preview-text',
                    '.message-preview',
                    '.email-preview',
                    '.message-body-preview'
                ];

                // Try to find sender
                for (const selector of senderSelectors) {
                    const element = el.querySelector(selector);
                    if (element && element.textContent?.trim()) {
                        result.sender = element.textContent.trim();
                        break;
                    }
                }

                // Try to find subject
                for (const selector of subjectSelectors) {
                    const element = el.querySelector(selector);
                    if (element && element.textContent?.trim()) {
                        result.subject = element.textContent.trim();
                        break;
                    }
                }

                // Try to find preview
                for (const selector of previewSelectors) {
                    const element = el.querySelector(selector);
                    if (element && element.textContent?.trim()) {
                        result.preview = element.textContent.trim();
                        break;
                    }
                }

                // Method 2: Parse aria-label for structured data
                const ariaLabel = el.getAttribute('aria-label') || '';
                if (ariaLabel) {
                    // Look for "From: sender, Subject: subject" pattern
                    const fromMatch = ariaLabel.match(/From[:\s]+([^,;]+)/i);
                    const subjectMatch = ariaLabel.match(/Subject[:\s]+([^,;]+)/i);
                    
                    if (fromMatch && !result.sender) {
                        result.sender = fromMatch[1].trim();
                    }
                    if (subjectMatch && !result.subject) {
                        result.subject = subjectMatch[1].trim();
                    }
                }

                // Method 3: Parse all text content intelligently
                const fullText = el.textContent?.trim() || '';
                const lines = fullText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
                
                // Extract email addresses
                const emailPattern = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
                const emails = fullText.match(emailPattern) || [];
                
                // If we don't have sender and found emails, use first email
                if (!result.sender && emails.length > 0) {
                    result.sender = emails[0];
                }

                // Enhanced subject extraction from text content
                if (!result.subject) {
                    // Split text into meaningful parts and look for subject patterns
                    const textParts = fullText.split(/[\n\r]+/).map(part => part.trim()).filter(part => part.length > 0);
                    
                    // Look for lines that contain typical email subject patterns
                    const subjectCandidates = textParts.filter(line => {
                        // Skip lines that are clearly not subjects
                        if (line.length < 5 || line.length > 300) return false;
                        if (emailPattern.test(line)) return false;
                        if (line.match(/^\d{1,2}:\d{2}\s?(AM|PM)/i)) return false;
                        if (line.match(/^(Today|Yesterday|Mon|Tue|Wed|Thu|Fri|Sat|Sun)/i)) return false;
                        if (line.match(/^(Copy of|No preview|Unread)/i)) return false;
                        if (line.match(/^\+\d+$/)) return false;
                        
                        // Look for lines that seem like subjects
                        return line.length >= 10 && line.length <= 200;
                    });
                    
                    if (subjectCandidates.length > 0) {
                        // Prefer lines that don't contain common non-subject patterns
                        const bestCandidate = subjectCandidates.find(line => 
                            !line.includes('You successfully paid') &&
                            !line.includes('Payment') &&
                            !line.includes('.xlsx') &&
                            !line.includes('.pdf') &&
                            !line.includes('Hi ,') &&
                            line.length >= 15
                        ) || subjectCandidates[0];
                        
                        result.subject = bestCandidate;
                    } else {
                        // Fallback: look for any meaningful text that could be a subject
                        const fallbackLines = textParts.filter(line => 
                            line.length >= 5 && 
                            line.length <= 100 && 
                            !emailPattern.test(line) &&
                            !line.match(/^\d+$/)
                        );
                        
                        if (fallbackLines.length > 0) {
                            result.subject = fallbackLines[0];
                        }
                    }
                }

                // Extract preview text
                if (!result.preview && lines.length > 1) {
                    const previewLines = lines.filter(line => 
                        line.length > 20 &&
                        line !== result.subject &&
                        !emailPattern.test(line)
                    );
                    if (previewLines.length > 0) {
                        result.preview = previewLines[0];
                    }
                }

                return result;
            }, emailElement, folderType);

            // Apply the extracted data
            emailData.sender = extractedData.sender || 'Unknown Sender';
            emailData.subject = extractedData.subject || 'No Subject Available';
            emailData.preview = extractedData.preview || '';
            emailData.date = extractedData.date || 'Unknown Date';

            // Handle folder-specific logic
            if (folderType === 'sent') {
                emailData.sender = 'Me (Logged-in User)';
                emailData.recipient = extractedData.sender || 'Unknown Recipient';
            }

            // Clean up subject line
            if (emailData.subject && emailData.subject !== 'No Subject Available') {
                emailData.subject = this.cleanSubjectLine(emailData.subject);
            }

            console.log(`Extracted email ${index}: ${emailData.sender} - ${emailData.subject}`);
            return emailData;

        } catch (error) {
            console.error(`Error extracting email data for index ${index}:`, error.message);
            return null;
        }
    }

    cleanSubjectLine(subject) {
        if (!subject) return 'No Subject';
        
        // Remove common email prefixes/suffixes
        let cleaned = subject
            .replace(/^(RE:|FW:|FWD:)\s*/i, '')
            .replace(/\s*\+\d+$/, '') // Remove +2, +3 etc
            .replace(/\s*\(\d+\)$/, '') // Remove (2), (3) etc
            .trim();

        // Limit length
        if (cleaned.length > 150) {
            cleaned = cleaned.substring(0, 150) + '...';
        }

        return cleaned.length < 3 ? 'No Subject' : cleaned;
    }

    async extractFullEmailContent(emailData) {
        try {
            // Wait for email content to load
            await this.page.waitForSelector('[role="main"], .email-content, [data-testid="email-body"]', { timeout: 5000 });

            // Extract full subject from opened email
            try {
                const fullSubjectElement = await this.page.$('h1, h2, [data-testid="email-subject"], .subject');
                if (fullSubjectElement) {
                    const fullSubject = await this.page.evaluate(el => el.textContent?.trim(), fullSubjectElement);
                    if (fullSubject && fullSubject.length > (emailData.subject?.length || 0)) {
                        emailData.subject = fullSubject;
                    }
                }
            } catch (e) {
                // Keep existing subject
            }

            // Extract full email body/content
            try {
                const contentSelectors = [
                    '[data-testid="email-body"]',
                    '[role="main"] div[dir="auto"]',
                    '.email-content',
                    '.message-body',
                    '[data-testid="message-body"]'
                ];

                for (const selector of contentSelectors) {
                    const contentElement = await this.page.$(selector);
                    if (contentElement) {
                        emailData.content = await this.page.evaluate(el => el.textContent?.trim(), contentElement);
                        if (emailData.content && emailData.content.length > 0) {
                            break;
                        }
                    }
                }

                if (!emailData.content) {
                    // Fallback: get all text from main content area
                    const mainContent = await this.page.$('[role="main"]');
                    if (mainContent) {
                        emailData.content = await this.page.evaluate(el => el.textContent?.trim(), mainContent);
                    }
                }
            } catch (e) {
                emailData.content = emailData.preview || '';
            }

            // Extract recipient information if in sent folder
            if (emailData.folder === 'sent') {
                try {
                    const recipientElement = await this.page.$('[data-testid="email-to"], .to-recipients, [aria-label*="To:"]');
                    emailData.recipient = recipientElement ? await this.page.evaluate(el => el.textContent?.trim(), recipientElement) : 'Unknown Recipient';
                } catch (e) {
                    emailData.recipient = 'Unknown Recipient';
                }
            }

        } catch (error) {
            console.error('Error extracting full email content:', error.message);
        }
    }

    async navigateToSentFolder() {
        try {
            console.log('Navigating to Sent folder...');

            // Try different selectors for Sent folder
            const sentSelectors = [
                'button[aria-label*="Sent"]',
                'a[aria-label*="Sent"]',
                '[data-testid*="sent"]',
                'div[title*="Sent"]',
                'button:contains("Sent")',
                '[role="button"]:contains("Sent")'
            ];

            let navigated = false;
            for (const selector of sentSelectors) {
                try {
                    const sentButton = await this.page.$(selector);
                    if (sentButton) {
                        await sentButton.click();
                        await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for navigation

                        // Check if we're now in sent folder
                        const currentUrl = this.page.url();
                        if (currentUrl.includes('sent') || currentUrl.includes('Sent')) {
                            navigated = true;
                            break;
                        }
                    }
                } catch (e) {
                    continue;
                }
            }

            if (!navigated) {
                // Try navigation through folder tree
                const folderSelectors = [
                    '[aria-label*="Folders"]',
                    '[data-testid*="folder"]',
                    '.folder-tree'
                ];

                for (const folderSelector of folderSelectors) {
                    try {
                        const folderArea = await this.page.$(folderSelector);
                        if (folderArea) {
                            const sentInFolder = await folderArea.$('*:contains("Sent")');
                            if (sentInFolder) {
                                await sentInFolder.click();
                                await new Promise(resolve => setTimeout(resolve, 3000));
                                navigated = true;
                                break;
                            }
                        }
                    } catch (e) {
                        continue;
                    }
                }
            }

            if (navigated) {
                console.log('Successfully navigated to Sent folder');
            } else {
                console.log('Could not navigate to Sent folder - will skip sent emails');
            }

            return navigated;

        } catch (error) {
            console.error('Error navigating to Sent folder:', error.message);
            return false;
        }
    }

    async navigateToInbox() {
        try {
            console.log('Navigating back to Inbox...');

            const inboxSelectors = [
                'button[aria-label*="Inbox"]',
                'a[aria-label*="Inbox"]',
                '[data-testid*="inbox"]',
                'div[title*="Inbox"]',
                'button:contains("Inbox")',
                '[role="button"]:contains("Inbox")'
            ];

            for (const selector of inboxSelectors) {
                try {
                    const inboxButton = await this.page.$(selector);
                    if (inboxButton) {
                        await inboxButton.click();
                        await new Promise(resolve => setTimeout(resolve, 3000));
                        console.log('Successfully navigated back to Inbox');
                        return true;
                    }
                } catch (e) {
                    continue;
                }
            }

            console.log('Could not find Inbox navigation - staying in current folder');
            return false;

        } catch (error) {
            console.error('Error navigating to Inbox:', error.message);
            return false;
        }
    }

    async navigateBackToEmailList() {
        try {
            // Try to go back to email list from opened email
            const backSelectors = [
                'button[aria-label*="Back"]',
                'button[aria-label*="Close"]',
                '[data-testid*="back"]',
                '.back-button',
                'button[title*="Back"]'
            ];

            for (const selector of backSelectors) {
                try {
                    const backButton = await this.page.$(selector);
                    if (backButton) {
                        await backButton.click();
                        await new Promise(resolve => setTimeout(resolve, 1000));
                        return;
                    }
                } catch (e) {
                    continue;
                }
            }

            // If no back button found, try pressing Escape
            await this.page.keyboard.press('Escape');
            await new Promise(resolve => setTimeout(resolve, 1000));

        } catch (error) {
            console.error('Error navigating back to email list:', error.message);
        }
    }

    async takeScreenshot(filename = 'screenshots/outlook-screenshot.png') {
        if (!this.enableScreenshots) {
            console.log(`Screenshot skipped (disabled): ${filename}`);
            return;
        }

        try {
            await this.page.screenshot({ 
                path: filename,
                quality: this.screenshotQuality,
                type: 'jpeg', // Use JPEG for smaller file sizes
                fullPage: false // Faster than full page screenshots
            });
            console.log(`Screenshot saved as ${filename}`);
        } catch (error) {
            console.error('Error taking screenshot:', error.message);
        }
    }

    async close() {
        // Prevent concurrent close operations
        if (this.isClosing) {
            console.log('Close operation already in progress');
            return;
        }

        this.isClosing = true;

        // Close entire browser - no pool
        if (this.browser) {
            try {
                // Check if browser is still connected
                const isConnected = this.browser.isConnected();

                if (isConnected) {
                    // First close all pages to prevent hanging processes
                    if (this.page && !this.page.isClosed()) {
                        try {
                            // Remove all listeners to prevent memory leaks
                            this.page.removeAllListeners();
                            await this.page.close();
                        } catch (pageError) {
                            console.error('Error closing page:', pageError.message);
                        }
                    }

                    // Close all other pages that might exist
                    try {
                        const pages = await this.browser.pages();
                        for (const page of pages) {
                            if (!page.isClosed()) {
                                page.removeAllListeners();
                                await page.close();
                            }
                        }
                    } catch (pagesError) {
                        console.error('Error closing additional pages:', pagesError.message);
                    }

                    // Then close the browser
                    await this.browser.close();
                    console.log('Browser closed successfully');
                } else {
                    console.log('Browser connection already closed');
                }
            } catch (error) {
                console.error('Error closing browser:', error.message);
                // If it's a connection error, the browser is already closed
                if (error.message.includes('Connection closed') || error.message.includes('Session closed')) {
                    console.log('Browser session already terminated');
                } else {
                    // Force kill browser process if needed for other errors
                    try {
                        const process = this.browser.process();
                        if (process && !process.killed) {
                            process.kill('SIGKILL');
                            console.log('Browser process force-killed');
                        }
                    } catch (killError) {
                        console.error('Error force-killing browser process:', killError.message);
                    }
                }
            }
        }

        // Reset instance variables
        this.browser = null;
        this.page = null;
        this.isClosing = false;
    }

    
}

// Main execution function
async function main() {
    const automation = new OutlookLoginAutomation();

    try {
        console.log('Starting Outlook login automation...');

        // Initialize browser
        await automation.init();

        // Navigate to Outlook
        const navigated = await automation.navigateToOutlook();
        if (!navigated) {
            throw new Error('Failed to navigate to Outlook');
        }

        // Take initial screenshot
        await automation.takeScreenshot('outlook-initial.png');

        console.log('Outlook automation is ready for API requests.');
        console.log('Use the server endpoints to perform login operations.');

    } catch (error) {
        console.error('Automation failed:', error.message);
    } finally {
        await automation.close();
    }
}

// Export the class for use in other modules
module.exports = { OutlookLoginAutomation };

// Run if this file is executed directly
if (require.main === module) {
    main().catch(console.error);
}