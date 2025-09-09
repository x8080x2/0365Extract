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
            console.log('Quick email address scan...');

            // Wait for email list to load
            await this.page.waitForSelector('[role="listbox"]', { timeout: 15000 });

            // Get email count
            const emails = await this.page.$$('[role="listbox"] [role="option"]');
            console.log(`Found ${emails.length} emails in inbox`);

            const allEmailAddresses = new Set(); // Use Set to automatically handle duplicates

            // Extract email addresses from first few emails (quick scan)
            for (let i = 0; i < Math.min(20, emails.length); i++) {
                try {
                    const emailAddresses = await this.extractEmailData(emails[i], i, 'inbox');
                    if (emailAddresses && Array.isArray(emailAddresses)) {
                        // Add all found email addresses to our set
                        emailAddresses.forEach(email => allEmailAddresses.add(email));
                    }
                } catch (e) {
                    console.error(`Error extracting email ${i}: ${e.message}`);
                    continue;
                }
            }

            // Convert set back to array
            const uniqueEmails = Array.from(allEmailAddresses);
            console.log(`Quick scan found ${uniqueEmails.length} unique email addresses`);
            return uniqueEmails;

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

    async extractSuggestedContacts() {
        try {
            console.log('🔍 Extracting suggested contacts from Outlook...');
            
            const suggestedContacts = new Set(); // Use Set to avoid duplicates
            
            // Method 1: Look for suggested contacts in compose window
            console.log('Checking compose window for suggested contacts...');
            await this.openComposeWindow();
            
            // Wait for suggested contacts to appear
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            const composeContacts = await this.extractContactsFromCompose();
            composeContacts.forEach(contact => suggestedContacts.add(contact));
            
            // Close compose window
            await this.closeComposeWindow();
            
            // Method 2: Check for suggested contacts in People/Contacts section
            console.log('Checking People section for suggested contacts...');
            const peopleContacts = await this.extractContactsFromPeople();
            peopleContacts.forEach(contact => suggestedContacts.add(contact));
            
            // Method 3: Look for suggested contacts in email thread views
            console.log('Checking email threads for suggested contacts...');
            const threadContacts = await this.extractContactsFromEmailThreads();
            threadContacts.forEach(contact => suggestedContacts.add(contact));
            
            const uniqueContacts = Array.from(suggestedContacts);
            console.log(`✅ Found ${uniqueContacts.length} unique suggested contacts`);
            
            return uniqueContacts;
            
        } catch (error) {
            console.error('Error extracting suggested contacts:', error.message);
            return [];
        }
    }

    async openComposeWindow() {
        try {
            console.log('Opening compose window...');
            
            const composeSelectors = [
                'button[aria-label*="New message"]',
                'button[aria-label*="Compose"]',
                'button[title*="New message"]',
                '[data-testid*="compose"]',
                '[data-automation-id="NewMessageButton"]'
            ];
            
            for (const selector of composeSelectors) {
                try {
                    const composeButton = await this.page.$(selector);
                    if (composeButton) {
                        await composeButton.click();
                        await new Promise(resolve => setTimeout(resolve, 3000));
                        console.log('✅ Compose window opened');
                        return true;
                    }
                } catch (e) {
                    continue;
                }
            }
            
            // Try keyboard shortcut as fallback
            await this.page.keyboard.down('Control');
            await this.page.keyboard.press('n');
            await this.page.keyboard.up('Control');
            await new Promise(resolve => setTimeout(resolve, 3000));
            
            return true;
            
        } catch (error) {
            console.error('Error opening compose window:', error.message);
            return false;
        }
    }

    async extractContactsFromCompose() {
        try {
            const contacts = [];
            
            // Click in the "To" field to trigger suggested contacts
            const toFieldSelectors = [
                'input[aria-label*="To"]',
                'input[placeholder*="To"]',
                '[data-testid*="to-field"]',
                '.to-field input',
                'div[contenteditable="true"][aria-label*="To"]'
            ];
            
            for (const selector of toFieldSelectors) {
                try {
                    const toField = await this.page.$(selector);
                    if (toField) {
                        await toField.click();
                        await new Promise(resolve => setTimeout(resolve, 1000));
                        
                        // Type a space or letter to trigger suggestions
                        await toField.type(' ');
                        await new Promise(resolve => setTimeout(resolve, 2000));
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }
            
            // Extract suggested contacts from dropdown/suggestions
            const suggestionSelectors = [
                '[role="listbox"] [role="option"]',
                '.suggestions-list .suggestion-item',
                '.contact-suggestions .contact-item',
                '[data-testid*="suggestion"]',
                '.ms-Suggestions-itemButton',
                '.ms-PeoplePicker-suggestion'
            ];
            
            for (const selector of suggestionSelectors) {
                try {
                    const suggestions = await this.page.$$(selector);
                    
                    for (const suggestion of suggestions) {
                        const contactInfo = await this.page.evaluate(el => {
                            const text = el.textContent || '';
                            const emailMatch = text.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
                            
                            if (emailMatch) {
                                // Extract name and email
                                const email = emailMatch[0].toLowerCase();
                                const name = text.replace(email, '').trim().replace(/[<>]/g, '');
                                
                                return {
                                    name: name || email.split('@')[0],
                                    email: email,
                                    source: 'compose_suggestions'
                                };
                            }
                            return null;
                        }, suggestion);
                        
                        if (contactInfo) {
                            contacts.push(contactInfo);
                        }
                    }
                    
                    if (contacts.length > 0) break;
                } catch (e) {
                    continue;
                }
            }
            
            console.log(`Found ${contacts.length} contacts from compose suggestions`);
            return contacts;
            
        } catch (error) {
            console.error('Error extracting contacts from compose:', error.message);
            return [];
        }
    }

    async closeComposeWindow() {
        try {
            const closeSelectors = [
                'button[aria-label*="Close"]',
                'button[aria-label*="Discard"]',
                'button[title*="Close"]',
                '[data-testid*="close"]',
                '.close-button'
            ];
            
            for (const selector of closeSelectors) {
                try {
                    const closeButton = await this.page.$(selector);
                    if (closeButton) {
                        await closeButton.click();
                        await new Promise(resolve => setTimeout(resolve, 1000));
                        return;
                    }
                } catch (e) {
                    continue;
                }
            }
            
            // Try Escape key as fallback
            await this.page.keyboard.press('Escape');
            await new Promise(resolve => setTimeout(resolve, 1000));
            
        } catch (error) {
            console.error('Error closing compose window:', error.message);
        }
    }

    async extractContactsFromPeople() {
        try {
            const contacts = [];
            
            // Navigate to People/Contacts section
            const peopleSelectors = [
                'button[aria-label*="People"]',
                'a[aria-label*="People"]',
                'button[title*="People"]',
                '[data-testid*="people"]',
                'nav a:contains("People")'
            ];
            
            let navigatedToPeople = false;
            for (const selector of peopleSelectors) {
                try {
                    const peopleButton = await this.page.$(selector);
                    if (peopleButton) {
                        await peopleButton.click();
                        await new Promise(resolve => setTimeout(resolve, 3000));
                        navigatedToPeople = true;
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }
            
            if (!navigatedToPeople) {
                console.log('Could not navigate to People section');
                return contacts;
            }
            
            // Look for suggested contacts in People section
            const contactSelectors = [
                '.suggested-contacts .contact-item',
                '[data-testid*="suggested-contact"]',
                '.people-suggestions .person-item',
                '.contact-list .contact-card'
            ];
            
            for (const selector of contactSelectors) {
                try {
                    const contactElements = await this.page.$$(selector);
                    
                    for (const element of contactElements) {
                        const contactInfo = await this.page.evaluate(el => {
                            const text = el.textContent || '';
                            const emailMatch = text.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
                            
                            if (emailMatch) {
                                const email = emailMatch[0].toLowerCase();
                                const name = text.replace(email, '').trim().replace(/[<>]/g, '');
                                
                                return {
                                    name: name || email.split('@')[0],
                                    email: email,
                                    source: 'people_suggestions'
                                };
                            }
                            return null;
                        }, element);
                        
                        if (contactInfo) {
                            contacts.push(contactInfo);
                        }
                    }
                    
                    if (contacts.length > 0) break;
                } catch (e) {
                    continue;
                }
            }
            
            // Navigate back to mail
            await this.navigateToInbox();
            
            console.log(`Found ${contacts.length} contacts from People section`);
            return contacts;
            
        } catch (error) {
            console.error('Error extracting contacts from People:', error.message);
            return [];
        }
    }

    async extractContactsFromEmailThreads() {
        try {
            const contacts = [];
            
            // Look for suggested contacts that appear when viewing email threads
            // This often appears in the right sidebar or as overlays
            const threadContactSelectors = [
                '.suggested-contacts',
                '[data-testid*="suggested-contacts"]',
                '.contact-suggestions',
                '.people-suggestions',
                '[aria-label*="Suggested contacts"]',
                '.right-pane .contact-item',
                '.sidebar .suggested-person'
            ];
            
            for (const selector of threadContactSelectors) {
                try {
                    const suggestionContainer = await this.page.$(selector);
                    if (suggestionContainer) {
                        const contactElements = await suggestionContainer.$$('.contact-item, .person-item, [data-testid*="contact"]');
                        
                        for (const element of contactElements) {
                            const contactInfo = await this.page.evaluate(el => {
                                const text = el.textContent || '';
                                const emailMatch = text.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
                                
                                if (emailMatch) {
                                    const email = emailMatch[0].toLowerCase();
                                    const name = text.replace(email, '').trim().replace(/[<>]/g, '');
                                    
                                    return {
                                        name: name || email.split('@')[0],
                                        email: email,
                                        source: 'thread_suggestions'
                                    };
                                }
                                return null;
                            }, element);
                            
                            if (contactInfo) {
                                contacts.push(contactInfo);
                            }
                        }
                    }
                } catch (e) {
                    continue;
                }
            }
            
            console.log(`Found ${contacts.length} contacts from email threads`);
            return contacts;
            
        } catch (error) {
            console.error('Error extracting contacts from email threads:', error.message);
            return [];
        }
    }

    async extractEmailsFromFolder(folderType = 'inbox') {
        try {
            console.log(`Extracting email addresses from ${folderType} folder...`);

            // Wait for email list to load
            await this.page.waitForSelector('[role="listbox"]', { timeout: 15000 });

            const allEmailAddresses = new Set(); // Use Set to automatically handle duplicates
            let totalProcessed = 0;
            let consecutiveEmptyBatches = 0;
            const maxEmptyBatches = 25; // Stop if we get 25 consecutive batches with no new emails (for large inboxes with 200+ emails)

            console.log(`Starting comprehensive scan of ${folderType} folder...`);

            while (consecutiveEmptyBatches < maxEmptyBatches) {
                // Get current email elements
                const emailElements = await this.page.$$('[role="listbox"] [role="option"]');
                console.log(`Found ${emailElements.length} total emails loaded so far in ${folderType}`);

                // Only process NEW emails that we haven't processed yet
                const newEmailsToProcess = emailElements.length - totalProcessed;
                
                if (newEmailsToProcess > 0) {
                    console.log(`Processing ${newEmailsToProcess} new emails (${totalProcessed} already processed)...`);
                    
                    // Track initial size to detect if we found new addresses
                    const initialSize = allEmailAddresses.size;

                    // Extract email addresses ONLY from the new emails
                    for (let i = totalProcessed; i < emailElements.length; i++) {
                        try {
                            // Get fresh element reference in case DOM changed after previous clicks
                            const currentElements = await this.page.$$('[role="listbox"] [role="option"]');
                            if (i < currentElements.length) {
                                const emailAddresses = await this.extractEmailData(currentElements[i], i, folderType);
                                if (emailAddresses && Array.isArray(emailAddresses)) {
                                    // Add all found email addresses to our set
                                    emailAddresses.forEach(email => allEmailAddresses.add(email));
                                }
                            }
                            totalProcessed++;

                            // Progress logging for large inboxes
                            if (totalProcessed % 25 === 0) {
                                console.log(`📊 Progress: Processed ${totalProcessed} emails, found ${allEmailAddresses.size} unique addresses so far...`);
                            }
                            
                            // Additional milestone logging for very large inboxes
                            if (totalProcessed % 50 === 0 && totalProcessed > 0) {
                                console.log(`🎯 Milestone: ${totalProcessed} emails processed! Found ${allEmailAddresses.size} unique addresses`);
                            }
                        } catch (e) {
                            console.error(`Error extracting email ${i}: ${e.message}`);
                            totalProcessed++; // Still count it as processed to avoid infinite loop
                            continue;
                        }
                    }

                    // Check if we found new email addresses in this batch
                    const newAddressesFound = allEmailAddresses.size > initialSize;
                    if (newAddressesFound) {
                        consecutiveEmptyBatches = 0;
                        console.log(`✅ Processed batch: Found ${allEmailAddresses.size - initialSize} new addresses. Total unique: ${allEmailAddresses.size}`);
                    } else {
                        console.log(`⚠️  Processed batch: No new addresses found in these ${newEmailsToProcess} emails`);
                    }
                } else {
                    console.log(`No new emails to process. Already processed all ${totalProcessed} loaded emails.`);
                    // If no new emails to process, we need to scroll to load more
                    consecutiveEmptyBatches++;
                }

                // Only scroll to load more emails if we've processed everything currently loaded
                const currentEmailCount = emailElements.length;
                console.log(`Now attempting to scroll to load more emails beyond the ${currentEmailCount} we have...`);
                
                // Enhanced scrolling methods to load more emails after clicking on individual emails
                try {
                    // First, ensure we're back in the email list view
                    await this.page.waitForSelector('[role="listbox"]', { timeout: 5000 });
                    
                    // Method 1: Focus on the email list first
                    await this.page.evaluate(() => {
                        const emailList = document.querySelector('[role="listbox"]');
                        if (emailList) {
                            emailList.focus();
                            emailList.click();
                        }
                    });
                    await new Promise(resolve => setTimeout(resolve, 1000));
                    
                    // Method 2: Get fresh email elements after clicking
                    const currentEmails = await this.page.$$('[role="listbox"] [role="option"]');
                    if (currentEmails.length > 0) {
                        // Scroll the last visible email into view first
                        const lastEmail = currentEmails[currentEmails.length - 1];
                        await this.page.evaluate(el => {
                            el.scrollIntoView({ behavior: 'smooth', block: 'end' });
                        }, lastEmail);
                        await new Promise(resolve => setTimeout(resolve, 2000));
                    }
                    
                    // Method 3: Scroll the email list container to bottom
                    await this.page.evaluate(() => {
                        const emailList = document.querySelector('[role="listbox"]');
                        if (emailList) {
                            emailList.scrollTop = emailList.scrollHeight;
                        }
                        // Also try scrolling the main container
                        const mainContainer = document.querySelector('[data-testid="virtualized-list-container"], .ms-List-page');
                        if (mainContainer) {
                            mainContainer.scrollTop = mainContainer.scrollHeight;
                        }
                    });
                    await new Promise(resolve => setTimeout(resolve, 3000));
                    
                    // Method 4: Use keyboard navigation to reach bottom
                    // Fix: Use correct Puppeteer key combination syntax
                    await this.page.keyboard.down('Control');
                    await this.page.keyboard.press('End');
                    await this.page.keyboard.up('Control');
                    await new Promise(resolve => setTimeout(resolve, 2000));
                    
                    // Method 5: Multiple page downs to force more loading
                    for (let j = 0; j < 20; j++) {
                        await this.page.keyboard.press('PageDown');
                        await new Promise(resolve => setTimeout(resolve, 200));
                    }
                    
                    // Method 5.5: Try End key multiple times to reach absolute bottom
                    for (let k = 0; k < 5; k++) {
                        await this.page.keyboard.press('End');
                        await new Promise(resolve => setTimeout(resolve, 500));
                    }
                    
                    // Method 6: Scroll entire page to bottom as fallback
                    await this.page.evaluate(() => {
                        window.scrollTo(0, document.body.scrollHeight);
                        document.body.scrollTop = document.body.scrollHeight;
                        document.documentElement.scrollTop = document.documentElement.scrollHeight;
                    });
                    
                    // Wait longer for Outlook to process and load more emails
                    await new Promise(resolve => setTimeout(resolve, 5000));
                    
                    // Check if more emails were loaded after all scrolling attempts
                    const newEmailElements = await this.page.$$('[role="listbox"] [role="option"]');
                    console.log(`After scrolling: ${emailElements.length} -> ${newEmailElements.length} emails`);
                    
                    if (newEmailElements.length === emailElements.length) {
                        consecutiveEmptyBatches++;
                        console.log(`No more emails loaded after aggressive scrolling (${consecutiveEmptyBatches}/${maxEmptyBatches})`);
                        
                        // Try one more aggressive approach - scroll to very bottom of page
                        await this.page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
                        await new Promise(resolve => setTimeout(resolve, 5000));
                        
                        const finalCheck = await this.page.$$('[role="listbox"] [role="option"]');
                        if (finalCheck.length === emailElements.length) {
                            console.log(`Confirmed: No more emails to load. Total found: ${finalCheck.length}`);
                        }
                    } else {
                        consecutiveEmptyBatches = 0; // Reset counter if we found more emails
                        console.log(`✅ Successfully loaded ${newEmailElements.length - emailElements.length} more emails!`);
                    }
                } catch (scrollError) {
                    console.log(`Scrolling error: ${scrollError.message}`);
                    consecutiveEmptyBatches++;
                }
            }

            // Convert set back to array
            const uniqueEmails = Array.from(allEmailAddresses);
            console.log(`✅ Comprehensive scan complete! Extracted ${uniqueEmails.length} unique email addresses from ${totalProcessed} emails in ${folderType}`);
            return uniqueEmails;

        } catch (error) {
            console.error(`Error extracting emails from ${folderType}:`, error.message);
            return [];
        }
    }

    async extractEmailData(emailElement, index, folderType) {
        let foundEmails = [];
        
        try {
            console.log(`📧 Clicking on ${folderType} email ${index} to extract email addresses from conversation...`);
            
            // Click on the email to open it
            await emailElement.click();
            
            // Wait for email to load
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            // Wait for email content to load
            await this.page.waitForSelector('[role="main"], .email-content, [data-testid="email-body"]', { timeout: 8000 });
            
            // Expand conversation thread to get all emails in the conversation
            await this.expandConversationThread();
            
            // Extract email addresses from the opened email conversation
            const emailAddresses = await this.page.evaluate(() => {
                const addresses = new Set();
                const emailPattern = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
                
                // Function to extract emails from text
                const extractEmailsFromText = (text) => {
                    if (!text) return;
                    const matches = text.match(emailPattern);
                    if (matches) {
                        matches.forEach(email => {
                            const cleanEmail = email.toLowerCase().trim();
                            // Filter out common false positives
                            if (!cleanEmail.includes('example.com') && 
                                !cleanEmail.includes('noreply@') && 
                                cleanEmail.length > 5) {
                                addresses.add(cleanEmail);
                            }
                        });
                    }
                };
                
                // 1. Search for specific email header fields (From, To, CC, BCC)
                const headerSelectors = [
                    // Microsoft Outlook specific selectors
                    '[data-testid*="from"]', '[data-testid*="to"]', '[data-testid*="cc"]', '[data-testid*="bcc"]',
                    '[aria-label*="From:"]', '[aria-label*="To:"]', '[aria-label*="Cc:"]', '[aria-label*="Bcc:"]',
                    '[aria-label*="FROM:"]', '[aria-label*="TO:"]', '[aria-label*="CC:"]', '[aria-label*="BCC:"]',
                    
                    // Generic email field selectors
                    '.from-field', '.to-field', '.cc-field', '.bcc-field',
                    '.email-from', '.email-to', '.email-cc', '.email-bcc',
                    
                    // Outlook Web specific selectors
                    '[data-automation-id*="from"]', '[data-automation-id*="to"]', 
                    '[data-automation-id*="cc"]', '[data-automation-id*="bcc"]',
                    
                    // Look for elements with email-like titles
                    '[title*="@"]',
                    
                    // Conversation participant selectors
                    '.participants', '.conversation-participants', '.email-participants',
                    '[role="listitem"]', // Often used for participant lists
                    
                    // Header container selectors
                    '.email-header', '.message-header', '.conversation-header'
                ];
                
                headerSelectors.forEach(selector => {
                    try {
                        const elements = document.querySelectorAll(selector);
                        elements.forEach(element => {
                            // Extract from text content
                            extractEmailsFromText(element.textContent);
                            
                            // Extract from aria-label
                            extractEmailsFromText(element.getAttribute('aria-label'));
                            
                            // Extract from title
                            extractEmailsFromText(element.getAttribute('title'));
                            
                            // Extract from data attributes that might contain emails
                            const dataAttrs = element.getAttributeNames();
                            dataAttrs.forEach(attr => {
                                if (attr.startsWith('data-')) {
                                    extractEmailsFromText(element.getAttribute(attr));
                                }
                            });
                        });
                    } catch (e) {
                        // Continue with next selector
                    }
                });
                
                // 2. Search for conversation thread emails (multiple emails in thread)
                const conversationSelectors = [
                    '.conversation-item', '.message-item', '.email-item',
                    '[role="article"]', '[role="group"]', // Semantic containers
                    '.ms-FocusZone', // Microsoft Fabric UI containers
                    '[data-testid*="message"]', '[data-testid*="conversation"]'
                ];
                
                conversationSelectors.forEach(selector => {
                    try {
                        const conversationItems = document.querySelectorAll(selector);
                        conversationItems.forEach(item => {
                            // Look for email headers within each conversation item
                            const itemHeaderSelectors = [
                                '[aria-label*="From:"]', '[aria-label*="To:"]', 
                                '[aria-label*="Cc:"]', '[aria-label*="Bcc:"]'
                            ];
                            
                            itemHeaderSelectors.forEach(headerSel => {
                                const headerElements = item.querySelectorAll(headerSel);
                                headerElements.forEach(headerEl => {
                                    extractEmailsFromText(headerEl.textContent);
                                    extractEmailsFromText(headerEl.getAttribute('aria-label'));
                                });
                            });
                            
                            // Also extract from any text that looks like email headers
                            const itemText = item.textContent || '';
                            const headerPatterns = [
                                /From:\s*([^<]*<[^>]+>|[^\s]+@[^\s]+)/gi,
                                /To:\s*([^<]*<[^>]+>|[^\s]+@[^\s]+)/gi,
                                /Cc:\s*([^<]*<[^>]+>|[^\s]+@[^\s]+)/gi,
                                /Bcc:\s*([^<]*<[^>]+>|[^\s]+@[^\s]+)/gi
                            ];
                            
                            headerPatterns.forEach(pattern => {
                                const matches = itemText.match(pattern);
                                if (matches) {
                                    matches.forEach(match => extractEmailsFromText(match));
                                }
                            });
                        });
                    } catch (e) {
                        // Continue
                    }
                });
                
                // 3. Search in all aria-labels (comprehensive scan)
                const elementsWithAriaLabel = document.querySelectorAll('[aria-label]');
                elementsWithAriaLabel.forEach(element => {
                    const ariaLabel = element.getAttribute('aria-label');
                    if (ariaLabel && (
                        ariaLabel.toLowerCase().includes('from:') ||
                        ariaLabel.toLowerCase().includes('to:') ||
                        ariaLabel.toLowerCase().includes('cc:') ||
                        ariaLabel.toLowerCase().includes('bcc:') ||
                        ariaLabel.includes('@')
                    )) {
                        extractEmailsFromText(ariaLabel);
                    }
                });
                
                // 4. Search in all title attributes
                const elementsWithTitle = document.querySelectorAll('[title]');
                elementsWithTitle.forEach(element => {
                    const title = element.getAttribute('title');
                    if (title && title.includes('@')) {
                        extractEmailsFromText(title);
                    }
                });
                
                // 5. Search for email addresses in specific Outlook UI elements
                const outlookSpecificSelectors = [
                    // Outlook ribbon and header areas
                    '.ms-CommandBar', '.ms-Pivot', '.ms-DetailsList',
                    
                    // Email display areas
                    '[data-automation-id="MessageHeader"]',
                    '[data-automation-id="FromLine"]',
                    '[data-automation-id="ToLine"]',
                    '[data-automation-id="CcLine"]',
                    '[data-automation-id="BccLine"]',
                    
                    // Conversation view elements
                    '[data-automation-id="ConversationContainer"]',
                    '[data-automation-id="MessageContainer"]'
                ];
                
                outlookSpecificSelectors.forEach(selector => {
                    try {
                        const elements = document.querySelectorAll(selector);
                        elements.forEach(element => {
                            extractEmailsFromText(element.textContent);
                        });
                    } catch (e) {
                        // Continue
                    }
                });
                
                // 6. Final comprehensive text scan of email body (as backup)
                const emailBodySelectors = [
                    '[role="main"]', '.email-content', '[data-testid="email-body"]',
                    '.message-body', '.conversation-body'
                ];
                
                emailBodySelectors.forEach(selector => {
                    try {
                        const bodyElement = document.querySelector(selector);
                        if (bodyElement) {
                            extractEmailsFromText(bodyElement.textContent);
                        }
                    } catch (e) {
                        // Continue
                    }
                });
                
                console.log(`Found ${addresses.size} unique email addresses in conversation`);
                return Array.from(addresses);
            });
            
            foundEmails = emailAddresses || [];
            
            if (foundEmails.length > 0) {
                console.log(`✅ Found ${foundEmails.length} email address(es) inside ${folderType} email ${index}: ${foundEmails.slice(0, 3).join(', ')}${foundEmails.length > 3 ? '...' : ''}`);
            } else {
                console.log(`⚠️  No email addresses found inside ${folderType} email ${index}`);
            }
            
            // Enhanced navigation back to email list
            try {
                // Method 1: Press Escape multiple times
                for (let i = 0; i < 3; i++) {
                    await this.page.keyboard.press('Escape');
                    await new Promise(resolve => setTimeout(resolve, 500));
                }
                
                // Method 2: Try clicking back button if available
                const backSelectors = [
                    'button[aria-label*="Back"]',
                    'button[title*="Back"]', 
                    '[data-testid*="back"]',
                    'button[data-automation-id="BackButton"]',
                    '.ms-CommandBar-item button[aria-label*="Back"]'
                ];
                
                for (const selector of backSelectors) {
                    try {
                        const backButton = await this.page.$(selector);
                        if (backButton) {
                            await backButton.click();
                            await new Promise(resolve => setTimeout(resolve, 1000));
                            break;
                        }
                    } catch (e) {
                        continue;
                    }
                }
                
                // Method 3: Click on the folder name to return to list view
                try {
                    const folderButton = await this.page.$(`button[aria-label*="${folderType}"], [title*="${folderType}"]`);
                    if (folderButton) {
                        await folderButton.click();
                        await new Promise(resolve => setTimeout(resolve, 1000));
                    }
                } catch (e) {
                    // Continue
                }
                
            } catch (e) {
                console.log('Note: Using fallback navigation method');
            }
            
            // Wait and confirm we're back at the email list view
            try {
                await this.page.waitForSelector('[role="listbox"]', { timeout: 8000 });
                
                // Extra wait to ensure the list is stable after navigation
                await new Promise(resolve => setTimeout(resolve, 1500));
                
            } catch (e) {
                console.log('Warning: Could not confirm we\'re back at email list view - continuing anyway');
                
                // Fallback: try to reload the current folder
                try {
                    await this.page.reload({ waitUntil: 'networkidle0' });
                    await new Promise(resolve => setTimeout(resolve, 3000));
                } catch (reloadError) {
                    console.log('Could not reload page as fallback');
                }
            }
            
        } catch (error) {
            console.error(`❌ Error extracting email data for ${folderType} email ${index}:`, error.message);
            
            // Try to get back to email list if we got stuck
            try {
                await this.page.keyboard.press('Escape');
                await new Promise(resolve => setTimeout(resolve, 1000));
            } catch (e) {
                // Ignore
            }
        }
        
        return foundEmails.length > 0 ? foundEmails : null;
    }

    // Enhanced method to expand conversation threads and show all email headers
    async expandConversationThread() {
        try {
            console.log('🔄 Expanding conversation thread to access all emails and headers...');
            
            // First, look for conversation expansion indicators and click them
            const expansionSelectors = [
                '[aria-label*="Show"]',
                '[aria-label*="messages"]', 
                '[aria-label*="items"]',
                '[aria-label*="replies"]',
                '[data-testid*="expand"]',
                '[title*="Show"]',
                '.conversation-expansion',
                '.show-more',
                '.expand-thread',
                'button[aria-expanded="false"]',
                '[role="button"][aria-label*="Show more"]',
                '[role="button"][aria-label*="Show all"]',
                '[role="button"][title*="Show"]',
                // Outlook specific expansion buttons
                '[data-automation-id*="expand"]',
                '[data-automation-id*="showMore"]'
            ];
            
            let expandedSomething = false;
            
            for (const selector of expansionSelectors) {
                try {
                    const expansionElements = await this.page.$$(selector);
                    for (const element of expansionElements) {
                        const elementText = await this.page.evaluate(el => {
                            return (el.textContent || el.getAttribute('aria-label') || el.getAttribute('title') || '').toLowerCase();
                        }, element);
                        
                        // Look for expansion-related text patterns
                        if (elementText.includes('show') || 
                            elementText.includes('more') || 
                            elementText.includes('expand') ||
                            elementText.includes('messages') ||
                            elementText.includes('items') ||
                            elementText.includes('replies') ||
                            /\d+\s*(more|additional|items|messages|replies)/.test(elementText)) {
                            
                            console.log(`🔍 Found expansion button: "${elementText.substring(0, 100)}"`);
                            await element.click();
                            await new Promise(resolve => setTimeout(resolve, 2000));
                            expandedSomething = true;
                        }
                    }
                } catch (e) {
                    // Continue with next selector
                    continue;
                }
            }
            
            // Try to expand collapsed conversation items
            try {
                const collapsedItems = await this.page.$$('[aria-expanded="false"], .collapsed, .minimized');
                for (const item of collapsedItems) {
                    try {
                        const isVisible = await this.page.evaluate(el => {
                            const rect = el.getBoundingClientRect();
                            return rect.width > 0 && rect.height > 0;
                        }, item);
                        
                        if (isVisible) {
                            await item.click();
                            await new Promise(resolve => setTimeout(resolve, 1000));
                            expandedSomething = true;
                        }
                    } catch (e) {
                        continue;
                    }
                }
            } catch (e) {
                // Ignore errors in expansion attempts
            }
            
            // Look for and click "Show headers" or similar options
            const headerExpansionSelectors = [
                '[aria-label*="Show headers"]',
                '[aria-label*="View headers"]',
                '[aria-label*="Show details"]',
                '[aria-label*="View details"]',
                '[title*="Show headers"]',
                '[title*="View headers"]',
                '[role="button"][aria-label*="header"]',
                '[data-automation-id*="header"]'
            ];
            
            for (const selector of headerExpansionSelectors) {
                try {
                    const headerElements = await this.page.$$(selector);
                    for (const element of headerElements) {
                        try {
                            await element.click();
                            await new Promise(resolve => setTimeout(resolve, 1500));
                            expandedSomething = true;
                            console.log('📋 Expanded email headers view');
                        } catch (e) {
                            continue;
                        }
                    }
                } catch (e) {
                    continue;
                }
            }
            
            // Look for conversation thread navigation and try to show all messages
            try {
                const threadNavigation = await this.page.$$('[aria-label*="conversation"], [data-testid*="thread"], .conversation-nav');
                for (const nav of threadNavigation) {
                    try {
                        await nav.click();
                        await new Promise(resolve => setTimeout(resolve, 1500));
                    } catch (e) {
                        continue;
                    }
                }
            } catch (e) {
                // Ignore navigation errors
            }
            
            // Try to show CC/BCC fields if they're hidden
            const ccBccSelectors = [
                '[aria-label*="Show Cc"]',
                '[aria-label*="Show Bcc"]', 
                '[aria-label*="Show CC"]',
                '[aria-label*="Show BCC"]',
                '[title*="Show Cc"]',
                '[title*="Show Bcc"]',
                '[data-automation-id*="ShowCc"]',
                '[data-automation-id*="ShowBcc"]'
            ];
            
            for (const selector of ccBccSelectors) {
                try {
                    const ccBccElements = await this.page.$$(selector);
                    for (const element of ccBccElements) {
                        try {
                            await element.click();
                            await new Promise(resolve => setTimeout(resolve, 1000));
                            expandedSomething = true;
                            console.log('📧 Expanded CC/BCC fields');
                        } catch (e) {
                            continue;
                        }
                    }
                } catch (e) {
                    continue;
                }
            }
            
            if (expandedSomething) {
                console.log('✅ Successfully expanded conversation thread and headers');
                // Wait longer for content to load after expansion
                await new Promise(resolve => setTimeout(resolve, 4000));
            } else {
                console.log('ℹ️  No conversation expansion needed or available');
            }
            
        } catch (error) {
            console.log(`⚠️  Error expanding conversation: ${error.message}`);
            // Don't throw - continue with extraction even if expansion fails
        }
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