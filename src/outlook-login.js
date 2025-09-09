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

    

    async harvestBccContacts() {
        try {
            console.log('🎯 Starting BCC contact harvesting method...');
            
            const harvestedContacts = new Set();
            let currentStep = 1;
            let totalSteps = 8; // Estimated steps for logging
            
            // Step 1: Navigate back to main mail view if not already there
            console.log(`📍 Step ${currentStep}/${totalSteps}: Ensuring we're in the main mail view...`);
            await this.navigateToInbox();
            await new Promise(resolve => setTimeout(resolve, 2000));
            currentStep++;
            
            // Step 2: Look for and click "New mail" button
            console.log(`📍 Step ${currentStep}/${totalSteps}: Looking for "New mail" button...`);
            const newMailSelectors = [
                'button[aria-label*="New mail"]',
                'button[title*="New mail"]',
                '[data-testid*="new-mail"]',
                'button:contains("New mail")',
                '.ms-CommandBar-item button[aria-label*="New"]',
                '[data-automation-id*="NewMessage"]',
                'div[role="button"][aria-label*="New mail"]'
            ];
            
            let newMailClicked = false;
            for (const selector of newMailSelectors) {
                try {
                    const newMailButton = await this.page.$(selector);
                    if (newMailButton) {
                        console.log(`✅ Found "New mail" button with selector: ${selector}`);
                        await newMailButton.click();
                        console.log('🔘 Clicked "New mail" button successfully');
                        newMailClicked = true;
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }
            
            if (!newMailClicked) {
                throw new Error('Could not find or click "New mail" button');
            }
            currentStep++;
            
            // Step 3: Wait for compose window to load
            console.log(`📍 Step ${currentStep}/${totalSteps}: Waiting for compose window to load...`);
            await new Promise(resolve => setTimeout(resolve, 3000));
            
            // Look for compose window indicators
            const composeSelectors = [
                '[role="dialog"]',
                '.compose-container',
                '[data-testid*="compose"]',
                '.ms-Panel-main',
                '[aria-label*="compose"]'
            ];
            
            let composeLoaded = false;
            for (const selector of composeSelectors) {
                const element = await this.page.$(selector);
                if (element) {
                    console.log(`✅ Compose window detected with selector: ${selector}`);
                    composeLoaded = true;
                    break;
                }
            }
            
            if (!composeLoaded) {
                console.log('⚠️ Could not confirm compose window loaded, proceeding anyway...');
            }
            currentStep++;
            
            // Step 4: Look for and click BCC field
            console.log(`📍 Step ${currentStep}/${totalSteps}: Looking for BCC field...`);
            
            // First try to find BCC field directly
            const bccSelectors = [
                'input[aria-label*="Bcc"]',
                'input[placeholder*="Bcc"]',
                '[data-testid*="bcc"]',
                'input[name*="bcc"]',
                '.bcc-field input',
                '[role="textbox"][aria-label*="Bcc"]',
                'div[role="textbox"][aria-label*="Bcc"]',
                'input[aria-label*="bcc"]',
                'div[aria-label*="bcc"]',
                '[role="combobox"][aria-label*="Bcc"]',
                '[aria-describedby*="bcc"]',
                'input[data-automation-id*="bcc"]',
                'div[data-automation-id*="bcc"]',
                '.ms-TextField input[aria-label*="Bcc"]'
            ];
            
            let bccField = null;
            for (const selector of bccSelectors) {
                const element = await this.page.$(selector);
                if (element) {
                    console.log(`✅ Found BCC field with selector: ${selector}`);
                    bccField = element;
                    break;
                }
            }
            
            // If BCC field not found, look for "Show BCC" or similar buttons
            if (!bccField) {
                console.log('🔍 BCC field not visible, looking for "Show BCC" or "Cc/Bcc" buttons...');
                
                const showBccSelectors = [
                    'a[title*="Bcc"]',
                    'a:contains("Bcc")',
                    'button[aria-label*="Bcc"]',
                    'button:contains("Bcc")',
                    'button:contains("Cc")',
                    '[data-testid*="show-bcc"]',
                    '.show-cc-bcc',
                    'button[title*="Bcc"]',
                    'button[aria-label*="Show additional fields"]',
                    'a[aria-label*="Bcc"]',
                    '[role="link"][title*="Bcc"]'
                ];
                
                // Try manual search for Bcc link in all text elements
                console.log('🔍 Searching for "Bcc" text in all page elements...');
                const bccLinks = await this.page.evaluate(() => {
                    const elements = Array.from(document.querySelectorAll('*'));
                    return elements.filter(el => {
                        const text = el.textContent || '';
                        const title = el.getAttribute('title') || '';
                        const ariaLabel = el.getAttribute('aria-label') || '';
                        return text.includes('Bcc') || title.includes('Bcc') || ariaLabel.includes('Bcc');
                    }).map(el => ({
                        tagName: el.tagName,
                        text: el.textContent?.trim(),
                        title: el.getAttribute('title'),
                        ariaLabel: el.getAttribute('aria-label'),
                        className: el.className
                    }));
                });
                console.log(`🔍 Found ${bccLinks.length} elements containing "Bcc":`, bccLinks);
                
                for (const selector of showBccSelectors) {
                    try {
                        const showButton = await this.page.$(selector);
                        if (showButton) {
                            console.log(`🔘 Found show BCC button: ${selector}`);
                            await showButton.click();
                            await new Promise(resolve => setTimeout(resolve, 2000));
                            
                            // Try to find BCC field again after clicking
                            for (const bccSelector of bccSelectors) {
                                const element = await this.page.$(bccSelector);
                                if (element) {
                                    console.log(`✅ BCC field now visible: ${bccSelector}`);
                                    bccField = element;
                                    break;
                                }
                            }
                            if (bccField) break;
                        }
                    } catch (e) {
                        continue;
                    }
                }
                
                // If still not found, try clicking on any element that contains "Bcc" text
                if (!bccField) {
                    console.log('🔍 Trying to click on any element containing "Bcc" text...');
                    const bccClickable = await this.page.evaluate(() => {
                        const elements = Array.from(document.querySelectorAll('a, button, span, div'));
                        const bccElement = elements.find(el => {
                            const text = el.textContent || '';
                            const title = el.getAttribute('title') || '';
                            return text.includes('Bcc') || title.includes('Bcc');
                        });
                        return bccElement;
                    });
                    
                    if (bccClickable) {
                        console.log('🔘 Found clickable Bcc element, attempting click...');
                        await this.page.evaluate((el) => el.click(), bccClickable);
                        await new Promise(resolve => setTimeout(resolve, 2000));
                        
                        // Try to find BCC field again
                        for (const bccSelector of bccSelectors) {
                            const element = await this.page.$(bccSelector);
                            if (element) {
                                console.log(`✅ BCC field now visible after text click: ${bccSelector}`);
                                bccField = element;
                                break;
                            }
                        }
                    }
                }
            }
            
            if (!bccField) {
                throw new Error('Could not find BCC field even after trying to show it');
            }
            currentStep++;
            
            // Step 5: Click on BCC field to trigger suggestions
            console.log(`📍 Step ${currentStep}/${totalSteps}: Clicking on BCC field to trigger contact suggestions...`);
            await bccField.click();
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            // Click on empty part of BCC field to trigger suggestions
            console.log('🔍 Clicking on empty BCC field area to trigger contact suggestions...');
            await bccField.click();
            await new Promise(resolve => setTimeout(resolve, 1000));
            
            // Try clicking at different positions within the BCC field
            const bccBox = await bccField.boundingBox();
            if (bccBox) {
                // Click at the beginning of the field
                await this.page.mouse.click(bccBox.x + 10, bccBox.y + bccBox.height / 2);
                await new Promise(resolve => setTimeout(resolve, 1000));
                
                // Click at the end of the field
                await this.page.mouse.click(bccBox.x + bccBox.width - 10, bccBox.y + bccBox.height / 2);
                await new Promise(resolve => setTimeout(resolve, 1000));
                
                // Click in the middle
                await this.page.mouse.click(bccBox.x + bccBox.width / 2, bccBox.y + bccBox.height / 2);
                await new Promise(resolve => setTimeout(resolve, 1000));
            }
            
            currentStep++;
            
            // Step 6: Start harvesting suggested contacts
            console.log(`📍 Step ${currentStep}/${totalSteps}: Starting contact suggestion harvesting...`);
            
            let harvestRound = 1;
            let maxRounds = 50; // Prevent infinite loops
            let consecutiveNoContacts = 0;
            let maxConsecutiveNoContacts = 3;
            
            while (harvestRound <= maxRounds && consecutiveNoContacts < maxConsecutiveNoContacts) {
                console.log(`🔄 Harvest round ${harvestRound}: Looking for suggested contacts...`);
                
                // Look for contact suggestion dropdowns/containers
                const suggestionSelectors = [
                    '[role="listbox"]',
                    '[role="menu"]',
                    '.suggestions-container',
                    '.contact-suggestions',
                    '[data-testid*="suggestions"]',
                    '[data-testid*="contacts"]',
                    '.ms-Suggestions-container',
                    '.peoplepicker-results',
                    '[aria-label*="suggestions"]',
                    '[aria-label*="contacts"]'
                ];
                
                let suggestionContainer = null;
                for (const selector of suggestionSelectors) {
                    const element = await this.page.$(selector);
                    if (element) {
                        // Check if this container has visible contact items
                        const contactItems = await element.$$('[role="option"], .contact-item, .suggestion-item, [data-testid*="contact"]');
                        if (contactItems.length > 0) {
                            console.log(`📋 Found suggestion container with ${contactItems.length} contacts: ${selector}`);
                            suggestionContainer = element;
                            break;
                        }
                    }
                }
                
                if (!suggestionContainer) {
                    console.log(`⚠️ No suggestion container found in round ${harvestRound}`);
                    consecutiveNoContacts++;
                    harvestRound++;
                    continue;
                }
                
                // Reset consecutive no contacts since we found some
                consecutiveNoContacts = 0;
                
                // Extract contact information from suggestions
                const contactItems = await suggestionContainer.$$('[role="option"], .contact-item, .suggestion-item, [data-testid*="contact"], [data-testid*="person"]');
                console.log(`👥 Found ${contactItems.length} contact suggestions in round ${harvestRound}`);
                
                let newContactsThisRound = 0;
                
                for (let i = 0; i < contactItems.length; i++) {
                    try {
                        const contactInfo = await this.page.evaluate(element => {
                            // Extract email and name from the contact element
                            const emailPattern = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
                            const text = element.textContent || '';
                            const ariaLabel = element.getAttribute('aria-label') || '';
                            const title = element.getAttribute('title') || '';
                            
                            const allText = `${text} ${ariaLabel} ${title}`;
                            const emailMatches = allText.match(emailPattern);
                            
                            return {
                                text: text.trim(),
                                ariaLabel: ariaLabel.trim(),
                                title: title.trim(),
                                emails: emailMatches || [],
                                fullText: allText
                            };
                        }, contactItems[i]);
                        
                        // Process found emails
                        if (contactInfo.emails.length > 0) {
                            for (const email of contactInfo.emails) {
                                if (!harvestedContacts.has(email.toLowerCase())) {
                                    harvestedContacts.add(email.toLowerCase());
                                    newContactsThisRound++;
                                    console.log(`✅ New contact harvested: ${email}`);
                                }
                            }
                        }
                        
                        // Click on the contact to potentially reveal more suggestions
                        console.log(`🔘 Clicking on contact ${i + 1}/${contactItems.length}...`);
                        await contactItems[i].click();
                        await new Promise(resolve => setTimeout(resolve, 800));
                        
                    } catch (e) {
                        console.log(`⚠️ Error processing contact ${i + 1}: ${e.message}`);
                        continue;
                    }
                }
                
                console.log(`📊 Round ${harvestRound} summary: ${newContactsThisRound} new contacts found. Total: ${harvestedContacts.size}`);
                
                // Clear the BCC field and try again to get fresh suggestions
                try {
                    // Select all and delete
                    await this.page.keyboard.down('Control');
                    await this.page.keyboard.press('KeyA');
                    await this.page.keyboard.up('Control');
                    await this.page.keyboard.press('Delete');
                    await new Promise(resolve => setTimeout(resolve, 500));
                    
                    // Click field again to ensure focus
                    await bccField.click();
                    await new Promise(resolve => setTimeout(resolve, 500));
                } catch (e) {
                    console.log(`⚠️ Error clearing BCC field: ${e.message}`);
                }
                
                harvestRound++;
                
                // Add a longer delay between rounds to allow suggestions to refresh
                await new Promise(resolve => setTimeout(resolve, 2000));
            }
            currentStep++;
            
            // Step 7: Finalize harvesting
            console.log(`📍 Step ${currentStep}/${totalSteps}: Finalizing contact harvest...`);
            currentStep++;
            
            // Step 8: Close compose window and finalize
            console.log(`📍 Step ${currentStep}/${totalSteps}: Closing compose window and finalizing harvest...`);
            
            // Close the compose window
            const closeSelectors = [
                'button[aria-label*="Close"]',
                'button[title*="Close"]',
                '[data-testid*="close"]',
                'button[aria-label*="Discard"]',
                '.ms-Panel-closeButton',
                'button.ms-Button--icon[aria-label*="Close"]'
            ];
            
            for (const selector of closeSelectors) {
                try {
                    const closeButton = await this.page.$(selector);
                    if (closeButton) {
                        console.log(`🔘 Closing compose window with: ${selector}`);
                        await closeButton.click();
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }
            
            // If close button didn't work, try Escape key
            await this.page.keyboard.press('Escape');
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            // Convert Set to Array and log final results
            const finalContacts = Array.from(harvestedContacts);
            
            console.log('🎉 BCC Contact Harvesting Complete!');
            console.log(`📊 Total contacts harvested: ${finalContacts.length}`);
            console.log(`🔄 Harvest rounds completed: ${harvestRound - 1}`);
            
            if (finalContacts.length > 0) {
                console.log('📧 Sample harvested contacts:');
                finalContacts.slice(0, 10).forEach((contact, index) => {
                    console.log(`   ${index + 1}. ${contact}`);
                });
                if (finalContacts.length > 10) {
                    console.log(`   ... and ${finalContacts.length - 10} more contacts`);
                }
            } else {
                console.log('⚠️ No contacts were harvested. This could be due to:');
                console.log('   - No contacts in the user\'s directory');
                console.log('   - Different Outlook interface layout');
                console.log('   - Permissions or policy restrictions');
            }
            
            // Take a final screenshot
            await this.takeScreenshot(`screenshots/bcc-harvest-complete-${Date.now()}.png`);
            
            return finalContacts;
            
        } catch (error) {
            console.error('❌ Error during BCC contact harvesting:', error.message);
            
            // Take error screenshot
            await this.takeScreenshot(`screenshots/bcc-harvest-error-${Date.now()}.png`);
            
            // Try to close any open dialogs
            try {
                await this.page.keyboard.press('Escape');
                await this.page.keyboard.press('Escape');
            } catch (e) {
                // Ignore cleanup errors
            }
            
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
            console.log(`📧 Clicking on ${folderType} email ${index} to extract internal email addresses...`);
            
            // Click on the email to open it
            await emailElement.click();
            
            // Wait for email to load
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            // Wait for email content to load
            await this.page.waitForSelector('[role="main"], .email-content, [data-testid="email-body"]', { timeout: 8000 });
            
            // Expand conversation thread to get all emails in the conversation
            await this.expandConversationThread();
            
            // Extract email addresses from the opened email
            const emailAddresses = await this.page.evaluate(() => {
                const addresses = new Set();
                
                // Email pattern to find all email addresses
                const emailPattern = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
                
                // Get all text content from the email
                const fullPageText = document.body.textContent || '';
                const emailMatches = fullPageText.match(emailPattern);
                if (emailMatches) {
                    emailMatches.forEach(email => addresses.add(email.toLowerCase().trim()));
                }
                
                // Look for specific email header elements
                const emailSelectors = [
                    '[data-testid="email-from"]',
                    '[data-testid="email-to"]', 
                    '[data-testid="email-cc"]',
                    '[data-testid="email-bcc"]',
                    '[aria-label*="From:"]',
                    '[aria-label*="To:"]',
                    '[aria-label*="Cc:"]',
                    '[aria-label*="Bcc:"]',
                    '.from-field',
                    '.to-field', 
                    '.cc-field',
                    '.bcc-field',
                    '[title*="@"]',
                    '[data-automation-id*="email"]'
                ];
                
                emailSelectors.forEach(selector => {
                    try {
                        const elements = document.querySelectorAll(selector);
                        elements.forEach(element => {
                            const text = element.textContent || element.getAttribute('aria-label') || element.getAttribute('title') || '';
                            const matches = text.match(emailPattern);
                            if (matches) {
                                matches.forEach(email => addresses.add(email.toLowerCase().trim()));
                            }
                        });
                    } catch (e) {
                        // Continue with next selector
                    }
                });
                
                // Also look for email addresses in all aria-label attributes
                const elementsWithAriaLabel = document.querySelectorAll('[aria-label]');
                elementsWithAriaLabel.forEach(element => {
                    const ariaLabel = element.getAttribute('aria-label');
                    if (ariaLabel) {
                        const matches = ariaLabel.match(emailPattern);
                        if (matches) {
                            matches.forEach(email => addresses.add(email.toLowerCase().trim()));
                        }
                    }
                });
                
                // Look in all title attributes too
                const elementsWithTitle = document.querySelectorAll('[title]');
                elementsWithTitle.forEach(element => {
                    const title = element.getAttribute('title');
                    if (title) {
                        const matches = title.match(emailPattern);
                        if (matches) {
                            matches.forEach(email => addresses.add(email.toLowerCase().trim()));
                        }
                    }
                });
                
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

    // New method to expand conversation threads to get all emails in a conversation
    async expandConversationThread() {
        try {
            console.log('🔄 Expanding conversation thread to access all emails...');
            
            // Look for conversation expansion indicators and click them
            const expansionSelectors = [
                '[aria-label*="Show"]',
                '[aria-label*="messages"]', 
                '[aria-label*="items"]',
                '[data-testid*="expand"]',
                '[title*="Show"]',
                '.conversation-expansion',
                '.show-more',
                '.expand-thread',
                'button[aria-expanded="false"]',
                '[role="button"][aria-label*="Show more"]'
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
                            /\d+\s*(more|additional|items|messages)/.test(elementText)) {
                            
                            console.log(`🔍 Found expansion button: "${elementText}"`);
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
            
            // Also try to expand any collapsed conversation items by looking for specific patterns
            try {
                const collapsedItems = await this.page.$$('[aria-expanded="false"], .collapsed, .minimized');
                for (const item of collapsedItems) {
                    try {
                        await item.click();
                        await new Promise(resolve => setTimeout(resolve, 1000));
                        expandedSomething = true;
                    } catch (e) {
                        continue;
                    }
                }
            } catch (e) {
                // Ignore errors in expansion attempts
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
            
            if (expandedSomething) {
                console.log('✅ Successfully expanded conversation thread');
                // Wait for content to load after expansion
                await new Promise(resolve => setTimeout(resolve, 3000));
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