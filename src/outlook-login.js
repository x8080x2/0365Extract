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

    // Simplified navigation - removed authentication logic
    async navigateToEmailInterface() {
        try {
            console.log('Navigating to email interface for scanning...');
            
            // Basic navigation without authentication
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            const currentUrl = this.page.url();
            console.log(`Current URL: ${currentUrl}`);
            
            return true;
        } catch (error) {
            console.error('Error navigating to email interface:', error.message);
            return false;
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
                        extractedEmails.push(emailData);
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
            // Single approach: extract all visible text and parse it
            const emailData = {
                id: `${folderType}_${index}_${Date.now()}`,
                folder: folderType,
                index: index
            };

            // Get all text content from the email element
            const fullText = await this.page.evaluate(el => {
                // Get all text content and aria-labels
                const textContent = el.textContent?.trim() || '';
                const ariaLabel = el.getAttribute('aria-label') || '';
                const title = el.getAttribute('title') || '';
                
                // Also check for email addresses in any attribute
                const allAttributes = Array.from(el.attributes).map(attr => attr.value).join(' ');
                
                return {
                    text: textContent,
                    aria: ariaLabel,
                    title: title,
                    attributes: allAttributes
                };
            }, emailElement);

            // Extract email address (sender/recipient)
            const emailPattern = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
            const allContent = `${fullText.text} ${fullText.aria} ${fullText.title} ${fullText.attributes}`;
            const emailMatches = allContent.match(emailPattern);
            emailData.sender = emailMatches ? emailMatches[0] : 'Unknown Sender';

            // Extract subject - usually the longest meaningful text
            const textLines = fullText.text.split('\n').map(line => line.trim()).filter(line => line.length > 0);
            const meaningfulLines = textLines.filter(line => 
                line.length > 5 && 
                !line.match(/^\d+$/) && 
                !line.match(/^(AM|PM|\d{1,2}:\d{2})/) &&
                !emailPattern.test(line)
            );
            emailData.subject = meaningfulLines.length > 0 ? meaningfulLines[0] : 'No Subject';

            // Extract date - look for time patterns
            const datePattern = /(\d{1,2}:\d{2}\s*(AM|PM)|\d{1,2}\/\d{1,2}\/\d{2,4}|yesterday|today|\d+\s*(minute|hour|day)s?\s*ago)/i;
            const dateMatch = allContent.match(datePattern);
            emailData.date = dateMatch ? dateMatch[0] : 'Unknown Date';

            // Extract preview - take remaining meaningful text
            const previewText = meaningfulLines.slice(1).join(' ');
            emailData.preview = previewText.substring(0, 200);

            console.log(`Extracted email ${index}: ${emailData.sender} - ${emailData.subject}`);
            return emailData;

        } catch (error) {
            console.error(`Error extracting email data for index ${index}:`, error.message);
            return null;
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