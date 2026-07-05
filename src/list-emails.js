const { chromium } = require('playwright');
const logger = require('./utils/logger');
const fs = require('fs-extra');
const path = require('path');
const { OUTLOOK_URL, MICROSOFT_LOGIN_URL_PATTERN } = require('./config');

/**
 * Wait for page to load with timeout
 * @param {import('playwright').Page} page
 * @param {number} timeout - Timeout in milliseconds
 */
async function waitForPageLoad(page, timeout = 20000) {
    logger.debug(`Waiting up to ${timeout}ms for page to load...`);
    try {
        await page.waitForLoadState('domcontentloaded', { timeout });
        logger.debug('DOM content loaded.');
    } catch (e) {
        logger.warn(`domcontentloaded timeout: ${e.message}`);
    }
    
    try {
        await page.waitForLoadState('networkidle', { timeout });
        logger.debug('Network idle reached.');
    } catch (e) {
        logger.warn(`networkidle timeout: ${e.message}`);
    }
}

/**
 * Handle Microsoft's intermediate login confirmation page
 * @param {import('playwright').Page} page
 * @returns {Promise<boolean>} true if the page was handled, false otherwise
 */
async function handleMicrosoftLoginConfirm(page) {
    const url = page.url();

    // Check if we're on a Microsoft login page (either authorize or common/oauth2)
    const isLoginUrl = url.includes('login.microsoftonline.com') || url.includes('login.live.com');

    if (!isLoginUrl) {
        return false;
    }

    logger.warn(`Detected Microsoft login page: ${url}`);
    logger.info('Waiting up to 20 seconds for confirmation page or mail interface...');

    try {
        // Wait for either the confirmation page OR the mail interface
        const result = await Promise.race([
            // Option 1: Confirmation/selection page handling
            (async () => {
                // Give time for dynamic content to load
                await page.waitForTimeout(5000);

                // Check if we have #newSessionLink directly (the intermediate login page)
                const hasNewSessionLink = await page.$eval('body', (body) => {
                    return document.querySelector('#newSessionLink') !== null;
                }).catch(() => false);

                if (hasNewSessionLink) {
                    logger.info('Microsoft intermediate login page detected with #newSessionLink.');

                    // Try to find and click #newSessionLink
                    try {
                        await page.waitForSelector('#newSessionLink', { state: 'visible', timeout: 15000 });
                        const link = page.locator('#newSessionLink').first();

                        logger.info('Found "#newSessionLink" - clicking to continue with existing session...');

                        // Wait for navigation after click
                        await Promise.all([
                            page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 30000 }),
                            link.click()
                        ]);

                        logger.success('Redirected after confirming session.');
                        return { handled: true, type: 'confirmation' };
                    } catch (e) {
                        logger.warn(`Could not click #newSessionLink: ${e.message}`);
                        return { handled: false, type: 'no_link' };
                    }
                }

                // Check if we have the "We found an account you can use" text
                const hasAccountText = await page.$eval('body', (body) => {
                    return body.innerText.includes('account you can use') ||
                           body.innerText.includes('sign in with');
                }).catch(() => false);

                if (hasAccountText) {
                    logger.info('Microsoft account confirmation page detected.');

                    // Try to find and click #newSessionLink
                    try {
                        await page.waitForSelector('#newSessionLink', { state: 'visible', timeout: 15000 });
                        const link = page.locator('#newSessionLink').first();

                        logger.info('Found "#newSessionLink" - clicking to continue with existing session...');

                        await Promise.all([
                            page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 30000 }),
                            link.click()
                        ]);

                        logger.success('Redirected after confirming session.');
                        return { handled: true, type: 'confirmation' };
                    } catch (e) {
                        logger.warn(`Could not click #newSessionLink: ${e.message}`);
                        return { handled: false, type: 'no_link' };
                    }
                }

                // Check if we're on the account SELECTION page (prompt=select_account)
                // This shows a list of accounts to choose from
                const isSelectionPage = await page.$eval('body', (body) => {
                    return body.innerText.includes('Sign in') &&
                           document.querySelector('[data-bind*="account"]') !== null;
                }).catch(() => false);

                if (isSelectionPage) {
                    logger.info('Microsoft account selection page detected.');
                    
                    // Try to automatically select the first available account button
                    try {
                        await page.waitForSelector('[data-bind*="account"], .entry-button, .btn-primary', { 
                            state: 'visible', 
                            timeout: 15000 
                        });

                        const clicked = await page.evaluate(() => {
                            // Try to find and click the first account button
                            const buttons = Array.from(document.querySelectorAll('button, a'));
                            
                            for (const btn of buttons) {
                                const text = btn.innerText.toLowerCase();
                                // Look for buttons with email addresses or "Sign in" text
                                if ((text.includes('@') && text.length < 50) || 
                                    text.includes('sign in') ||
                                    btn.getAttribute('data-bind')?.includes('email')) {
                                    try {
                                        btn.click();
                                        return true;
                                    } catch (e) {
                                        // Try alternative click method
                                        const event = new MouseEvent('click', { bubbles: true });
                                        btn.dispatchEvent(event);
                                        return true;
                                    }
                                }
                            }
                            
                            // Fallback: click any primary button
                            for (const btn of buttons) {
                                if (btn.classList.contains('btn-primary') || 
                                    btn.classList.contains('button') ||
                                    btn.getAttribute('type') === 'submit') {
                                    try {
                                        btn.click();
                                        return true;
                                    } catch (e) {
                                        const event = new MouseEvent('click', { bubbles: true });
                                        btn.dispatchEvent(event);
                                        return true;
                                    }
                                }
                            }
                            
                            return false;
                        });

                        if (clicked) {
                            logger.success('Selected account from selection page.');
                            // Wait for navigation to mail page
                            await page.waitForURL(url => url.includes('/mail/'), { timeout: 30000 });
                            return { handled: true, type: 'selection' };
                        }
                    } catch (e) {
                        logger.warn(`Could not select account: ${e.message}`);
                    }
                    
                    return { handled: false, type: 'no_selection' };
                }

                // Check if we're already at the mail page
                if (url.includes('/mail/')) {
                    return { handled: true, type: 'already_at_mail' };
                }

                return { handled: false, type: 'not_confirmed' };
            })(),

            // Option 2: Direct navigation to mail page (no confirmation needed)
            (async () => {
                await page.waitForURL(url => url.includes('/mail/'), { timeout: 20000 });
                logger.success('Reached Outlook mail page directly.');
                return { handled: true, type: 'direct' };
            })()
        ]);

        return result.handled;
    } catch (e) {
        logger.warn(`Microsoft login confirmation handler encountered an issue: ${e.message}`);
        
        // If we ended up at the mail page despite the error, consider it handled
        if (page.url().includes('/mail/')) {
            logger.success('Reached Outlook mail page after handling attempt.');
            return true;
        }
        
        return false;
    }
}

/**
 * Detect if we're on the Outlook mail page
 * @param {import('playwright').Page} page
 * @returns {Promise<boolean>}
 */
async function isOutlookMailPage(page) {
    const url = page.url();

    // Check for Outlook mail URL or wait for mail-specific elements
    if (url.includes('/mail/')) {
        return true;
    }

    // Fallback: look for mail UI elements
    try {
        await page.waitForSelector('[aria-label*="message list"], [role="grid"][aria-label*="mail"]', { 
            state: 'visible', 
            timeout: 5000 
        });
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * Navigate to Outlook and handle any intermediate pages
 * @param {import('playwright').Page} page
 */
async function navigateToOutlook(page) {
    logger.info(`Navigating to Outlook: ${OUTLOOK_URL}`);

    // Navigate to Outlook URL - this may redirect through Microsoft login
    logger.info('Navigating to Outlook...');
    await Promise.all([
        page.goto(OUTLOOK_URL, { waitUntil: 'domcontentloaded', timeout: 30000 }),
        waitForPageLoad(page, 20000)
    ]);

    const currentUrl = page.url();
    
    // Give extra time for any client-side redirects
    logger.info('Waiting 15 seconds for any intermediate redirects...');
    await page.waitForTimeout(15000);

    logger.debug(`After navigation, URL is: ${currentUrl}`);

    // If we're on a Microsoft login page, handle the confirmation flow
    if (currentUrl.includes('login.microsoftonline.com') || currentUrl.includes('login.live.com')) {
        logger.info('Microsoft login flow detected. Waiting 15 seconds for confirmation page...');
        
        // Wait longer for dynamic content to load
        await page.waitForTimeout(15000);

        // Check if we have the account confirmation text or #newSessionLink
        const hasAccountText = await page.$eval('body', (body) => {
            return body.innerText.includes('account you can use') ||
                   body.innerText.includes('sign in with') ||
                   document.querySelector('#newSessionLink') !== null;
        }).catch(() => false);

        if (hasAccountText) {
            logger.info('Microsoft account confirmation page detected.');
            
            // Wait 15 seconds before clicking to ensure page is stable
            await page.waitForTimeout(10000);

            // Try to find and click #newSessionLink
            try {
                await page.waitForSelector('#newSessionLink', { state: 'visible', timeout: 20000 });
                
                const linkElement = page.locator('#newSessionLink').first();
                
                // Get the account info for logging
                const accountInfo = await page.evaluate(() => {
                    const title = document.querySelector('#NewSessionTitle')?.textContent || '';
                    const fullName = document.querySelector('[data-bind*="newSessionFullName"]')?.textContent || '';
                    const email = document.querySelector('[data-bind*="newSessionDisplayName"]')?.textContent || '';
                    return { title, fullName, email };
                });
                
                logger.info(`Account confirmation page shows: ${accountInfo.fullName} (${accountInfo.email})`);
                logger.info('Found "#newSessionLink" - clicking to continue with existing session...');
                
                // Wait for navigation after click
                await Promise.all([
                    page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 30000 }),
                    linkElement.click()
                ]);

                logger.success('Redirected after confirming session.');
                
                // Wait for the mail interface to load
                logger.info('Waiting 15 seconds for Outlook mail interface to load...');
                await page.waitForTimeout(10000);
                await waitForPageLoad(page, 20000);
            } catch (e) {
                logger.warn(`Could not click #newSessionLink: ${e.message}`);
                
                // Try JavaScript-based click as fallback
                try {
                    const clicked = await page.evaluate(() => {
                        const elements = Array.from(document.querySelectorAll('a, button'));
                        const target = elements.find(el =>
                            el.id === 'newSessionLink' ||
                            (el.getAttribute('href') && el.getAttribute('href').includes('newSession'))
                        );
                        if (target) {
                            target.click();
                            return true;
                        }
                        return false;
                    });
                    
                    if (clicked) {
                        logger.success('Successfully clicked #newSessionLink via JavaScript fallback.');
                        await page.waitForTimeout(10000);
                    } else {
                        logger.warn('Could not find #newSessionLink even via JS scan.');
                    }
                } catch (jsError) {
                    logger.warn(`JavaScript click also failed: ${jsError.message}`);
                }
            }
        } else {
            logger.info('No account confirmation page detected, waiting for mail interface...');
            
            // Wait 15 seconds then check again
            await page.waitForTimeout(10000);
            const newUrl = page.url();
            if (newUrl.includes('/mail/')) {
                logger.success(`Reached mail interface: ${newUrl}`);
            }
        }

        // Final wait for the page to stabilize
        logger.info('Waiting 15 seconds for final page stabilization...');
        await page.waitForTimeout(10000);
    }
}

/**
 * Scrape emails from the current page
 * @param {import('playwright').Page} page
 * @returns {Promise<Array>} Array of email objects
 */
async function scrapeEmails(page) {
    logger.info('Scraping emails from mail list...');

    // Check if we're on Microsoft login page (due to client-side redirect)
    const currentUrl = page.url();
    
    // Wait 15 seconds before checking for login page
    logger.info('Waiting 15 seconds for any redirects...');
    await page.waitForTimeout(15000);
    
    if (currentUrl.includes('login.microsoftonline.com') || currentUrl.includes('login.live.com')) {
        logger.warn(`We are on Microsoft login page (${currentUrl}) instead of Outlook mail.`);
        
        // Wait 15 seconds for the confirmation/selection page to load
        await page.waitForTimeout(15000);

        try {
            // First, check for confirmation page with #newSessionLink
            const hasNewSessionLink = await page.$eval('body', (body) => {
                return document.querySelector('#newSessionLink') !== null;
            }).catch(() => false);

            if (hasNewSessionLink) {
                const link = page.locator('#newSessionLink').first();
                logger.info('Found "#newSessionLink" - clicking to confirm session...');

                // Wait for navigation after click
                await Promise.all([
                    page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 30000 }),
                    link.click()
                ]);

                logger.success('Redirected after confirming session.');
            } else {
                // Check for account selection page (prompt=select_account)
                const isSelectionPage = await page.$eval('body', (body) => {
                    return body.innerText.includes('Sign in') &&
                           document.querySelector('[data-bind*="account"]') !== null;
                }).catch(() => false);

                if (isSelectionPage) {
                    logger.info('Account selection page detected - attempting to select account...');

                    // Try to click the first account button with email text
                    const clicked = await page.evaluate(() => {
                        const buttons = Array.from(document.querySelectorAll('button, a'));
                        
                        for (const btn of buttons) {
                            const text = btn.innerText.toLowerCase();
                            if ((text.includes('@') && text.length < 50) || 
                                text.includes('sign in')) {
                                try {
                                    btn.click();
                                    return true;
                                } catch (e) {
                                    const event = new MouseEvent('click', { bubbles: true });
                                    btn.dispatchEvent(event);
                                    return true;
                                }
                            }
                        }
                        
                        // Fallback: click any primary button
                        for (const btn of buttons) {
                            if (btn.classList.contains('btn-primary') || 
                                btn.getAttribute('type') === 'submit') {
                                try {
                                    btn.click();
                                    return true;
                                } catch (e) {
                                    const event = new MouseEvent('click', { bubbles: true });
                                    btn.dispatchEvent(event);
                                    return true;
                                }
                            }
                        }
                        
                        return false;
                    });

                    if (clicked) {
                        logger.success('Selected account from selection page.');
                    } else {
                        logger.warn('Could not find account button to click.');
                    }
                } else {
                    logger.warn('Neither confirmation nor selection page detected, waiting for redirect...');
                }
            }

            // Wait 15 seconds for the mail interface to load
            logger.info('Waiting 15 seconds for Outlook mail interface...');
            await page.waitForTimeout(15000);

            // Verify we're at a mail URL
            const newUrl = page.url();
            if (newUrl.includes('/mail/')) {
                logger.success(`Reached mail interface: ${newUrl}`);
            } else {
                logger.warn(`Unexpected URL after confirmation: ${newUrl}`);
            }
        } catch (e) {
            logger.error(`Failed to handle Microsoft login confirmation in scrapeEmails: ${e.message}`);
            
            // Try JavaScript fallback
            try {
                const clicked = await page.evaluate(() => {
                    const elements = Array.from(document.querySelectorAll('a, button'));
                    const target = elements.find(el =>
                        el.id === 'newSessionLink' ||
                        (el.getAttribute('href') && el.getAttribute('href').includes('newSession'))
                    );
                    if (target) {
                        target.click();
                        return true;
                    }
                    return false;
                });
                
                if (clicked) {
                    logger.success('Successfully clicked #newSessionLink via JavaScript fallback.');
                    await page.waitForTimeout(15000);
                } else {
                    logger.warn('Could not find #newSessionLink even via JS scan.');
                }
            } catch (jsError) {
                logger.error(`JavaScript click also failed: ${jsError.message}`);
            }
            
            throw e;
        }
    }

    // Wait for stable state before scraping
    logger.info('Waiting 15 seconds for page to stabilize...');
    await page.waitForTimeout(15000);
    
    logger.debug(`Current URL after stabilization: ${page.url()}`);

    // Scrape email data using the actual Outlook Web App DOM structure
    const emails = await page.evaluate(() => {
        const results = [];

        // Find all message rows - these are divs with data-convid attribute
        const msgRows = document.querySelectorAll('[data-convid]');

        for (const row of msgRows) {
            try {
                // Extract sender from span with title attribute containing email
                const senderEl = row.querySelector('span[title*="@"]');
                const sender = senderEl ? senderEl.getAttribute('title') || '' : '';

                // Extract subject from TtcXM class or aria-label
                const subjectEl = row.querySelector('.TtcXM, [aria-label*="Unread"], [aria-label*="Read"]');
                let subject = '';
                if (subjectEl) {
                    subject = subjectEl.textContent.trim() || subjectEl.getAttribute('aria-label') || '';
                }

                // Extract date from qq2gS class
                const dateEl = row.querySelector('.qq2gS._rWRU');
                const date = dateEl ? dateEl.getAttribute('title') || '' : '';

                // Extract preview text
                const previewEl = row.querySelector('.FqgPc, .Zgp3k span');
                const preview = previewEl ? previewEl.textContent.trim().substring(0, 150) : '';

                // Only include if we have at least a subject or sender
                if (subject || sender) {
                    results.push({
                        receivedDate: date,
                        subject,
                        sender,
                        firstBodyLine: preview
                    });
                }
            } catch (e) {
                console.log(`Error extracting row: ${e.message}`);
            }
        }

        return results;
    });

    logger.success(`Scraped ${emails.length} emails.`);
    return emails;
}

/**
 * List Outlook emails
 * @param {Object} options - Command options
 * @param {string} options.authFile - Path to authentication JSON file
 * @param {boolean} [options.notheadless] - Run in visible browser mode
 * @param {boolean} [options.dodump] - Dump HTML content for debugging
 * @param {number} [options.maxResults] - Maximum number of results to return
 * @returns {Promise<Array>} Array of email objects
 */
async function listEmails(options = {}) {
    logger.info('Connecting to Outlook...');
    
    const headless = !options.notheadless;
    logger.debug(`Launching browser (headless: ${headless})...`);
    
    // Check auth file exists
    if (!(await fs.pathExists(options.authFile))) {
        throw new Error(`Authentication file not found: ${options.authFile}`);
    }
    
    const browser = await chromium.launch({ headless });
    const context = await browser.newContext({ storageState: options.authFile });

    try {
        const page = await context.newPage();

        // Navigate directly to Outlook - the auth state should handle any intermediate redirects
        logger.info(`Navigating to Outlook: ${OUTLOOK_URL}`);
        await page.goto(OUTLOOK_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
        
        // Wait for any client-side redirects to complete
        logger.info('Waiting 15 seconds for any intermediate redirects...');
        await page.waitForTimeout(15000);
        
        logger.success(`Reached mail interface: ${page.url()}`);

        // Dump debug content if requested (after navigation is complete)
        if (options.dodump) {
            const dumpDir = await logger.getDumpDir();
            const displayPath = logger.getDumpDisplayPath();
            logger.warn(`Dumping main page content to ${displayPath}/debug_page.html...`);
            try {
                const content = await page.content();
                await fs.writeFile(path.join(dumpDir, 'debug_page.html'), content);
            } catch (e) {
                logger.warn(`Could not dump page content: ${e.message}`);
            }
        }

        // Scrape emails
        let emails = await scrapeEmails(page);

        // Apply max results limit if specified
        if (options.maxResults && options.maxResults > 0) {
            emails = emails.slice(0, parseInt(options.maxResults));
            logger.info(`Limited to first ${options.maxResults} results.`);
        }

        return emails;

    } catch (e) {
        logger.error('Error listing emails:', e);
        throw e;
    } finally {
        await browser.close();
    }
}

module.exports = { listEmails, handleMicrosoftLoginConfirm };
