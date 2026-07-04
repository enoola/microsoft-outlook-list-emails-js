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
            await page.waitForTimeout(15000);

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
                await page.waitForTimeout(15000);
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
                        await page.waitForTimeout(15000);
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
            await page.waitForTimeout(15000);
            const newUrl = page.url();
            if (newUrl.includes('/mail/')) {
                logger.success(`Reached mail interface: ${newUrl}`);
            }
        }

        // Final wait for the page to stabilize
        logger.info('Waiting 15 seconds for final page stabilization...');
        await page.waitForTimeout(15000);
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

    // Try multiple selectors in sequence with increasing timeouts
    const EMAIL_SELECTORS = [
        '[aria-label*="message list"]',
        '[role="grid"][aria-label*="mail"]',
        'table[role="grid"]',
        '.owaRoot table',
        '#appContainer table',
        'div[data-automationid*="msg"]'
    ];

    let foundSelector = null;
    for (const selector of EMAIL_SELECTORS) {
        try {
            logger.info(`Trying selector: ${selector}`);
            await page.waitForSelector(selector, { state: 'visible', timeout: 20000 });
            foundSelector = selector;
            logger.success(`Email list detected using selector: ${selector}`);
            break;
        } catch (e) {
            // Try next selector
        }
    }

    if (!foundSelector) {
        // Last resort: look for any table with email-like content
        logger.info('Trying to find any table with email content...');
        
        try {
            await page.waitForFunction(() => {
                const tables = document.querySelectorAll('table');
                for (const table of tables) {
                    const hasEmail = Array.from(table.querySelectorAll('td, th')).some(td => 
                        td.innerText.includes('@') && td.innerText.length > 5
                    );
                    if (hasEmail) return true;
                }
                return false;
            }, { timeout: 30000 });
            
            logger.success('Found table with email content.');
        } catch (e) {
            logger.warn(`Could not find any table with email content: ${e.message}`);
        }
    }

    // Scrape email data
    const emails = await page.evaluate(() => {
        const results = [];

        // Try multiple selectors for Outlook mail list items
        const selectors = [
            '[aria-label*="message list"] tr',
            '[role="grid"][aria-label*="mail"] tr',
            'table[role="grid"] tr',
            '.owaRoot table tr',
            '#appContainer table tr',
            'div[data-automationid*="msg"]'
        ];

        let rows = [];
        for (const selector of selectors) {
            try {
                rows = Array.from(document.querySelectorAll(selector));
                if (rows.length > 0) {
                    break;
                }
            } catch (e) {}
        }

        console.log(`Found ${rows.length} potential email rows.`);

        // Process each row
        for (const row of rows) {
            try {
                // Try to extract date
                let receivedDate = '';
                const dateCell = row.querySelector('[aria-label*="date"], [data-automationid*="date"]') ||
                                row.querySelector('td:nth-child(1)') ||
                                row.querySelector('.receivedTime');
                if (dateCell) {
                    receivedDate = dateCell.textContent.trim() || '';
                }

                // Try to extract subject
                let subject = '';
                const subjectCell = row.querySelector('[aria-label*="subject"], [data-automationid*="subject"]') ||
                                  row.querySelector('td:nth-child(2), td:nth-child(3)') ||
                                  row.querySelector('.subject');
                if (subjectCell) {
                    subject = subjectCell.textContent.trim() || '';
                }

                // Try to extract sender
                let sender = '';
                const senderCell = row.querySelector('[aria-label*="from"], [data-automationid*="from"]') ||
                                 row.querySelector('td:nth-child(3), td:nth-child(4)') ||
                                 row.querySelector('.senderName');
                if (senderCell) {
                    sender = senderCell.textContent.trim() || '';
                }

                // Try to extract first body line (if available in preview)
                let firstBodyLine = '';
                const bodyCell = row.querySelector('[aria-label*="body"], [data-automationid*="body"]') ||
                               row.querySelector('.previewText');
                if (bodyCell) {
                    firstBodyLine = bodyCell.textContent.trim().split('\n')[0] || '';
                }

                // Only add if we have at least subject
                if (subject) {
                    results.push({
                        receivedDate,
                        subject,
                        sender,
                        firstBodyLine
                    });
                }
            } catch (e) {
                console.log(`Error parsing row: ${e.message}`);
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

        // Navigate to Outlook
        await navigateToOutlook(page);
        
        // Wait for page to be stable before any operations
        logger.info('Waiting for page to stabilize...');
        await page.waitForTimeout(3000);

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
