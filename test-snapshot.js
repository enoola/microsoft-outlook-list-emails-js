const { chromium } = require('playwright');
const fs = require('fs-extra');

async function main() {
    const authFile = '../microsoft-webauth-playwright-js/auth.json';
    
    if (!(await fs.pathExists(authFile))) {
        console.error(`Auth file not found: ${authFile}`);
        return;
    }
    
    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext({ storageState: authFile });
    const page = await context.newPage();
    
    try {
        // Navigate to Outlook
        console.log('Navigating to Outlook...');
        await page.goto('https://outlook.cloud.microsoft/mail/');
        
        // Wait for navigation to complete
        await page.waitForLoadState('networkidle', { timeout: 30000 });
        
        console.log(`Current URL after networkidle: ${page.url()}`);
        
        // Wait a bit more for any client-side redirects
        await page.waitForTimeout(5000);
        console.log(`Current URL after wait: ${page.url()}`);
        
        // Dump HTML content
        const html = await page.content();
        await fs.writeFile('/tmp/outlook-mail.html', html);
        console.log('HTML dumped to /tmp/outlook-mail.html');
        
        // Take screenshot
        await page.screenshot({ path: '/tmp/outlook-mail.png' });
        console.log('Screenshot saved to /tmp/outlook-mail.png');
        
        // Try various selectors
        const selectors = [
            '[aria-label*="message list"]',
            'table[role="grid"]',
            '.owaRoot table',
            '#appContainer table',
            'div[data-automationid*="msg"]',
            'tr[role="row"]'
        ];
        
        for (const selector of selectors) {
            try {
                const count = await page.evaluate((sel) => document.querySelectorAll(sel).length, selector);
                console.log(`Selector "${selector}" found ${count} elements`);
            } catch (e) {
                console.log(`Selector "${selector}" error: ${e.message}`);
            }
        }
        
        // Wait before closing
        await page.waitForTimeout(30000);
    } finally {
        await browser.close();
    }
}

main().catch(console.error);
