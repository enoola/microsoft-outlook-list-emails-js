const { chromium } = require('playwright');
const fs = require('fs-extra');

async function test() {
    const authFile = '../microsoft-webauth-playwright-js/auth.json';
    
    console.log('Launching browser...');
    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext({ storageState: authFile });
    const page = await context.newPage();
    
    try {
        console.log('Navigating to Outlook...');
        await page.goto('https://outlook.cloud.microsoft/mail/', { waitUntil: 'domcontentloaded', timeout: 30000 });
        
        // Wait for any redirects
        await page.waitForTimeout(15000);
        const url = page.url();
        console.log('Current URL:', url);
        
        // Check if we're on the intermediate Microsoft login page with #newSessionLink
        if (url.includes('login.microsoftonline.com')) {
            console.log('Waiting for page to fully load...');
            await page.waitForTimeout(15000);
            
            // Try to find and click #newSessionLink using JavaScript
            const clicked = await page.evaluate(() => {
                const link = document.querySelector('#newSessionLink');
                if (link) {
                    console.log('Found #newSessionLink, clicking...');
                    link.click();
                    return true;
                }
                console.log('#newSessionLink not found');
                return false;
            });
            
            if (clicked) {
                console.log('Waiting for navigation after click...');
                try {
                    await page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 30000 });
                    console.log('Navigation complete!');
                } catch (e) {
                    console.log('Navigation error:', e.message);
                }
            }
        }
        
        // Wait for mail page
        await page.waitForTimeout(15000);
        console.log('Final URL:', page.url());
        
    } catch (e) {
        console.error('Error:', e.message);
    } finally {
        await browser.close();
    }
}

test().catch(console.error);
