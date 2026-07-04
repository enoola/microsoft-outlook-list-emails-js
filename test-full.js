const { chromium } = require('playwright');
const fs = require('fs-extra');

async function test() {
    const authFile = '../microsoft-webauth-playwright-js/auth.json';
    
    console.log('=== Starting full test ===');
    console.log('Launching browser...');
    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext({ storageState: authFile });
    const page = await context.newPage();
    
    try {
        console.log('Navigating to Outlook...');
        await Promise.all([
            page.goto('https://outlook.cloud.microsoft/mail/', { waitUntil: 'domcontentloaded', timeout: 30000 }),
            page.waitForTimeout(15000) // Wait for redirect
        ]);
        
        const url = page.url();
        console.log('Current URL:', url);
        
        if (url.includes('login.microsoftonline.com')) {
            console.log('=== Waiting 15 seconds for page to load ===');
            await page.waitForTimeout(15000);
            
            // Check for #newSessionLink
            const hasNewSessionLink = await page.$eval('body', () => 
                document.querySelector('#newSessionLink') !== null
            ).catch(() => false);
            
            if (hasNewSessionLink) {
                console.log('=== Found #newSessionLink, clicking... ===');
                await Promise.all([
                    page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 30000 }),
                    page.click('#newSessionLink')
                ]);
                console.log('=== Navigation complete! ===');
            }
        }
        
        // Final wait
        console.log('=== Waiting 15 seconds for stabilization ===');
        await page.waitForTimeout(15000);
        
        const finalUrl = page.url();
        console.log('Final URL:', finalUrl);
        console.log('Is mail page:', finalUrl.includes('/mail/'));
        
    } catch (e) {
        console.error('Error:', e.message);
    } finally {
        await browser.close();
    }
}

test().catch(console.error);
