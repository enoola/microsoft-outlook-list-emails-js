const { chromium } = require('playwright');
const fs = require('fs-extra');

async function test() {
    const authFile = '../microsoft-webauth-playwright-js/auth.json';
    
    console.log('=== Starting click test ===');
    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext({ storageState: authFile });
    const page = await context.newPage();
    
    try {
        // Navigate
        console.log('Navigating to Outlook...');
        await Promise.all([
            page.goto('https://outlook.cloud.microsoft/mail/', { waitUntil: 'domcontentloaded', timeout: 30000 }),
            page.waitForTimeout(15000)
        ]);
        
        const url = page.url();
        console.log('URL:', url);
        
        // Check for #newSessionLink
        const hasNewSessionLink = await page.$eval('body', () => 
            document.querySelector('#newSessionLink') !== null
        ).catch(() => false);
        console.log('hasNewSessionLink:', hasNewSessionLink);
        
        if (hasNewSessionLink) {
            // Get element info before clicking
            const elemInfo = await page.$eval('#newSessionLink', (el) => ({
                id: el.id,
                href: el.href,
                target: el.target,
                tagName: el.tagName,
                innerText: el.innerText.substring(0, 100)
            }));
            console.log('Element info:', JSON.stringify(elemInfo));
            
            // Try clicking with different methods
            console.log('=== Attempting click ===');
            
            try {
                await page.click('#newSessionLink', { timeout: 30000 });
                console.log('Click succeeded, waiting for navigation...');
                
                // Wait for URL to change or navigation
                const newUrl = page.url();
                console.log('URL after click:', newUrl);
                
                if (newUrl !== url) {
                    console.log('=== Navigation successful! ===');
                } else {
                    console.log('URL did not change, trying alternative...');
                    
                    // Try JavaScript click
                    await page.evaluate(() => {
                        const link = document.querySelector('#newSessionLink');
                        if (link) {
                            link.removeAttribute('target');
                            link.click();
                        }
                    });
                    
                    await page.waitForTimeout(5000);
                    console.log('URL after JS click:', page.url());
                }
            } catch (e) {
                console.error('Click error:', e.message);
            }
        }
        
    } catch (e) {
        console.error('Error:', e.message);
    } finally {
        await browser.close();
    }
}

test().catch(console.error);
