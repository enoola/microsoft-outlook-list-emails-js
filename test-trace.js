const { chromium } = require('playwright');
const fs = require('fs-extra');

async function test() {
    const authFile = '../microsoft-webauth-playwright-js/auth.json';
    
    console.log('=== Starting trace ===');
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
        console.log('URL after navigation:', url);
        
        // Check for #newSessionLink immediately
        console.log('Checking for #newSessionLink immediately...');
        const hasNewSessionLink = await page.$eval('body', () => 
            document.querySelector('#newSessionLink') !== null
        ).catch(() => false);
        console.log('hasNewSessionLink:', hasNewSessionLink);
        
        // If not found, wait and try again
        if (!hasNewSessionLink) {
            console.log('Not found yet, waiting 15 seconds...');
            await page.waitForTimeout(15000);
            
            const hasNewSessionLink2 = await page.$eval('body', () => 
                document.querySelector('#newSessionLink') !== null
            ).catch(() => false);
            console.log('hasNewSessionLink after wait:', hasNewSessionLink2);
            
            if (hasNewSessionLink2) {
                console.log('=== Clicking #newSessionLink ===');
                await Promise.all([
                    page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 30000 }),
                    page.click('#newSessionLink')
                ]);
                console.log('=== Navigation complete! ===');
            }
        }
        
        // Final wait
        await page.waitForTimeout(15000);
        console.log('Final URL:', page.url());
        
    } catch (e) {
        console.error('Error:', e.message);
    } finally {
        await browser.close();
    }
}

test().catch(console.error);
