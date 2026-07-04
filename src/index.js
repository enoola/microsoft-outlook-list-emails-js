#!/usr/bin/env node
const { program } = require('commander');
const logger = require('./utils/logger');
const { listEmails } = require('./list-emails');

program
    .name('outlook-list')
    .description('List Microsoft Outlook emails via Playwright')
    .version('1.0.0');

program
    .command('list')
    .description('List available Outlook emails')
    .requiredOption('--auth-file <path>', 'Path to authentication JSON file (auth.json)')
    .option('--notheadless', 'Run in visible browser mode for debugging')
    .option('--dodump', 'Dump HTML content to files for debugging')
    .option('--max-results <n>', 'Limit number of results returned')
    .action(async (options) => {
        try {
            const emails = await listEmails(options);
            
            // Output as JSON
            console.log(JSON.stringify(emails, null, 2));
            
            if (emails.length === 0) {
                logger.warn('No emails found.');
            } else {
                logger.success(`Listed ${emails.length} email(s).`);
            }
        } catch (e) {
            logger.error('Failed to list emails.', e);
            process.exit(1);
        }
    });

program.parse();
