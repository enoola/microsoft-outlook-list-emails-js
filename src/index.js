#!/usr/bin/env node
const { program } = require('commander');
const logger = require('./utils/logger');
const fs = require('fs-extra');
const { listEmails, listEmailsCount } = require('./list-emails');

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
    .option('--output-file <filename>', 'Write email list to JSON file instead of stdout')
    .option('--count', 'Return only the total count of emails')
    .action(async (options) => {
        try {
            if (options.count) {
                const count = await listEmailsCount(options);
                console.log(count);
                logger.success(`Total emails: ${count}`);
            } else {
                const emails = await listEmails(options);

                // Output as JSON
                const jsonOutput = JSON.stringify(emails, null, 2);

                if (options.outputFile) {
                    await fs.writeFile(options.outputFile, jsonOutput, 'utf8');
                    logger.success(`Email list written to ${options.outputFile}`);
                } else {
                    console.log(jsonOutput);
                }

                if (emails.length === 0) {
                    logger.warn('No emails found.');
                } else {
                    logger.success(`Listed ${emails.length} email(s).`);
                }
            }
        } catch (e) {
            logger.error('Failed to list emails.', e);
            process.exit(1);
        }
    });

program.parse();
