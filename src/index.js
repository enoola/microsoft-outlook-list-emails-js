#!/usr/bin/env node
const { program } = require('commander');
const logger = require('./utils/logger');
const fs = require('fs-extra');
const chalk = require('chalk');
const { listEmails, listEmailsCount } = require('./list-emails');

// Track start time for duration calculation
const startTime = new Date();

/**
 * Format duration from start time to now
 * @returns {string} Formatted duration string
 */
function formatDuration() {
    const endTime = new Date();
    const diffMs = endTime - startTime;
    
    const hours = Math.floor(diffMs / 3600000);
    const minutes = Math.floor((diffMs % 3600000) / 60000);
    const seconds = Math.floor((diffMs % 60000) / 1000);
    
    let duration = '';
    if (hours > 0) duration += `${hours} hour${hours > 1 ? 's' : ''} `;
    if (minutes > 0) duration += `${minutes} minute${minutes > 1 ? 's' : ''} `;
    duration += `${seconds} second${seconds > 1 ? 's' : ''}`;
    
    return duration.trim();
}

/**
 * Log command start
 */
function logCommandStart() {
    const now = new Date().toISOString();
    logger.info(`Command started at ${now}`);
}

/**
 * Log command end with duration
 */
function logCommandEnd(success) {
    const now = new Date().toISOString();
    if (success) {
        console.log(chalk.green(`[SUCCESS] Command completed at ${now} (Duration: ${formatDuration()})`));
    } else {
        console.log(chalk.red(`[ERROR] Command failed at ${now} (Duration: ${formatDuration()})`));
    }
}

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
        logCommandStart();
        
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
            
            logCommandEnd(true);
        } catch (e) {
            logger.error('Failed to list emails.', e);
            logCommandEnd(false);
            process.exit(1);
        }
    });

program.parse();
