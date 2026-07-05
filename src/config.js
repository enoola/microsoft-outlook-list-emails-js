const path = require('path');

/**
 * Returns the directory where auth files are stored.
 */
function getAuthDir() {
    return path.resolve(__dirname, '..');
}

const AUTH_DIR = getAuthDir();

// Auth file location - looks for auth.json in parent webauth project first
const AUTH_FILE = path.join(AUTH_DIR, 'auth.json');

// Outlook URLs
const OUTLOOK_URL = 'https://outlook.live.com/mail/';
const MICROSOFT_LOGIN_URL_PATTERN = 'login.microsoftonline.com';

module.exports = {
    AUTH_FILE,
    OUTLOOK_URL,
    MICROSOFT_LOGIN_URL_PATTERN,
    AUTH_DIR,
};
