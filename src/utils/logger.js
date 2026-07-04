const chalk = require('chalk');
const path = require('path');
const fs = require('fs-extra');

/**
 * Logger utility for consistent output formatting
 */
class Logger {
    constructor() {
        this.dumpDir = null;
    }

    /**
     * Get or create the dump directory for debug files
     */
    async getDumpDir() {
        if (!this.dumpDir) {
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
            this.dumpDir = path.join(process.cwd(), 'diag-dumps', `dump-${timestamp}`);
            await fs.ensureDir(this.dumpDir);
        }
        return this.dumpDir;
    }

    /**
     * Get display path for dump files
     */
    getDumpDisplayPath() {
        if (!this.dumpDir) return '<not set>';
        return path.relative(process.cwd(), this.dumpDir);
    }

    /**
     * Log an info message
     */
    info(message) {
        console.log(chalk.blue('[INFO]') + ` ${message}`);
    }

    /**
     * Log a debug message (only if DEBUG is set)
     */
    debug(message) {
        if (process.env.DEBUG) {
            console.log(chalk.gray('[DEBUG]') + ` ${message}`);
        }
    }

    /**
     * Log a success/warning message
     */
    warn(message) {
        console.log(chalk.yellow('[WARN]') + ` ${message}`);
    }

    /**
     * Log an error message
     */
    error(message, error = null) {
        console.error(chalk.red('[ERROR]') + ` ${message}`);
        if (error && process.env.DEBUG) {
            console.error(error);
        }
    }

    /**
     * Log a success message
     */
    success(message) {
        console.log(chalk.green('[SUCCESS]') + ` ${message}`);
    }
}

module.exports = new Logger();
