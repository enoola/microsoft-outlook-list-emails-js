# microsoft-outlook-list-emails

List Microsoft Outlook emails via Playwright.

## Installation

```bash
cd microsoft-outlook-list-emails-js
npm install
```

## Usage

### Prerequisites

You need to authenticate first using the `microsoft-webauth-playwright` tool:

```bash
cd ../microsoft-webauth-playwright-js
node src/index.js login --against outlook
```

This will create an `auth.json` file in the webauth project directory.

### List Emails

```bash
cd microsoft-outlook-list-emails-js
node src/index.js list --auth-file ../microsoft-webauth-playwright-js/auth.json --max-results 10
```

### Options

- `--auth-file <path>` - Path to authentication JSON file (required)
- `--notheadless` - Run in visible browser mode for debugging
- `--dodump` - Dump HTML content to files for debugging
- `--max-results <n>` - Limit number of results returned (default: no limit)

### Output Format

JSON array with objects containing:
- `receivedDate` - Email received date
- `subject` - Email subject
- `sender` - Email sender
- `firstBodyLine` - First line of email body
