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
- `emailId` - Email unique item ID (can be used for deep linking)
- `conversationId` - Conversation unique ID
- `receivedDate` - Email received date
- `subject` - Email subject
- `sender` - Email sender
- `firstBodyLine` - First line of email body
- `status` - Status of email ("read" or "unread")
- `category` - Category of email ("Focus" or "Other")

### Deep Linking to Specific Emails

To directly access an email or conversation, construct the URL using one of the following working formats (the ID **must** be URL-encoded, e.g., replacing `=` with `%3D`, `/` with `%2F`, etc.):

1. **Official Deep Link Format (via `emailId`)**:
   `https://outlook.live.com/mail/deeplink/read/?ItemID=<URL_ENCODED_EMAIL_ID>`
   *Note: This format uses the individual message ID (`emailId` starting with `AQAAB...` or `AQMk...`) and redirects to the active account context.*

2. **Direct Inbox URL Format (via `conversationId`)**:
   `https://outlook.live.com/mail/inbox/id/<URL_ENCODED_CONVERSATION_ID>`
   or
   `https://outlook.live.com/mail/0/inbox/id/<URL_ENCODED_CONVERSATION_ID>`
   *Note: The standard `/inbox/id/` path route in OWA only accepts the **`conversationId`** (starting with `AQQk...`). Attempting to use the individual message ID (`emailId`) with the `/inbox/id/` path will result in a redirect to the main mail folder.*

### Conversation View vs. Individual Emails

By default, Outlook groups emails into conversation threads. This affects how many entries the script scrapes:

- **Grouped by Conversation (Default)**: OWA only renders one row per conversation thread in the UI. If you have 527 total messages grouped into 191 threads, the script will only see and scrape those 191 conversation rows.
- **Show as Individual Messages**: OWA renders a separate row for every single message. In this mode, the script will scrape all 527 individual emails.

To change this setting to list all individual emails:
1. Open Outlook Web App in your browser.
2. Click the **Gear Icon (Settings)** in the top-right corner.
3. Go to **Mail** -> **Layout**.
4. Under **Message organization**, select **Show email as individual messages** (instead of *Show email grouped by conversation*).
5. Run the scrape script again.
