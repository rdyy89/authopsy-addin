# Authopsy - Outlook Add-in

Authopsy is an Outlook add-in that analyzes email authentication status including DMARC, DKIM, and SPF verification.

## Features

- **DMARC Analysis**: Verifies DMARC authentication status
- **DKIM Verification**: Checks DKIM signature validation
- **SPF Authentication**: Analyzes SPF record compliance
- **Visual Indicators**: Clear pass/fail/unknown status with icons
- **Detailed Information**: Click on any result to see detailed explanations
- **Works Everywhere**: Compatible with Outlook Web, Outlook Desktop, and Outlook Mobile

## Installation

### Option 1: Install from Manifest (Sideloading)

1. Download the `manifest.xml` file from this repository
2. In Outlook Web:
   - Go to Settings (gear icon) → View all Outlook settings
   - Navigate to Mail → Manage add-ins
   - Click "Add from file" and upload the manifest.xml
3. In Outlook Desktop:
   - Go to File → Manage Add-ins
   - Click "My add-ins" → "Add a custom add-in" → "Add from file"
   - Select the manifest.xml file

### Option 2: Direct URL Installation

Use this URL to install directly: `https://rdyy89.github.io/authopsy-addin/manifest.xml`

## Usage

1. Open any received email in Outlook
2. Click the "Authopsy" button in the ribbon or add-in panel
3. The add-in will analyze the email headers and display:
   - DMARC status with icon
   - DKIM status with icon  
   - SPF status with icon
4. Click the "Details" button next to any result for more information

## Status Icons

- 🟢 **Pass**: Authentication succeeded
- 🔴 **Fail**: Authentication failed
- 🟡 **Unknown**: Status could not be determined

## Technical Details

This add-in analyzes email headers including:
- `Authentication-Results`
- `ARC-Authentication-Results`
- `X-MS-Exchange-Authentication-Results`
- `DKIM-Signature`
- `Received-SPF`

## Development

To run locally for development:

```bash
npm install
npm start
```

The add-in will be served at `http://localhost:3000`

### File Structure

```
authopsy-addin/
├── manifest.xml          # Add-in manifest
├── taskpane.html         # Main UI
├── taskpane.css          # Styling
├── taskpane.js           # Main functionality
├── commands.html         # Commands page
├── commands.js           # Command handlers
├── package.json          # Node.js dependencies
├── icons/               # Icon assets
│   ├── authopsy.png     # Main add-in icon
│   ├── pass.png         # Pass status icon
│   ├── fail.png         # Fail status icon
│   └── unknown.png      # Unknown status icon
└── README.md            # This file
```

## Supported Platforms

- ✅ Outlook on the Web
- ✅ Outlook 2016 or later (Windows)
- ✅ Outlook 2016 or later (Mac)
- ✅ Outlook Mobile

## Privacy & Security

This add-in:
- Only reads email headers for analysis
- Does not store or transmit any email content
- Processes all data locally in your browser
- No external API calls for authentication analysis

## Contributing

1. Fork this repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

MIT License - see LICENSE file for details

## Support

For issues and questions, please visit: https://github.com/rdyy89/authopsy-addin/issues
