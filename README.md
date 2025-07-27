# Authopsy - Outlook Add-in

Authopsy is an Outlook add-in that analyzes email authentication status including DMARC, DKIM, and SPF verification by checking email headers.

## Features

- **📊 Header Analysis**: Analyzes email headers for authentication information
- **🔍 DMARC Analysis**: Verifies DMARC authentication status
- **🔐 DKIM Verification**: Checks DKIM signature validation
- **📧 SPF Authentication**: Analyzes SPF record compliance
- **📋 Dropdown Menu**: Access via dropdown with multiple analysis options
- **📌 Pinnable Interface**: Can be pinned for quick access in Outlook Web
- **⚡ Quick Check**: Fast header analysis with summary results
- **📱 Full Analysis**: Detailed analysis with explanations
- **🎯 Visual Indicators**: Clear pass/fail/unknown status with icons
- **🌐 Works Everywhere**: Compatible with Outlook Web, Desktop, and Mobile

## Installation

### For Work & School Accounts (Recommended)

1. **Download the manifest**: Save the `manifest.xml` file from this repository
2. **In Outlook Web** (outlook.office.com):
   - Click the Apps icon (9 dots) → Admin
   - Go to Settings → Integrated apps
   - Click "Upload custom apps" → "Upload from file"
   - Select the `manifest.xml` file
3. **In Outlook Desktop**:
   - Go to File → Manage Add-ins → "My add-ins"
   - Click "Add a custom add-in" → "Add from file"
   - Select the `manifest.xml` file

### For Personal Accounts

1. **In Outlook Web**:
   - Go to Settings (gear icon) → View all Outlook settings
   - Navigate to Mail → Manage add-ins
   - Click "Add from file" and upload `manifest.xml`
2. **In Outlook Desktop**:
   - File → Manage Add-ins → My add-ins
   - Add custom add-in → From file → Select `manifest.xml`

### Direct URL Installation

For advanced users: `https://rdyy89.github.io/authopsy-addin/manifest.xml`

## Usage

### Method 1: Dropdown Menu (Recommended)
1. Open any received email in Outlook
2. Click the "Authopsy" dropdown button in the ribbon
3. Choose from:
   - **Full Analysis**: Opens detailed panel with explanations
   - **Quick Check**: Shows summary in notification

### Method 2: Pinned Panel
1. Right-click the "Authopsy" button
2. Select "Pin" to keep the panel always visible
3. Click any email to automatically analyze headers

### What You'll See

The add-in displays:
- 🟢 **DMARC**: Pass/Fail/Unknown status with detailed explanation
- 🔐 **DKIM**: Signature verification status with details
- 📧 **SPF**: Sender authentication status with explanation
- 📊 **Summary Score**: X/3 checks passed
- 💡 **Click Details**: Get technical explanations for each result

## Status Icons

- 🟢 **Pass**: Authentication succeeded - email is legitimate
- 🔴 **Fail**: Authentication failed - potential security risk
- 🟡 **Unknown**: Status could not be determined - investigate further

## Header Analysis Details

This add-in analyzes the following email headers:
- `Authentication-Results` - Primary authentication results
- `ARC-Authentication-Results` - Authenticated Received Chain results  
- `X-MS-Exchange-Authentication-Results` - Microsoft Exchange results
- `DKIM-Signature` - Digital signature information
- `Received-SPF` - SPF validation results

### What Each Check Means:
- **DMARC**: Domain-based Message Authentication, Reporting & Conformance
- **DKIM**: DomainKeys Identified Mail (digital signature)
- **SPF**: Sender Policy Framework (IP authorization)

## Troubleshooting

### Installation Issues
- **Work Account**: Use admin installation method through IT department
- **Permission Error**: Ensure you have rights to install add-ins
- **Manifest Error**: Check that manifest.xml downloaded completely

### Analysis Issues  
- **No Results**: Email may not have authentication headers
- **Unknown Status**: Headers present but results unclear
- **Error Message**: Select a received email (not sent items)

## Development

To run locally for development:

```bash
npm install
npm start
```

The add-in will be served at `http://localhost:3000`

For testing, update manifest URLs to localhost before installation.

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
