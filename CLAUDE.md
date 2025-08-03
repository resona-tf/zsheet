# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a Zoho CRM Widget application (ZSheet) that generates spreadsheets by replacing template variables with actual CRM data. The widget integrates with Zoho CRM and Zoho Sheet to create dynamic documents based on templates and CRM records.

## Architecture

### Key Components

1. **Frontend (app/)**
   - Single Page Application with vanilla JavaScript
   - Zoho Embedded App SDK integration
   - Template selection UI with multi-select capability
   - Progress tracking for document generation

2. **Backend (server/)**
   - Express.js HTTPS server on port 5000
   - Serves static files and plugin manifest
   - Required for Zoho widget development

3. **Core Libraries**
   - `ZohoEmbededAppSDK.min.js` - Zoho CRM integration
   - `xlsx-populate` - Excel file manipulation
   - `pdf-lib` - PDF generation
   - `i18next` - Internationalization support

### Variable Replacement System

The system uses a custom variable syntax `<Z%variable%Z>` documented in `docs/variable_format.md`:
- Basic field references: `<Z%FieldName%Z>`
- Lookup field references: `<Z%LookupField->TargetField%Z>`
- Related records: `<Z%RelatedList@->Field%Z>`
- Subform references: `<Z%Subform#->Field%Z>`
- Conditional replacements: `<Z%Field??Condition%Z>`
- Special variables for execution time, stamps, and invoice numbers

### API Integration

- **zohoapi.js**: Caching layer for Zoho CRM API calls
  - Reduces API calls through intelligent caching
  - Handles records, fields, modules, and related data
- **sheetapi.js**: Zoho Sheet API wrapper (large file, use search for specific functions)
- **varEngine.js**: Core variable replacement engine

## Development Commands

```bash
# Install dependencies
npm install

# Start development server (HTTPS on port 5000)
npm start

# The server runs at https://127.0.0.1:5000
# You must authorize the self-signed certificate in your browser
```

## Common Development Tasks

### Adding New Variable Types
1. Update variable parser in `varEngine.js`
2. Add documentation to `docs/variable_format.md`
3. Test with sample templates

### Debugging API Calls
- API call counter is implemented in `zohoapi.js`
- Check browser console for API logging
- Use `apiCounter()` function to track API usage

### Template Management
- Templates are configured through widget settings
- Multiple templates can be selected for batch processing
- Template URLs are stored in `SETTINGS.SheetTemplateUrl`

## Important Notes

1. **SSL Certificate**: The development server uses self-signed certificates (`cert.pem`, `key.pem`). You must manually trust these in your browser.

2. **Zoho Widget Constraints**: 
   - Must run on HTTPS
   - Requires proper CORS headers (already configured)
   - Widget height/width dynamically adjusted based on content

3. **API Rate Limiting**: The system implements delays for bulk operations (>30 records) to avoid Zoho API rate limits.

4. **Row Deletion Logic**: Complex algorithm in `clearingRows()` function that:
   - Preserves rows without repeat keys
   - Removes unchanged template rows
   - Maintains minimum row counts based on repeat key numbers

5. **Multi-language Support**: Uses i18next with translations in `app/translations/`