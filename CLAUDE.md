# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

FixMyTime is a Microsoft Excel Add-in built with TypeScript and the Office.js API. Currently a basic template that needs development to implement time-related functionality.

## Development Commands

### Building and Running
- `npm install` - Install dependencies
- `npm start` - Start debugging session with Office (loads manifest.xml)
- `npm run dev-server` - Run webpack dev server on https://localhost:3000
- `npm run build` - Production build
- `npm run build:dev` - Development build
- `npm run watch` - Watch mode for automatic rebuilds

### Code Quality
- `npm run lint` - Check code with office-addin-lint
- `npm run lint:fix` - Auto-fix linting issues
- `npm run validate` - Validate the manifest.xml file

### Debugging
- `npm run stop` - Stop debugging session

## Architecture

### Technology Stack
- **Office.js** - Excel API integration
- **TypeScript** - Primary language (targets ES5)
- **Webpack 5** - Bundling and build system
- **Babel** - TypeScript/ES6 transpilation
- **Office UI Fabric** - UI components

### Project Structure
- `/src/taskpane/` - Main task pane UI and logic
  - `taskpane.ts` - Core business logic
  - `taskpane.html` - UI template
  - `taskpane.css` - Styles
- `/src/commands/` - Ribbon command handlers
  - `commands.ts` - Command implementations
  - `commands.html` - Command context HTML
- `/assets/` - Add-in icons (16x16 to 128x128)
- `manifest.xml` - Office Add-in manifest

### Key Configuration Files
- `manifest.xml` - Defines add-in metadata, permissions, and UI extension points
- `webpack.config.js` - Build configuration with HTTPS dev server setup
- `tsconfig.json` - TypeScript compiler options
- `package.json` - Dependencies and scripts

## Development Notes

### HTTPS Requirement
Office Add-ins require HTTPS. The dev server runs on https://localhost:3000 with self-signed certificates managed by office-addin-dev-certs.

### Entry Points
The add-in has two webpack entry points:
1. **taskpane** - Main UI loaded in Excel's task pane
2. **commands** - Functions called from ribbon buttons

### Office.js Initialization
Always ensure Office.js is initialized before accessing Excel APIs:
```typescript
Office.onReady((info) => {
  // Your code here
});
```

### Current Implementation
The add-in now includes:
- Matter profile management with customizable formatting settings
- Quick actions (Format Spreadsheet, Add Charge Column, Color Code Rows)
- Tab-based UI with Main and Settings tabs
- Dynamic matter selection UI that moves from Main to Settings tab when a matter is loaded
- Persistent storage of matter profiles using localStorage
- **Name Standardisation Rule**: Automatically expands first names to full names based on Fee Earners list
- **Notes Column**: Tracks which rules have been applied to each row (e.g., "Name Standardised")
- **Undo Functionality**: Allows reverting Name Standardisation rule applications

### Testing
No test framework is currently configured. Consider adding Jest or Mocha for unit tests when implementing features.

## Development Workflow

### Repository Management
- GitHub repository: https://github.com/smercette/FixMyTime
- Keep Claude.md file updated as we go (ie, at every commit)
- Created with comprehensive .gitignore for Node.js/TypeScript projects
- Uses gh CLI for GitHub operations

## Data Integrity Rules

### Column Protection
- Data in the following columns should NEVER be changed by the add in: 'Name' 'Date' 'Role' 'Rate'
- The data in the Narrative and Time / Original Narrative and Original Time columns should also not be changed
- Any changes that are made by Rules should be done in the Amended Narrative and/or Amended Time columns respectively

## Rules Management

### Rule Creation and Behavior
- When adding a Rule, it should be added to the Rules dropdown on the Settings tab
- Rules must be toggleable (on/off functionality)
- A save button at the bottom of the Rules dropdown should save the Rules to the matter profile
- When amending 'Time' or 'Narrative', this refers to modifying the 'Amended Time' and/or 'Amended Narrative' columns
- When adding a rule, carefully consider:
  - Whether the rule needs to be configurable
  - If configurable, determine the specific configuration parameters
  - Ensure the rule's implementation supports toggling and saving to matter profile
- When adding Rules, please follow the approach for Name Standardisation (dropdown with a checkbox shows in the Rules dropdown).

## Interaction Guidelines

### AI Interaction Principles
- Don't be sycophantic - I don't need praise. Only agree with me if you think I'm right.

## User Interface Guidelines

### Dropdown Behavior
- All dropdown menus should be minimised by default (ie, the user should have to expand them manually).