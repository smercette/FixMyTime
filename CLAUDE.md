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
The template currently only highlights selected cells in yellow. The commands.ts file contains placeholder Outlook code that should be replaced with Excel-specific functionality.

### Testing
No test framework is currently configured. Consider adding Jest or Mocha for unit tests when implementing features.

## Development Workflow

### Repository Management
- GitHub repository: https://github.com/smercette/FixMyTime
- Keep Claude.md file updated as we go (ie, at every commit)
- Created with comprehensive .gitignore for Node.js/TypeScript projects
- Uses gh CLI for GitHub operations