# rubberduck-mcp Project Structure

This document outlines a recommended directory structure and key files for implementing the rubberduck-mcp TypeScript MCP server.

## Directory Structure

```
rubberduck-mcp/
├── .github/                        # GitHub-specific files
│   ├── workflows/                  # CI/CD workflows
│   │   └── main.yml                # Main workflow for build, test, and publish
│   └── ISSUE_TEMPLATE/             # Issue templates
├── src/                            # Source code
│   ├── index.ts                    # Main entry point
│   ├── server/                     # MCP server implementation
│   │   ├── index.ts                # Server initialization and configuration
│   │   ├── tools.ts                # MCP tools definitions and implementations
│   │   └── resources.ts            # MCP resources definitions and implementations
│   ├── com/                        # COM interoperability layer
│   │   ├── index.ts                # COM layer initialization
│   │   ├── rubberduck.ts           # Rubberduck COM interface wrappers
│   │   ├── vbe.ts                  # VBE COM interface wrappers
│   │   └── utils.ts                # COM utility functions
│   ├── git/                        # Git integration
│   │   ├── index.ts                # Git operations wrapper
│   │   └── commands.ts             # Git command implementations
│   └── types/                      # TypeScript type definitions
│       ├── index.ts                # Type exports
│       ├── rubberduck.ts           # Rubberduck interface types
│       ├── vbe.ts                  # VBE interface types
│       └── mcp.ts                  # MCP-specific types
├── test/                           # Test files
│   ├── unit/                       # Unit tests
│   │   ├── server.test.ts          # Server tests
│   │   ├── com.test.ts             # COM layer tests
│   │   └── git.test.ts             # Git integration tests
│   ├── integration/                # Integration tests
│   │   └── e2e.test.ts             # End-to-end tests
│   └── mocks/                      # Test mocks
│       ├── com.ts                  # COM object mocks
│       └── git.ts                  # Git command mocks
├── docs/                           # Documentation
│   ├── api/                        # API documentation
│   ├── examples/                   # Example usage
│   └── guides/                     # User guides
├── scripts/                        # Build and utility scripts
│   ├── build.js                    # Build script
│   └── release.js                  # Release script
├── RULES.md                        # Implementation rules
├── DOCUMENTATION_LINKS.md          # Documentation references
├── README.md                       # Project overview
├── LICENSE                         # License file
├── package.json                    # NPM package definition
├── tsconfig.json                   # TypeScript configuration
├── .eslintrc.js                    # ESLint configuration
├── .prettierrc                     # Prettier configuration
└── jest.config.js                  # Jest test configuration
```

## Key Configuration Files

### package.json

```json
{
  "name": "rubberduck-mcp",
  "version": "0.1.0",
  "description": "An MCP server for integrating AI assistants with Rubberduck VBA",
  "type": "module",
  "main": "dist/index.js",
  "types": "dist/index.d.ts",
  "bin": {
    "rubberduck-mcp": "./dist/index.js"
  },
  "files": [
    "dist"
  ],
  "scripts": {
    "build": "tsc",
    "test": "jest",
    "lint": "eslint src/**/*.ts",
    "format": "prettier --write src/**/*.ts",
    "start": "node dist/index.js",
    "dev": "ts-node-dev --respawn src/index.ts",
    "prepublishOnly": "npm run build"
  },
  "keywords": [
    "mcp",
    "model-context-protocol",
    "rubberduck",
    "vba",
    "git",
    "source-control"
  ],
  "author": "",
  "license": "MIT",
  "dependencies": {
    "@modelcontextprotocol/sdk": "^0.2.0",
    "node-win32com": "^0.1.0"
  },
  "devDependencies": {
    "@types/jest": "^29.5.0",
    "@types/node": "^18.15.11",
    "@typescript-eslint/eslint-plugin": "^5.57.1",
    "@typescript-eslint/parser": "^5.57.1",
    "eslint": "^8.37.0",
    "jest": "^29.5.0",
    "prettier": "^2.8.7",
    "ts-jest": "^29.1.0",
    "ts-node": "^10.9.1",
    "ts-node-dev": "^2.0.0",
    "typescript": "^5.0.3"
  },
  "engines": {
    "node": ">=16.0.0"
  }
}
```

### tsconfig.json

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "esModuleInterop": true,
    "forceConsistentCasingInFileNames": true,
    "strict": true,
    "skipLibCheck": true,
    "declaration": true,
    "sourceMap": true,
    "outDir": "dist",
    "rootDir": "src",
    "noImplicitAny": true,
    "strictNullChecks": true,
    "noImplicitThis": true,
    "alwaysStrict": true
  },
  "include": ["src/**/*"],
  "exclude": ["node_modules", "dist", "test"]
}
```

### Main Server File (src/index.ts)

```typescript
#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { initializeTools } from './server/tools.js';
import { initializeResources } from './server/resources.js';

async function main() {
  try {
    // Initialize the MCP server
    const server = new Server(
      {
        name: 'rubberduck-mcp',
        version: '0.1.0',
      },
      {
        capabilities: {
          tools: {},
          resources: {},
        },
      }
    );

    // Set up tools and resources
    initializeTools(server);
    initializeResources(server);

    // Connect to stdio transport
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    console.error('Rubberduck MCP server running on stdio');

    // Handle graceful shutdown
    process.on('SIGINT', async () => {
      await server.close();
      process.exit(0);
    });
  } catch (error) {
    console.error('Failed to start server:', error);
    process.exit(1);
  }
}

main().catch(console.error);
```

## Implementation Notes

1. The structure separates concerns as per the rules in RULES.md:
   - MCP server implementation in `src/server/`
   - COM interoperability in `src/com/`
   - Git integration in `src/git/`
   - Type definitions in `src/types/`

2. Follow these initial implementation steps:
   - Create the COM interoperability layer first
   - Implement COM object lifecycle management
   - Create type definitions for Rubberduck interfaces
   - Implement MCP tools and resources using the COM layer
   - Add error handling and authentication

3. Testing approach:
   - Unit tests with mocked COM objects
   - Integration tests with real Rubberduck instances (when available)
   - End-to-end tests for complete workflows

4. Documentation:
   - API documentation using JSDoc comments
   - Example usage in the docs/examples directory
   - User guides in the docs/guides directory
