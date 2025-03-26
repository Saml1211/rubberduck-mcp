# rubberduck-mcp

A TypeScript-based Model Context Protocol (MCP) server that bridges AI assistants with Rubberduck VBA's source control and code analysis functionality.

## Overview

rubberduck-mcp enables AI systems to access, analyze, modify, and manage VBA codebases through standardized MCP interfaces. This bridge allows AI assistants like Claude to work directly with VBA projects using Rubberduck's powerful source control and code analysis features.

## Features

- **VBA Module Management**: Export and import VBA modules as text
- **Git Integration**: Perform Git operations (commit, branch, push, pull) directly on VBA projects
- **Code Analysis**: Run Rubberduck's code analysis and get detailed results
- **Automated Refactoring**: Execute Rubberduck refactorings on VBA code
- **Project Structure**: Access and navigate VBA project structure
- **Secure Authentication**: API key validation for secure access

## Prerequisites

- Node.js 16.0.0 or later
- Rubberduck VBA Add-in installed and configured
- Windows environment (for COM interoperability)
- VBA host application (Excel, Word, etc.) with a VBA project loaded

## Installation

### From npm (when published)

```bash
npm install -g rubberduck-mcp
```

### From Source

```bash
git clone https://github.com/yourusername/rubberduck-mcp.git
cd rubberduck-mcp
npm install
npm run build
```

## Usage

### As a command-line tool:

```bash
rubberduck-mcp [options]
```

Available options:
- `--api-key <key>` - API key for authentication (can also use env var `RUBBERDUCK_API_KEY`)
- `--debug` - Enable debug logging (can also use env var `RUBBERDUCK_DEBUG=true`)
- `--com-timeout <ms>` - Timeout for COM operations in ms (can also use env var `RUBBERDUCK_COM_TIMEOUT`)
- `--max-retries <num>` - Maximum retries for COM operations (can also use env var `RUBBERDUCK_MAX_RETRIES`)
- `-h, --help` - Show help information
- `-v, --version` - Show version information

### With Claude Desktop:

1. Add to your Claude Desktop configuration file (typically `~/.claude/config.json` on macOS/Linux or `%APPDATA%\Claude\config.json` on Windows):

```json
{
  "mcpServers": {
    "rubberduck": {
      "command": "rubberduck-mcp",
      "args": [],
      "env": {
        "RUBBERDUCK_API_KEY": "your-api-key-here",
        "RUBBERDUCK_DEBUG": "true"
      }
    }
  }
}
```

2. In Claude, you can now use the following MCP tools to interact with your VBA projects:

- **exportVBAModule**: Export a VBA module to text format with optional attributes
- **importVBAModule**: Import VBA code back into a module, optionally creating it
- **analyzeCode**: Perform Rubberduck code analysis with configurable rulesets
- **executeRefactoring**: Execute Rubberduck refactorings on specified code
- **gitCommit**: Commit changes with a specified message and optional file selection
- **gitBranch**: Create and/or checkout Git branches
- **gitSync**: Perform Git synchronization operations (push, pull, fetch)

3. You can also access these MCP resources:

- **rubberduck://project-structure**: JSON representation of VBA project hierarchy
- **rubberduck://modules-list**: Array of available modules with metadata
- **rubberduck://git-status**: Current repository status with staged/unstaged changes
- **rubberduck://code-history/{moduleName}**: Commit history for specific modules
- **rubberduck://inspection-results/{moduleName}**: Latest code analysis results
- **rubberduck://refactoring-options**: Available automated refactorings with descriptions

## Example: Using rubberduck-mcp with Claude

Here are some examples of how you might use rubberduck-mcp with Claude:

1. **Export a VBA module:**

```
I'd like to see the code from the "Module1" module in my VBA project.
```

Claude can use the `exportVBAModule` tool:

```
Let me get that code for you from Module1.
```

2. **Analyze code for issues:**

```
Can you check my "ErrorHandler" module for any code quality issues?
```

Claude can use the `analyzeCode` tool:

```
I'll analyze your ErrorHandler module for potential issues.
```

3. **Commit changes to Git:**

```
Please commit my recent changes to the repository with the message "Fix error handling in validation routines"
```

Claude can use the `gitCommit` tool:

```
I'll commit your changes to Git with that message.
```

## Architecture

This project follows a layered architecture:

- **Server Layer**: Implements the Model Context Protocol server interface
- **COM Interoperability Layer**: Manages interactions with Rubberduck and VBE through COM
- **Git Integration Layer**: Provides Git operation capabilities
- **Types**: Strongly typed interfaces for all components

## Project Structure

The project has the following structure:

- **src/server/**: MCP server implementation
- **src/com/**: COM interoperability layer
- **src/git/**: Git integration
- **src/types/**: TypeScript type definitions

## Implementation Guidelines

The project adheres to strict implementation rules, ensuring:

- Proper COM resource management
- Robust error handling
- Strong type safety
- Comprehensive documentation

## Development

### Building

```bash
npm run build
```

### Testing

```bash
npm test
```

### ESLint

```bash
npm run lint
```

### Formatting

```bash
npm run format
```

## License

MIT

## Acknowledgements

- [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck) for their amazing VBA tooling
- [Model Context Protocol](https://github.com/modelcontextprotocol) for the protocol specification
