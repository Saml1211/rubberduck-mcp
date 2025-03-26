# Documentation Links for rubberduck-mcp Project

This document contains links to relevant documentation for developing the rubberduck-mcp TypeScript-based MCP server.

## MCP Protocol Documentation

### Core Specification and Documentation
- [Model Context Protocol GitHub Organization](https://github.com/modelcontextprotocol) - The main organization repository for MCP
- [MCP Documentation Repository](https://github.com/modelcontextprotocol/docs) - Official documentation and specification
- [MCP Introduction](https://modelcontextprotocol.io/introduction) - Overview and introduction to the protocol
- [MCP Examples](https://modelcontextprotocol.io/examples) - Example servers and implementations

### TypeScript SDK
- [MCP TypeScript SDK](https://github.com/modelcontextprotocol/typescript-sdk) - Official TypeScript SDK for MCP servers and clients
- [TypeScript SDK Documentation](https://modelcontextprotocol.io/docs/typescript-sdk) - Reference documentation for the SDK
- [Creating an MCP Server Tutorial](https://michaelwapp.medium.com/creating-a-model-context-protocol-server-a-step-by-step-guide-4c853fbf5ff2) - Step-by-step guide to creating an MCP server

### MCP Concepts
- [MCP Resources Documentation](https://modelcontextprotocol.io/docs/concepts/resources) - Documentation on MCP resources
- [MCP Prompts Documentation](https://modelcontextprotocol.io/docs/concepts/prompts) - Documentation on MCP prompts implementation
- [MCP Tools Documentation](https://modelcontextprotocol.io/docs/concepts/tools) - Documentation on MCP tools implementation

## Rubberduck VBA Documentation

### Source Control Integration
- [Rubberduck Source Control Blog Post](https://rubberduckvba.blog/tag/source-control/) - Overview of Rubberduck's source control capabilities
- [IDE-Integrated Git Source Control](https://rubberduckvba.blog/2016/03/20/ide-integrated-git-source-control/) - Details on Rubberduck's Git integration
- [Source Control API Wiki](https://github.com/rubberduck-vba/Rubberduck/wiki/Source-Control-API) - Documentation on Rubberduck's source control API
- [Rubberduck and Git Integration Question](https://stackoverflow.com/questions/43269102/how-to-manage-a-local-git-repository-using-rubberduck) - Community information on using Rubberduck with Git

### General Features
- [Rubberduck Features Overview](https://rubberduckvba.com/features) - General feature documentation for Rubberduck
- [Rubberduck GitHub Repository](https://github.com/rubberduck-vba/Rubberduck) - Main repository with source code and issues
- [Rubberduck News](https://rubberduckvba.blog/) - Official blog with updates and feature announcements

## COM Interoperability

### Node.js COM Libraries
- [node-win32com](https://github.com/idobatter/node-win32com) - Asynchronous, non-blocking win32com wrapper for Node.js
- [win32-api NPM Package](https://www.npmjs.com/package/win32-api) - FFI Definitions of Windows win32 api for Node.js
- [Node API for .NET](https://github.com/microsoft/node-api-dotnet) - Advanced interoperability between .NET and JavaScript

### TypeScript Development
- [TypeScript with npm Documentation](https://learn.microsoft.com/en-us/visualstudio/javascript/compile-typescript-code-npm?view=vs-2022) - Microsoft documentation on using TypeScript with npm
- [Introduction to TypeScript in Node.js](https://nodejs.org/en/learn/typescript/introduction) - Node.js official documentation on TypeScript integration
- [TypeScript NPM Package Publishing Guide](https://pauloe-me.medium.com/typescript-npm-package-publishing-a-beginners-guide-40b95908e69c) - Guide for publishing TypeScript packages

## Implementation Best Practices

### Error Handling and Security
- Review the RULES.md file created for specific requirements on error handling
- Implement standard MCP error codes as defined in the MCP specification

### COM Resource Management
- Focus on proper COM object lifecycle management to prevent memory leaks
- Implement reference counting and cleanup patterns

### Type Definitions
- Create strong TypeScript interfaces for all COM interactions
- Leverage TypeScript's type safety features for robust code

## Additional Resources

- [npm Documentation](https://docs.npmjs.com/cli/v9/configuring-npm/package-json/) - Official npm documentation for package.json configuration
- [Building MCP with LLMs](https://modelcontextprotocol.io/tutorials/building-mcp-with-llms) - Tutorial on building MCP servers with LLM integration
