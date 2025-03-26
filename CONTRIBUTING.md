# Contributing to rubberduck-mcp

Thank you for your interest in contributing to rubberduck-mcp! This document provides guidelines and instructions for contributing to the project.

## Code of Conduct

Please be respectful and considerate of others when contributing to this project. We aim to foster an inclusive and welcoming community.

## Development Process

### Setting Up the Development Environment

1. Fork the repository
2. Clone your fork locally
3. Install dependencies: `npm install`
4. Build the project: `npm run build`

### Prerequisites

- Node.js 16.0.0 or later
- TypeScript understanding
- Rubberduck VBA Add-in installed for testing
- Windows environment (for COM interoperability)

### Development Workflow

1. Create a new branch for your feature or bugfix
2. Make your changes
3. Write or update tests
4. Run tests: `npm test`
5. Run linting: `npm run lint`
6. Format code: `npm run format`
7. Commit your changes with a clear and descriptive message
8. Push to your fork
9. Submit a pull request

## Pull Request Process

1. Ensure all tests pass and linting issues are resolved
2. Update documentation if necessary
3. Provide a clear description of the changes and the problem they solve
4. Link to any related issues
5. Wait for review and address any feedback

## Coding Standards

### TypeScript Guidelines

- Use strict TypeScript typing
- Follow the existing code style
- Document public APIs with JSDoc comments
- Write meaningful variable and function names
- Keep functions small and focused

### COM Interoperability

- Always properly release COM objects to prevent memory leaks
- Use try-finally blocks for COM resource management
- Handle COM errors appropriately
- Avoid holding references to COM objects longer than necessary

### Testing

- Write tests for new functionality
- Ensure existing tests pass
- Mock COM objects for testing

### Error Handling

- Use appropriate MCP error codes
- Provide descriptive error messages
- Don't expose raw COM exceptions to MCP clients

## Architecture Guidelines

### Separation of Concerns

- Keep layers separate:
  - MCP server layer
  - COM interoperability layer
  - Git integration layer
  - Type definitions

### Resource Management

- Follow resource cleanup patterns
- Release COM objects immediately after use
- Use the utility functions provided in `src/com/utils.ts`

### Error Propagation

- Transform all errors into standardized MCP error codes
- Provide context in error messages
- Implement proper error boundaries

## Documentation

- Update README.md if needed
- Document all public APIs
- Include examples for new functionality
- Keep documentation up-to-date with code changes

## Release Process

The release process is handled by the project maintainers. When contributing, you don't need to worry about versioning or publishing.

## Questions?

If you have any questions about contributing, please open an issue in the repository.

Thank you for your contributions!
