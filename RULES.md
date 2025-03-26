# rubberduck-mcp Implementation Rules & Guidelines

## Architecture Rules

1. **Separation of Concerns**
   - Maintain distinct layers for MCP protocol handling, COM interop, and business logic
   - Create clear boundaries between Rubberduck API integration and MCP interface exposure
   - Implement an adapter pattern to isolate COM dependency from core logic

2. **Error Propagation**
   - Never expose raw COM exceptions to MCP clients
   - Transform all errors into standardized MCP error codes with contextual messages
   - Implement global error boundaries at transport interfaces

3. **Resource Management**
   - Release COM objects immediately after use to prevent memory leaks
   - Implement reference counting for shared COM objects
   - Use disposal patterns with try/finally blocks for all COM interactions

## COM Interoperability Rules

1. **Interface Stability**
   - Create resilient wrappers around Rubberduck COM interfaces to handle version differences
   - Implement feature detection for optional Rubberduck capabilities
   - Provide fallback implementations where possible for missing features

2. **State Management**
   - Maintain minimal state between operations
   - Reconnect to COM objects when needed rather than maintaining long-lived references
   - Validate COM object state before performing operations

3. **Process Isolation**
   - Run COM operations in a separate process when possible to prevent crashes affecting the MCP server
   - Implement timeouts for all COM operations
   - Provide retry mechanisms for transient COM failures

## MCP Protocol Implementation

1. **Schema Validation**
   - Validate all inputs against published schemas before processing
   - Return clear validation errors with specific field references
   - Maintain schema version compatibility with the MCP specification

2. **Resource Freshness**
   - Implement proper caching with TTL for expensive resources
   - Provide resource update notifications when underlying data changes
   - Include timestamps in all resource responses

3. **Tool Execution**
   - Support cancellation for long-running tools
   - Provide progress updates for operations exceeding 1 second
   - Implement concurrency limits for resource-intensive operations

## TypeScript Development Standards

1. **Type Safety**
   - Use strict TypeScript configuration with `noImplicitAny` and `strictNullChecks`
   - Define explicit interfaces for all COM objects and return types
   - Avoid type assertions except in clearly documented edge cases

2. **Async Patterns**
   - Use async/await exclusively for asynchronous code
   - Implement proper error handling in all async functions
   - Return typed Promises for all asynchronous operations

3. **Code Organization**
   - Use ES modules with explicit imports/exports
   - Organize code by feature rather than by type
   - Keep files focused on single responsibilities

## Testing Requirements

1. **Unit Testing**
   - Achieve 90%+ code coverage for core functionality
   - Mock all COM interfaces for deterministic testing
   - Include error case testing for all public methods

2. **Integration Testing**
   - Implement tests with real Rubberduck instances when possible
   - Create test VBA projects for end-to-end verification
   - Test versioning compatibility with multiple Rubberduck releases

3. **Performance Testing**
   - Benchmark core operations for regression testing
   - Test memory usage patterns during extended operations
   - Verify responsiveness under concurrent usage

## Documentation Standards

1. **API Documentation**
   - Document all public interfaces with JSDoc comments
   - Include usage examples for all tools and resources
   - Document error codes and recovery strategies

2. **Implementation Notes**
   - Document COM interaction quirks and workarounds
   - Maintain a decision log for architectural choices
   - Include references to Rubberduck documentation where relevant

3. **User Guides**
   - Provide clear installation and configuration instructions
   - Include troubleshooting guides for common issues
   - Document example flows for typical use cases

## Security Considerations

1. **Authentication**
   - Implement configurable API key validation
   - Support environment variable configuration for secrets
   - Validate all authentication on every request

2. **Input Sanitization**
   - Sanitize all VBA code input to prevent injection attacks
   - Validate paths to prevent directory traversal
   - Implement strict validation for Git operation parameters

3. **Operation Limits**
   - Implement rate limiting for resource-intensive operations
   - Add configurable timeout limits for all operations
   - Provide mechanisms to abort long-running operations

## Deployment Guidelines

1. **Package Management**
   - Use semantic versioning for releases
   - Include explicit dependency pinning
   - Document Node.js version requirements

2. **Configuration**
   - Support environment variable configuration
   - Provide sensible defaults for all settings
   - Include configuration validation on startup

3. **Logging**
   - Implement structured logging with levels
   - Include correlation IDs for request tracking
   - Avoid logging sensitive information

## Performance Optimization

1. **Caching Strategy**
   - Cache expensive COM operation results
   - Implement cache invalidation for relevant Git operations
   - Use memory-efficient data structures for large projects

2. **Resource Efficiency**
   - Minimize COM object creation and destruction
   - Implement lazy loading for infrequently used features
   - Release memory aggressively when not in active use

3. **Concurrency Management**
   - Process concurrent requests efficiently where possible
   - Queue operations that require exclusive access
   - Provide timeout mechanisms for deadlock prevention
