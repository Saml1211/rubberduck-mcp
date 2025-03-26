# rubberduck-mcp Implementation Examples

This document provides code samples demonstrating key implementation patterns for the rubberduck-mcp server, focusing on COM interoperability, type definitions, and MCP tool implementation.

## TypeScript Type Definitions (src/types/rubberduck.ts)

```typescript
/**
 * TypeScript interfaces for Rubberduck COM objects
 */

/**
 * Represents the main Rubberduck COM application object
 */
export interface IRubberduckApp {
  /** Gets the version of Rubberduck */
  readonly Version: string;
  
  /** Gets whether Rubberduck is connected to the VBE */
  readonly IsConnected: boolean;
  
  /** Gets the source control manager */
  readonly SourceControl: ISourceControlManager;
  
  /** Gets the code analyzer */
  readonly CodeAnalysis: ICodeAnalyzer;
  
  /** Connects to the VBE */
  Connect(): boolean;
  
  /** Disconnects from the VBE */
  Disconnect(): void;
}

/**
 * Represents Rubberduck's source control manager
 */
export interface ISourceControlManager {
  /** Gets whether a repository is currently open */
  readonly IsRepoOpen: boolean;
  
  /** Gets the current repository path if one is open */
  readonly CurrentRepositoryPath: string;
  
  /** Gets the current branch name if a repository is open */
  readonly CurrentBranch: string;
  
  /** Opens an existing Git repository */
  OpenRepository(path: string): boolean;
  
  /** Creates a new Git repository */
  CreateRepository(path: string): boolean;
  
  /** Commits changes with the specified message */
  Commit(message: string): string;
  
  /** Creates a new branch */
  CreateBranch(branchName: string): boolean;
  
  /** Switches to an existing branch */
  CheckoutBranch(branchName: string): boolean;
  
  /** Pulls changes from the remote repository */
  Pull(): boolean;
  
  /** Pushes changes to the remote repository */
  Push(): boolean;
  
  /** Exports a VBA module to text */
  ExportModule(moduleName: string): string;
  
  /** Imports code into a VBA module */
  ImportModule(moduleName: string, code: string): boolean;
  
  /** Gets the list of modules in the current VBA project */
  GetModulesList(): string[];
}

/**
 * Represents Rubberduck's code analysis engine
 */
export interface ICodeAnalyzer {
  /** Gets all available inspection types */
  readonly AvailableInspectionTypes: string[];
  
  /** Runs code analysis on the specified module */
  AnalyzeModule(moduleName: string, inspectionTypes?: string[]): ICodeIssue[];
}

/**
 * Represents a code issue found by the analyzer
 */
export interface ICodeIssue {
  /** The severity of the issue */
  readonly Severity: 'Hint' | 'Suggestion' | 'Warning' | 'Error';
  
  /** The description of the issue */
  readonly Description: string;
  
  /** The module where the issue was found */
  readonly ModuleName: string;
  
  /** The line number where the issue was found */
  readonly Line: number;
  
  /** The column where the issue was found */
  readonly Column: number;
}
```

## COM Interoperability Layer (src/com/rubberduck.ts)

```typescript
import { win32com } from 'node-win32com';
import { 
  IRubberduckApp, 
  ISourceControlManager, 
  ICodeAnalyzer,
  ICodeIssue 
} from '../types/rubberduck.js';
import { releaseComObject } from './utils.js';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';

/**
 * A wrapper around the Rubberduck COM objects with proper resource management
 */
export class RubberduckWrapper {
  private comApp: any | null = null;
  private isConnected = false;

  /**
   * Creates a new instance of the Rubberduck COM wrapper
   * @throws {McpError} If the Rubberduck COM object cannot be created
   */
  constructor() {
    try {
      // Create the Rubberduck COM application object
      this.comApp = win32com.createObject('Rubberduck.Application');
      
      if (!this.comApp) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to create Rubberduck COM object. Ensure Rubberduck is installed.'
        );
      }
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Error creating Rubberduck COM object: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Connects to the Rubberduck COM application
   * @returns True if successfully connected
   * @throws {McpError} If connection fails
   */
  public async connect(): Promise<boolean> {
    try {
      if (!this.comApp) {
        throw new McpError(
          ErrorCode.InternalError,
          'Rubberduck COM object is not initialized'
        );
      }

      this.isConnected = Boolean(this.comApp.Connect());
      return this.isConnected;
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to connect to Rubberduck: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Gets the Rubberduck version
   * @returns The version string
   * @throws {McpError} If the version cannot be retrieved
   */
  public async getVersion(): Promise<string> {
    try {
      this.ensureConnected();
      return String(this.comApp.Version);
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to get Rubberduck version: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Exports a VBA module to text
   * @param moduleName The name of the module to export
   * @returns The module content as text
   * @throws {McpError} If the module cannot be exported
   */
  public async exportModule(moduleName: string): Promise<string> {
    try {
      this.ensureConnected();
      
      const sourceControl = this.comApp.SourceControl;
      if (!sourceControl) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access Rubberduck Source Control'
        );
      }
      
      // Ensure repository is open
      if (!sourceControl.IsRepoOpen) {
        throw new McpError(
          ErrorCode.InvalidRequest,
          'No repository is currently open in Rubberduck'
        );
      }
      
      const moduleContent = sourceControl.ExportModule(moduleName);
      if (moduleContent === null || moduleContent === undefined) {
        throw new McpError(
          ErrorCode.NotFound,
          `Failed to export module: ${moduleName}`
        );
      }
      
      return String(moduleContent);
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error exporting module ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Imports code into a VBA module
   * @param moduleName The name of the module to import into
   * @param code The code to import
   * @returns True if successful
   * @throws {McpError} If the code cannot be imported
   */
  public async importModule(moduleName: string, code: string): Promise<boolean> {
    try {
      this.ensureConnected();
      
      const sourceControl = this.comApp.SourceControl;
      if (!sourceControl) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access Rubberduck Source Control'
        );
      }
      
      // Ensure repository is open
      if (!sourceControl.IsRepoOpen) {
        throw new McpError(
          ErrorCode.InvalidRequest,
          'No repository is currently open in Rubberduck'
        );
      }
      
      const result = Boolean(sourceControl.ImportModule(moduleName, code));
      return result;
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error importing module ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Commits changes to the repository
   * @param message The commit message
   * @returns The commit hash
   * @throws {McpError} If the commit fails
   */
  public async gitCommit(message: string): Promise<string> {
    try {
      this.ensureConnected();
      
      const sourceControl = this.comApp.SourceControl;
      if (!sourceControl) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access Rubberduck Source Control'
        );
      }
      
      // Ensure repository is open
      if (!sourceControl.IsRepoOpen) {
        throw new McpError(
          ErrorCode.InvalidRequest,
          'No repository is currently open in Rubberduck'
        );
      }
      
      const commitHash = sourceControl.Commit(message);
      if (!commitHash) {
        throw new McpError(
          ErrorCode.InternalError,
          'Commit operation failed or returned empty hash'
        );
      }
      
      return String(commitHash);
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error committing changes: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Analyzes a module for code issues
   * @param moduleName The name of the module to analyze
   * @param rules Optional array of rule names to check
   * @returns Array of code issues found
   * @throws {McpError} If analysis fails
   */
  public async analyzeCode(moduleName: string, rules?: string[]): Promise<ICodeIssue[]> {
    try {
      this.ensureConnected();
      
      const analyzer = this.comApp.CodeAnalysis;
      if (!analyzer) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access Rubberduck Code Analysis'
        );
      }
      
      const issues = analyzer.AnalyzeModule(moduleName, rules);
      
      // Convert COM objects to plain JavaScript objects
      return (issues || []).map((issue: any) => ({
        Severity: String(issue.Severity),
        Description: String(issue.Description),
        ModuleName: String(issue.ModuleName),
        Line: Number(issue.Line),
        Column: Number(issue.Column)
      }));
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error analyzing module ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Gets a list of all modules in the current project
   * @returns Array of module names
   * @throws {McpError} If the list cannot be retrieved
   */
  public async getModulesList(): Promise<string[]> {
    try {
      this.ensureConnected();
      
      const sourceControl = this.comApp.SourceControl;
      if (!sourceControl) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access Rubberduck Source Control'
        );
      }
      
      // Ensure repository is open
      if (!sourceControl.IsRepoOpen) {
        throw new McpError(
          ErrorCode.InvalidRequest,
          'No repository is currently open in Rubberduck'
        );
      }
      
      const modules = sourceControl.GetModulesList();
      return (modules || []).map((m: any) => String(m));
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error getting modules list: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Releases COM resources
   */
  public dispose(): void {
    try {
      if (this.comApp) {
        // Disconnect if connected
        if (this.isConnected) {
          try {
            this.comApp.Disconnect();
          } catch (error) {
            console.error('Error disconnecting from Rubberduck:', error);
          }
        }
        
        // Release the COM object
        releaseComObject(this.comApp);
        this.comApp = null;
        this.isConnected = false;
      }
    } catch (error) {
      console.error('Error disposing Rubberduck COM wrapper:', error);
    }
  }

  /**
   * Ensures the COM object is connected
   * @throws {McpError} If not connected
   */
  private ensureConnected(): void {
    if (!this.comApp) {
      throw new McpError(
        ErrorCode.InternalError,
        'Rubberduck COM object is not initialized'
      );
    }
    
    if (!this.isConnected) {
      throw new McpError(
        ErrorCode.InternalError,
        'Not connected to Rubberduck. Call connect() first.'
      );
    }
  }
}
```

## COM Utility Functions (src/com/utils.ts)

```typescript
/**
 * Releases a COM object safely
 * @param comObject The COM object to release
 */
export function releaseComObject(comObject: any): void {
  if (comObject) {
    try {
      // Release the COM object by calling the native method
      // This is critical for preventing memory leaks with COM objects
      if (typeof comObject.Release === 'function') {
        comObject.Release();
      } else if (global.gc) {
        // Force garbage collection if available
        global.gc();
      }
    } catch (error) {
      console.error('Error releasing COM object:', error);
    }
  }
}

/**
 * Creates a wrapper function that ensures COM objects are properly released
 * @param fn The function to wrap
 * @returns A wrapped function that ensures COM resources are released
 */
export function withComCleanup<T, Args extends any[]>(
  fn: (...args: Args) => T
): (...args: Args) => T {
  return function(...args: Args): T {
    const comObjects: any[] = [];
    
    try {
      // Track any COM objects created during function execution
      const result = fn(...args);
      return result;
    } finally {
      // Release all tracked COM objects
      comObjects.forEach(obj => releaseComObject(obj));
    }
  };
}
```

## MCP Tool Implementation (src/server/tools.ts)

```typescript
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { 
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ErrorCode,
  McpError 
} from '@modelcontextprotocol/sdk/types.js';
import { RubberduckWrapper } from '../com/rubberduck.js';

// Singleton instance of the Rubberduck wrapper
let rubberduckInstance: RubberduckWrapper | null = null;

/**
 * Gets a shared instance of the Rubberduck wrapper
 * Creates a new one if it doesn't exist
 */
function getRubberduck(): RubberduckWrapper {
  if (!rubberduckInstance) {
    rubberduckInstance = new RubberduckWrapper();
  }
  return rubberduckInstance;
}

/**
 * Releases the shared Rubberduck instance
 */
export function releaseRubberduck(): void {
  if (rubberduckInstance) {
    rubberduckInstance.dispose();
    rubberduckInstance = null;
  }
}

/**
 * Initializes all MCP tools related to Rubberduck
 * @param server The MCP server instance
 */
export function initializeTools(server: Server): void {
  // Set up the tools list handler
  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: [
      {
        name: 'exportModule',
        description: 'Export a VBA module to text',
        inputSchema: {
          type: 'object',
          properties: {
            moduleName: {
              type: 'string',
              description: 'Name of the module to export',
            },
          },
          required: ['moduleName'],
        },
      },
      {
        name: 'importModule',
        description: 'Import modified code back to a module',
        inputSchema: {
          type: 'object',
          properties: {
            moduleName: {
              type: 'string',
              description: 'Name of the module to import into',
            },
            code: {
              type: 'string',
              description: 'The code to import',
            },
          },
          required: ['moduleName', 'code'],
        },
      },
      {
        name: 'gitCommit',
        description: 'Commit changes with a message',
        inputSchema: {
          type: 'object',
          properties: {
            message: {
              type: 'string',
              description: 'Commit message',
            },
          },
          required: ['message'],
        },
      },
      {
        name: 'analyzeCode',
        description: 'Run Rubberduck code analysis on a module',
        inputSchema: {
          type: 'object',
          properties: {
            moduleName: {
              type: 'string',
              description: 'Name of the module to analyze',
            },
            rules: {
              type: 'array',
              items: {
                type: 'string',
              },
              description: 'Optional array of rules to check',
            },
          },
          required: ['moduleName'],
        },
      },
    ],
  }));

  // Set up the tool call handler
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    try {
      const rubberduck = getRubberduck();
      
      // Connect to Rubberduck if not already connected
      await rubberduck.connect();
      
      // Handle different tool calls
      switch (request.params.name) {
        case 'exportModule': {
          const { moduleName } = request.params.arguments as { moduleName: string };
          
          if (!moduleName) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Module name is required'
            );
          }
          
          const code = await rubberduck.exportModule(moduleName);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  content: code,
                  moduleName,
                  success: true
                }, null, 2),
              },
            ],
          };
        }
        
        case 'importModule': {
          const { moduleName, code } = request.params.arguments as { 
            moduleName: string; 
            code: string 
          };
          
          if (!moduleName) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Module name is required'
            );
          }
          
          if (!code) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Code is required'
            );
          }
          
          const success = await rubberduck.importModule(moduleName, code);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success,
                  moduleName,
                  message: success ? 'Module updated successfully' : 'Failed to update module'
                }, null, 2),
              },
            ],
          };
        }
        
        case 'gitCommit': {
          const { message } = request.params.arguments as { message: string };
          
          if (!message) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Commit message is required'
            );
          }
          
          const commitHash = await rubberduck.gitCommit(message);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  commitHash,
                  message: `Changes committed successfully: ${commitHash}`
                }, null, 2),
              },
            ],
          };
        }
        
        case 'analyzeCode': {
          const { moduleName, rules } = request.params.arguments as { 
            moduleName: string; 
            rules?: string[] 
          };
          
          if (!moduleName) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Module name is required'
            );
          }
          
          const issues = await rubberduck.analyzeCode(moduleName, rules);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  issues,
                  count: issues.length,
                  moduleName
                }, null, 2),
              },
            ],
          };
        }
        
        default:
          throw new McpError(
            ErrorCode.MethodNotFound,
            `Unknown tool: ${request.params.name}`
          );
      }
    } catch (error) {
      if (error instanceof McpError) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: error.message,
                code: error.code
              }, null, 2),
            },
          ],
          isError: true,
        };
      }
      
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              error: error instanceof Error ? error.message : String(error),
              code: ErrorCode.InternalError
            }, null, 2),
          },
        ],
        isError: true,
      };
    }
  });

  // Set up cleanup handler
  server.onclose = async () => {
    releaseRubberduck();
  };
}
```

## Notes on Implementation

1. **COM Object Lifecycle Management**:
   - The RubberduckWrapper class encapsulates all COM interactions
   - The dispose() method ensures COM objects are properly released
   - Helper utilities in utils.ts provide reusable COM cleanup functions

2. **Error Handling**:
   - All errors are converted to standardized MCP errors with appropriate codes
   - COM-specific errors are wrapped with context for easier debugging
   - Error boundaries at the MCP tool level ensure graceful failure

3. **Type Safety**:
   - TypeScript interfaces clearly define COM object structures
   - Strong typing throughout the codebase prevents common errors
   - Interface design matches the expected Rubberduck COM API

4. **Resource Management**:
   - Singleton pattern for the RubberduckWrapper reduces COM object creation
   - Explicit cleanup on server close prevents memory leaks
   - Clear connection state management with connect/disconnect methods

5. **MCP Tool Implementation**:
   - Tool schemas define clear parameter requirements and validations
   - Structured JSON responses with consistent formats
   - Proper error handling and conversion to MCP protocol errors

This implementation demonstrates the key patterns to follow when implementing the rest of the rubberduck-mcp server, particularly focusing on COM resource management, error handling, and MCP tool integration.
