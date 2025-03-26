/**
 * MCP tools implementation for rubberduck-mcp
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
  ErrorCode,
  McpError
} from '@modelcontextprotocol/sdk/types.js';

import { 
  RubberduckWrapper,
  RubberduckWrapperOptions 
} from '../com/rubberduck.js';
import { VBEWrapper } from '../com/vbe.js';
import { GitCommands } from '../git/commands.js';
import {
  ExportVBAModuleArgs,
  ImportVBAModuleArgs,
  AnalyzeCodeArgs,
  ExecuteRefactoringArgs,
  GitCommitArgs,
  GitBranchArgs,
  GitSyncArgs
} from '../types/mcp.js';

/**
 * Singleton instance of RubberduckWrapper
 */
let rubberduckInstance: RubberduckWrapper | null = null;

/**
 * Singleton instance of VBEWrapper
 */
let vbeInstance: VBEWrapper | null = null;

/**
 * Singleton instance of GitCommands
 */
let gitCommands: GitCommands | null = null;

/**
 * Gets a shared instance of the Rubberduck wrapper
 * @param options Rubberduck configuration options
 * @returns Instance of RubberduckWrapper
 */
function getRubberduck(options?: RubberduckWrapperOptions): RubberduckWrapper {
  if (!rubberduckInstance) {
    rubberduckInstance = new RubberduckWrapper(options);
  }
  return rubberduckInstance;
}

/**
 * Gets a shared instance of the VBE wrapper
 * @returns Instance of VBEWrapper
 */
function getVBE(): VBEWrapper {
  if (!vbeInstance) {
    vbeInstance = new VBEWrapper();
  }
  return vbeInstance;
}

/**
 * Gets a shared instance of the Git commands
 * @param options Git configuration options
 * @returns Instance of GitCommands
 */
function getGitCommands(options?: { apiKey?: string; debug?: boolean }): GitCommands {
  if (!gitCommands) {
    gitCommands = new GitCommands(getRubberduck(), options);
  }
  return gitCommands;
}

/**
 * Releases all COM resources
 */
export function releaseComResources(): void {
  if (rubberduckInstance) {
    rubberduckInstance.dispose();
    rubberduckInstance = null;
  }
  
  if (vbeInstance) {
    vbeInstance.dispose();
    vbeInstance = null;
  }
  
  gitCommands = null;
}

/**
 * Validates the API key if one is configured
 * @param rubberduck Rubberduck wrapper instance
 * @param apiKey API key to validate
 * @throws McpError if API key is invalid
 */
function validateApiKey(rubberduck: RubberduckWrapper, apiKey?: string): void {
  if (apiKey) {
    try {
      rubberduck.validateApiKey(apiKey);
    } catch (error) {
      throw new McpError(
        ErrorCode.Unauthorized,
        'Invalid API key'
      );
    }
  }
}

/**
 * Initializes all MCP tools
 * @param server MCP server instance
 * @param options Configuration options
 */
export function initializeTools(
  server: Server,
  options: { 
    apiKey?: string;
    debug?: boolean; 
    comTimeout?: number;
    maxRetries?: number;
  } = {}
): void {
  // Set up the tools list handler
  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: [
      {
        name: 'exportVBAModule',
        description: 'Exports a VBA module to text format with optional attributes',
        inputSchema: {
          type: 'object',
          properties: {
            moduleName: {
              type: 'string',
              description: 'Name of the module to export',
            },
            includeAttributes: {
              type: 'boolean',
              description: 'Whether to include module attributes',
              default: false,
            },
          },
          required: ['moduleName'],
        },
      },
      {
        name: 'importVBAModule',
        description: 'Imports VBA code back into a module, optionally creating it',
        inputSchema: {
          type: 'object',
          properties: {
            moduleName: {
              type: 'string',
              description: 'Name of the module to import into',
            },
            content: {
              type: 'string',
              description: 'The code to import',
            },
            createIfNotExists: {
              type: 'boolean',
              description: 'Whether to create the module if it doesn\'t exist',
              default: false,
            },
          },
          required: ['moduleName', 'content'],
        },
      },
      {
        name: 'analyzeCode',
        description: 'Performs Rubberduck code analysis with configurable rulesets',
        inputSchema: {
          type: 'object',
          properties: {
            target: {
              type: 'string',
              description: 'Module name or "project" for full project analysis',
            },
            rulesets: {
              type: 'array',
              items: {
                type: 'string',
              },
              description: 'Optional array of ruleset names',
            },
          },
          required: ['target'],
        },
      },
      {
        name: 'executeRefactoring',
        description: 'Executes Rubberduck refactorings on specified code',
        inputSchema: {
          type: 'object',
          properties: {
            refactoringType: {
              type: 'string',
              description: 'Type of refactoring to execute',
            },
            target: {
              type: 'object',
              properties: {
                module: {
                  type: 'string',
                  description: 'Module to refactor',
                },
                selection: {
                  type: 'object',
                  properties: {
                    startLine: { type: 'number' },
                    startColumn: { type: 'number' },
                    endLine: { type: 'number' },
                    endColumn: { type: 'number' },
                  },
                  description: 'Optional selection within the module',
                },
              },
              required: ['module'],
              description: 'Target for the refactoring',
            },
            options: {
              type: 'object',
              additionalProperties: true,
              description: 'Optional refactoring-specific options',
            },
          },
          required: ['refactoringType', 'target'],
        },
      },
      {
        name: 'gitCommit',
        description: 'Commits changes with specified message and optional file selection',
        inputSchema: {
          type: 'object',
          properties: {
            message: {
              type: 'string',
              description: 'Commit message',
            },
            files: {
              type: 'array',
              items: {
                type: 'string',
              },
              description: 'Optional array of specific files to commit',
            },
          },
          required: ['message'],
        },
      },
      {
        name: 'gitBranch',
        description: 'Creates and/or checks out Git branches',
        inputSchema: {
          type: 'object',
          properties: {
            name: {
              type: 'string',
              description: 'Branch name',
            },
            create: {
              type: 'boolean',
              description: 'Whether to create the branch if it doesn\'t exist',
              default: false,
            },
            checkout: {
              type: 'boolean',
              description: 'Whether to checkout the branch',
              default: true,
            },
          },
          required: ['name'],
        },
      },
      {
        name: 'gitSync',
        description: 'Performs Git synchronization operations',
        inputSchema: {
          type: 'object',
          properties: {
            operation: {
              type: 'string',
              enum: ['push', 'pull', 'fetch'],
              description: 'Operation to perform',
            },
            remote: {
              type: 'string',
              description: 'Optional remote name',
            },
            branch: {
              type: 'string',
              description: 'Optional branch name',
            },
          },
          required: ['operation'],
        },
      },
    ],
  }));

  // Set up the tool call handler
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    try {
      const rubberduck = getRubberduck({
        apiKey: options.apiKey,
        debug: options.debug,
        comTimeout: options.comTimeout,
        maxRetries: options.maxRetries,
      });
      
      // Always connect to ensure we have a valid connection
      try {
        await rubberduck.connect();
      } catch (error) {
        console.error('Failed to connect to Rubberduck:', error);
        throw new McpError(
          ErrorCode.ServerUnavailable,
          'Failed to connect to Rubberduck'
        );
      }
      
      // Handle different tool calls
      switch (request.params.name) {
        case 'exportVBAModule': {
          // If API key is configured, validate it
          validateApiKey(rubberduck, options.apiKey);
          
          const { moduleName, includeAttributes = false } = 
            request.params.arguments as ExportVBAModuleArgs;
          
          if (!moduleName) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Module name is required'
            );
          }
          
          const result = await rubberduck.exportVBAModule(moduleName, includeAttributes);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2),
              },
            ],
          };
        }
        
        case 'importVBAModule': {
          // If API key is configured, validate it
          validateApiKey(rubberduck, options.apiKey);
          
          const { moduleName, content, createIfNotExists = false } = 
            request.params.arguments as ImportVBAModuleArgs;
          
          if (!moduleName) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Module name is required'
            );
          }
          
          if (content === undefined) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Content is required'
            );
          }
          
          const result = await rubberduck.importVBAModule(moduleName, content, createIfNotExists);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2),
              },
            ],
          };
        }
        
        case 'analyzeCode': {
          // If API key is configured, validate it
          validateApiKey(rubberduck, options.apiKey);
          
          const { target, rulesets } = request.params.arguments as AnalyzeCodeArgs;
          
          if (!target) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Target is required'
            );
          }
          
          const result = await rubberduck.analyzeCode(target, rulesets);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2),
              },
            ],
          };
        }
        
        case 'executeRefactoring': {
          // If API key is configured, validate it
          validateApiKey(rubberduck, options.apiKey);
          
          const { refactoringType, target, options: refactoringOptions } = 
            request.params.arguments as ExecuteRefactoringArgs;
          
          if (!refactoringType) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Refactoring type is required'
            );
          }
          
          if (!target || !target.module) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Target module is required'
            );
          }
          
          // Convert selection format if provided
          let codeSelection = undefined;
          if (target.selection) {
            codeSelection = {
              StartLine: target.selection.startLine,
              StartColumn: target.selection.startColumn,
              EndLine: target.selection.endLine,
              EndColumn: target.selection.endColumn
            };
          }
          
          const result = await rubberduck.executeRefactoring(
            refactoringType,
            target.module,
            codeSelection,
            refactoringOptions
          );
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: result.Success,
                  result: result.Description,
                  errorMessage: result.ErrorMessage
                }, null, 2),
              },
            ],
          };
        }
        
        case 'gitCommit': {
          // If API key is configured, validate it
          validateApiKey(rubberduck, options.apiKey);
          
          const git = getGitCommands({ 
            apiKey: options.apiKey,
            debug: options.debug
          });
          
          const { message, files } = request.params.arguments as GitCommitArgs;
          
          if (!message) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Commit message is required'
            );
          }
          
          const result = await git.commit(message, files);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2),
              },
            ],
          };
        }
        
        case 'gitBranch': {
          // If API key is configured, validate it
          validateApiKey(rubberduck, options.apiKey);
          
          const git = getGitCommands({ 
            apiKey: options.apiKey,
            debug: options.debug
          });
          
          const { name, create = false, checkout = true } = 
            request.params.arguments as GitBranchArgs;
          
          if (!name) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Branch name is required'
            );
          }
          
          const result = await git.branch(name, create, checkout);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2),
              },
            ],
          };
        }
        
        case 'gitSync': {
          // If API key is configured, validate it
          validateApiKey(rubberduck, options.apiKey);
          
          const git = getGitCommands({ 
            apiKey: options.apiKey,
            debug: options.debug
          });
          
          const { operation, remote, branch } = request.params.arguments as GitSyncArgs;
          
          if (!operation) {
            throw new McpError(
              ErrorCode.InvalidParams,
              'Operation is required'
            );
          }
          
          const result = await git.sync(operation, remote, branch);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2),
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
}
