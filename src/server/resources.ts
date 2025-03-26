/**
 * MCP resources implementation for rubberduck-mcp
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import {
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
  ListResourceTemplatesRequestSchema,
  ErrorCode,
  McpError
} from '@modelcontextprotocol/sdk/types.js';

import { 
  RubberduckWrapper,
  RubberduckWrapperOptions 
} from '../com/rubberduck.js';
import { VBEWrapper } from '../com/vbe.js';
import { GitCommands } from '../git/commands.js';
import { releaseComResources } from './tools.js';

// Import the shared instances from tools.js
// (These would be defined there or in a shared module)
let rubberduckInstance: RubberduckWrapper | null = null;
let vbeInstance: VBEWrapper | null = null;
let gitCommands: GitCommands | null = null;

/**
 * Gets a shared instance of the Rubberduck wrapper
 * @param options Configuration options
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
 * Initializes all MCP resources
 * @param server MCP server instance
 * @param options Configuration options
 */
export function initializeResources(
  server: Server,
  options: { 
    apiKey?: string;
    debug?: boolean; 
    comTimeout?: number;
    maxRetries?: number;
  } = {}
): void {
  // Set up the resources list handler
  server.setRequestHandler(ListResourcesRequestSchema, async () => ({
    resources: [
      {
        uri: 'rubberduck://project-structure',
        name: 'VBA Project Structure',
        mimeType: 'application/json',
        description: 'JSON representation of VBA project hierarchy',
      },
      {
        uri: 'rubberduck://modules-list',
        name: 'VBA Modules List',
        mimeType: 'application/json',
        description: 'Array of available modules with metadata',
      },
      {
        uri: 'rubberduck://git-status',
        name: 'Git Repository Status',
        mimeType: 'application/json',
        description: 'Current repository status with staged/unstaged changes',
      },
      {
        uri: 'rubberduck://refactoring-options',
        name: 'Available Refactorings',
        mimeType: 'application/json',
        description: 'Available automated refactorings with descriptions',
      },
    ],
  }));

  // Set up the resource templates list handler
  server.setRequestHandler(ListResourceTemplatesRequestSchema, async () => ({
    resourceTemplates: [
      {
        uriTemplate: 'rubberduck://code-history/{moduleName}',
        name: 'Module Code History',
        mimeType: 'application/json',
        description: 'Commit history for a specific module',
      },
      {
        uriTemplate: 'rubberduck://inspection-results/{moduleName}',
        name: 'Module Inspection Results',
        mimeType: 'application/json',
        description: 'Latest code analysis results for a specific module',
      },
    ],
  }));

  // Set up the read resource handler
  server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
    try {
      const rubberduck = getRubberduck({
        apiKey: options.apiKey,
        debug: options.debug,
        comTimeout: options.comTimeout,
        maxRetries: options.maxRetries,
      });
      
      // Validate API key if configured
      validateApiKey(rubberduck, options.apiKey);
      
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
      
      // Parse the URI
      const uri = request.params.uri;
      
      // Static resources
      if (uri === 'rubberduck://project-structure') {
        return await getProjectStructure(rubberduck);
      }
      else if (uri === 'rubberduck://modules-list') {
        return await getModulesList(rubberduck);
      }
      else if (uri === 'rubberduck://git-status') {
        return await getGitStatus(rubberduck);
      }
      else if (uri === 'rubberduck://refactoring-options') {
        return await getRefactoringOptions(rubberduck);
      }
      
      // Dynamic resources
      const codeHistoryMatch = uri.match(/^rubberduck:\/\/code-history\/(.+)$/);
      if (codeHistoryMatch) {
        const moduleName = decodeURIComponent(codeHistoryMatch[1]);
        return await getModuleHistory(rubberduck, moduleName);
      }
      
      const inspectionResultsMatch = uri.match(/^rubberduck:\/\/inspection-results\/(.+)$/);
      if (inspectionResultsMatch) {
        const moduleName = decodeURIComponent(inspectionResultsMatch[1]);
        return await getInspectionResults(rubberduck, moduleName);
      }
      
      throw new McpError(
        ErrorCode.NotFound,
        `Resource not found: ${uri}`
      );
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      
      throw new McpError(
        ErrorCode.InternalError,
        `Error reading resource: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  });
}

/**
 * Gets the VBA project structure
 * @param rubberduck Rubberduck wrapper instance
 * @returns Project structure as a resource response
 */
async function getProjectStructure(
  rubberduck: RubberduckWrapper
): Promise<{ contents: { uri: string; mimeType: string; text: string }[] }> {
  try {
    // Use VBE wrapper to get project information
    const vbe = getVBE();
    await vbe.connect();
    
    // Get active project details
    const projectInfo = await vbe.getActiveProject();
    
    // Get modules in the project
    const components = await vbe.getComponents();
    
    // Format the response
    const structure = {
      name: projectInfo.name,
      fileName: projectInfo.fileName,
      description: projectInfo.description,
      modules: components.map(component => ({
        name: component.name,
        type: component.type.toString(),
        lineCount: component.lineCount,
      })),
      timestamp: new Date().toISOString(),
    };
    
    return {
      contents: [
        {
          uri: 'rubberduck://project-structure',
          mimeType: 'application/json',
          text: JSON.stringify(structure, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new McpError(
      ErrorCode.InternalError,
      `Error getting project structure: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

/**
 * Gets the list of VBA modules
 * @param rubberduck Rubberduck wrapper instance
 * @returns Modules list as a resource response
 */
async function getModulesList(
  rubberduck: RubberduckWrapper
): Promise<{ contents: { uri: string; mimeType: string; text: string }[] }> {
  try {
    // Get modules list
    const modules = await rubberduck.getModulesList();
    
    // Format the response
    const modulesList = {
      modules: modules.map(module => ({
        name: module.Name,
        type: module.Type,
        path: module.Path || null,
      })),
      count: modules.length,
      timestamp: new Date().toISOString(),
    };
    
    return {
      contents: [
        {
          uri: 'rubberduck://modules-list',
          mimeType: 'application/json',
          text: JSON.stringify(modulesList, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new McpError(
      ErrorCode.InternalError,
      `Error getting modules list: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

/**
 * Gets the current Git status
 * @param rubberduck Rubberduck wrapper instance
 * @returns Git status as a resource response
 */
async function getGitStatus(
  rubberduck: RubberduckWrapper
): Promise<{ contents: { uri: string; mimeType: string; text: string }[] }> {
  try {
    // Get Git status
    const git = getGitCommands();
    const status = await git.getStatus();
    
    // Format the response
    const gitStatus = {
      currentBranch: status.CurrentBranch,
      hasChanges: status.HasChanges,
      stagedChanges: status.StagedChanges.map(change => ({
        path: change.Path,
        status: change.Status,
      })),
      unstagedChanges: status.UnstagedChanges.map(change => ({
        path: change.Path,
        status: change.Status,
      })),
      timestamp: new Date().toISOString(),
    };
    
    return {
      contents: [
        {
          uri: 'rubberduck://git-status',
          mimeType: 'application/json',
          text: JSON.stringify(gitStatus, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new McpError(
      ErrorCode.InternalError,
      `Error getting Git status: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

/**
 * Gets available refactoring options
 * @param rubberduck Rubberduck wrapper instance
 * @returns Refactoring options as a resource response
 */
async function getRefactoringOptions(
  rubberduck: RubberduckWrapper
): Promise<{ contents: { uri: string; mimeType: string; text: string }[] }> {
  try {
    // Get available refactoring types
    const types = await rubberduck.getAvailableRefactoringTypes();
    
    // This would normally include descriptions, but we're simplifying here
    const refactoringOptions = {
      refactorings: types.map(type => ({
        name: type,
        // In a real implementation, we would include more metadata
        requiresSelection: false, // Placeholder
        description: `${type} refactoring`, // Placeholder
      })),
      count: types.length,
      timestamp: new Date().toISOString(),
    };
    
    return {
      contents: [
        {
          uri: 'rubberduck://refactoring-options',
          mimeType: 'application/json',
          text: JSON.stringify(refactoringOptions, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new McpError(
      ErrorCode.InternalError,
      `Error getting refactoring options: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

/**
 * Gets commit history for a specific module
 * @param rubberduck Rubberduck wrapper instance
 * @param moduleName Module name
 * @returns Module history as a resource response
 */
async function getModuleHistory(
  rubberduck: RubberduckWrapper,
  moduleName: string
): Promise<{ contents: { uri: string; mimeType: string; text: string }[] }> {
  try {
    if (!moduleName) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        'Module name is required'
      );
    }
    
    // Get module history
    const git = getGitCommands();
    const history = await git.getModuleHistory(moduleName);
    
    // Format the response
    const moduleHistory = {
      moduleName,
      commits: history.map(commit => ({
        hash: commit.Hash,
        message: commit.Message,
        author: commit.Author,
        date: commit.Date,
      })),
      count: history.length,
      timestamp: new Date().toISOString(),
    };
    
    return {
      contents: [
        {
          uri: `rubberduck://code-history/${encodeURIComponent(moduleName)}`,
          mimeType: 'application/json',
          text: JSON.stringify(moduleHistory, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new McpError(
      ErrorCode.InternalError,
      `Error getting module history: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

/**
 * Gets inspection results for a specific module
 * @param rubberduck Rubberduck wrapper instance
 * @param moduleName Module name
 * @returns Inspection results as a resource response
 */
async function getInspectionResults(
  rubberduck: RubberduckWrapper,
  moduleName: string
): Promise<{ contents: { uri: string; mimeType: string; text: string }[] }> {
  try {
    if (!moduleName) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        'Module name is required'
      );
    }
    
    // Run code analysis
    const analysis = await rubberduck.analyzeCode(moduleName);
    
    // Format the response
    const inspectionResults = {
      moduleName,
      issues: analysis.issues.map(issue => ({
        severity: issue.Severity,
        message: issue.Description,
        line: issue.Line,
        column: issue.Column,
        type: issue.InspectionType,
      })),
      metrics: analysis.metrics,
      count: analysis.issues.length,
      timestamp: new Date().toISOString(),
    };
    
    return {
      contents: [
        {
          uri: `rubberduck://inspection-results/${encodeURIComponent(moduleName)}`,
          mimeType: 'application/json',
          text: JSON.stringify(inspectionResults, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new McpError(
      ErrorCode.InternalError,
      `Error getting inspection results: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
