/**
 * TypeScript interfaces for MCP-specific types used in rubberduck-mcp
 */

/**
 * Arguments for exporting a VBA module
 */
export interface ExportVBAModuleArgs {
  /** Name of the module to export */
  moduleName: string;
  
  /** Whether to include module attributes */
  includeAttributes?: boolean;
}

/**
 * Result of exporting a VBA module
 */
export interface ExportVBAModuleResult {
  /** Module content as text */
  content: string;
  
  /** Module metadata */
  metadata: {
    /** Type of module (Standard, Class, Form, etc.) */
    type: string;
    
    /** Module attributes */
    attributes: Record<string, any>;
  };
}

/**
 * Arguments for importing a VBA module
 */
export interface ImportVBAModuleArgs {
  /** Name of the module to import */
  moduleName: string;
  
  /** Code content to import */
  content: string;
  
  /** Whether to create the module if it doesn't exist */
  createIfNotExists?: boolean;
}

/**
 * Result of importing a VBA module
 */
export interface ImportVBAModuleResult {
  /** Whether the import was successful */
  success: boolean;
  
  /** Any warnings generated during import */
  warnings: string[];
}

/**
 * Arguments for analyzing code
 */
export interface AnalyzeCodeArgs {
  /** Target module name or 'project' for the entire project */
  target: string;
  
  /** Optional rulesets to apply during analysis */
  rulesets?: string[];
}

/**
 * Result of code analysis
 */
export interface AnalyzeCodeResult {
  /** Issues found during analysis */
  issues: Array<{
    /** Issue severity (Error, Warning, Suggestion, Hint) */
    severity: string;
    
    /** Issue message */
    message: string;
    
    /** Issue location */
    location: {
      /** Module where the issue was found */
      module: string;
      
      /** Line number */
      line: number;
      
      /** Column number */
      column: number;
    };
  }>;
  
  /** Code metrics from analysis */
  metrics: Record<string, any>;
}

/**
 * Arguments for executing a refactoring
 */
export interface ExecuteRefactoringArgs {
  /** Type of refactoring to execute */
  refactoringType: string;
  
  /** Target for the refactoring */
  target: {
    /** Module to refactor */
    module: string;
    
    /** Optional selection within the module */
    selection?: {
      /** Start line of selection */
      startLine: number;
      
      /** Start column of selection */
      startColumn: number;
      
      /** End line of selection */
      endLine: number;
      
      /** End column of selection */
      endColumn: number;
    };
  };
  
  /** Optional refactoring options */
  options?: Record<string, any>;
}

/**
 * Result of executing a refactoring
 */
export interface ExecuteRefactoringResult {
  /** Whether the refactoring was successful */
  success: boolean;
  
  /** Result description or code changes */
  result: string;
}

/**
 * Arguments for Git commit
 */
export interface GitCommitArgs {
  /** Commit message */
  message: string;
  
  /** Optional specific files to commit */
  files?: string[];
}

/**
 * Result of Git commit
 */
export interface GitCommitResult {
  /** Commit hash */
  commitHash: string;
  
  /** Summary of changes */
  summary: string;
}

/**
 * Arguments for Git branch operations
 */
export interface GitBranchArgs {
  /** Branch name */
  name: string;
  
  /** Whether to create the branch if it doesn't exist */
  create?: boolean;
  
  /** Whether to checkout the branch */
  checkout?: boolean;
}

/**
 * Result of Git branch operations
 */
export interface GitBranchResult {
  /** Whether the operation was successful */
  success: boolean;
  
  /** Current branch after the operation */
  currentBranch: string;
}

/**
 * Arguments for Git synchronization operations
 */
export interface GitSyncArgs {
  /** Operation to perform (push, pull, fetch) */
  operation: 'push' | 'pull' | 'fetch';
  
  /** Optional remote name */
  remote?: string;
  
  /** Optional branch name */
  branch?: string;
}

/**
 * Result of Git synchronization operations
 */
export interface GitSyncResult {
  /** Whether the operation was successful */
  success: boolean;
  
  /** Details of the operation result */
  details: string;
}

/**
 * Project structure resource
 */
export interface ProjectStructureResource {
  /** Project name */
  name: string;
  
  /** Project description */
  description: string;
  
  /** VBA modules in the project */
  modules: Array<{
    /** Module name */
    name: string;
    
    /** Module type */
    type: string;
    
    /** Module path on disk */
    path?: string;
  }>;
  
  /** Project references */
  references: Array<{
    /** Reference name */
    name: string;
    
    /** Reference version */
    version: string;
    
    /** Reference GUID */
    guid: string;
  }>;
}

/**
 * Modules list resource
 */
export interface ModulesListResource {
  /** Array of modules */
  modules: Array<{
    /** Module name */
    name: string;
    
    /** Module type */
    type: string;
    
    /** Line count */
    lineCount?: number;
    
    /** Last modified timestamp */
    lastModified?: string;
  }>;
}

/**
 * Git status resource
 */
export interface GitStatusResource {
  /** Current branch */
  currentBranch: string;
  
  /** Array of staged changes */
  stagedChanges: Array<{
    /** File path */
    path: string;
    
    /** Change type (Added, Modified, Deleted, etc.) */
    status: string;
  }>;
  
  /** Array of unstaged changes */
  unstagedChanges: Array<{
    /** File path */
    path: string;
    
    /** Change type (Added, Modified, Deleted, etc.) */
    status: string;
  }>;
  
  /** Array of untracked files */
  untrackedFiles: string[];
  
  /** Timestamp of the status check */
  timestamp: string;
}

/**
 * Code history resource
 */
export interface CodeHistoryResource {
  /** Module name */
  moduleName: string;
  
  /** Array of commits */
  commits: Array<{
    /** Commit hash */
    hash: string;
    
    /** Commit message */
    message: string;
    
    /** Author name */
    author: string;
    
    /** Commit date */
    date: string;
  }>;
}

/**
 * Inspection results resource
 */
export interface InspectionResultsResource {
  /** Module name */
  moduleName: string;
  
  /** Timestamp of the analysis */
  timestamp: string;
  
  /** Array of issues */
  issues: Array<{
    /** Issue severity */
    severity: string;
    
    /** Issue message */
    message: string;
    
    /** Line number */
    line: number;
    
    /** Column number */
    column: number;
    
    /** Inspection type */
    inspectionType: string;
  }>;
}

/**
 * Refactoring options resource
 */
export interface RefactoringOptionsResource {
  /** Available refactorings */
  refactorings: Array<{
    /** Refactoring name */
    name: string;
    
    /** Refactoring description */
    description: string;
    
    /** Whether this refactoring requires a selection */
    requiresSelection: boolean;
    
    /** Available options for this refactoring */
    options?: Array<{
      /** Option name */
      name: string;
      
      /** Option description */
      description: string;
      
      /** Option type */
      type: 'string' | 'boolean' | 'number';
      
      /** Default value */
      defaultValue?: any;
    }>;
  }>;
}
