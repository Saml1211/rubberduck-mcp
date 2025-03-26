// Note: node-win32com package needs to be installed
// npm install node-win32com
import { win32com } from 'node-win32com';
import { 
  IRubberduckApp, 
  ISourceControlManager, 
  ICodeAnalyzer,
  ICodeIssue,
  IRefactoringEngine,
  IRefactoringResult,
  ICodeSelection,
  IModuleInfo,
  IGitStatus,
  ICodeMetrics,
  IGitCommit,
  ModuleType
} from '../types/rubberduck.js';
import { 
  releaseComObject, 
  retryComOperation, 
  withTimeout,
  usingComObject
} from './utils.js';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';

/**
 * Configuration options for the Rubberduck wrapper
 */
export interface RubberduckWrapperOptions {
  /** API key for authentication */
  apiKey?: string;
  
  /** Timeout in milliseconds for COM operations */
  comTimeout?: number;
  
  /** Maximum retries for COM operations */
  maxRetries?: number;
  
  /** Whether to run in debug mode with additional logging */
  debug?: boolean;
}

/**
 * A wrapper around the Rubberduck COM objects with proper resource management
 */
export class RubberduckWrapper {
  private comApp: any | null = null;
  private isConnected = false;
  private readonly options: Required<RubberduckWrapperOptions>;
  
  /**
   * Default options for the Rubberduck wrapper
   */
  private static readonly DEFAULT_OPTIONS: Required<RubberduckWrapperOptions> = {
    apiKey: '',
    comTimeout: 30000, // 30 seconds
    maxRetries: 3,
    debug: false
  };

  /**
   * Creates a new instance of the Rubberduck COM wrapper
   * @param options Configuration options
   * @throws {McpError} If the Rubberduck COM object cannot be created
   */
  constructor(options: RubberduckWrapperOptions = {}) {
    this.options = { ...RubberduckWrapper.DEFAULT_OPTIONS, ...options };
    
    if (this.options.debug) {
      console.log('Initializing RubberduckWrapper with options:', this.options);
    }
  }

  /**
   * Initializes the COM connection to Rubberduck
   * @throws {McpError} If the Rubberduck COM object cannot be created
   */
  private async initializeCom(): Promise<void> {
    if (this.comApp) {
      return;
    }
    
    try {
      // Create the Rubberduck COM application object
      this.comApp = win32com.createObject('Rubberduck.Application');
      
      if (!this.comApp) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to create Rubberduck COM object. Ensure Rubberduck is installed.'
        );
      }
      
      if (this.options.debug) {
        console.log('Successfully created Rubberduck COM object');
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
      await this.initializeCom();
      
      // Wrap the connection attempt with retry and timeout
      this.isConnected = await retryComOperation(
        async () => {
          return await withTimeout(
            Promise.resolve(Boolean(this.comApp.Connect())),
            this.options.comTimeout
          );
        },
        this.options.maxRetries
      );
      
      if (this.isConnected) {
        if (this.options.debug) {
          console.log('Successfully connected to Rubberduck');
          console.log('Rubberduck version:', await this.getVersion());
        }
      } else {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to connect to Rubberduck'
        );
      }
      
      return this.isConnected;
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
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
  /**
   * Ensures the COM object is connected before performing operations
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
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to get Rubberduck version: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Validates the API key if one is configured
   * @throws {McpError} If the API key is invalid
   */
  public validateApiKey(apiKey: string): void {
    if (!this.options.apiKey) {
      return;
    }
    
    if (apiKey !== this.options.apiKey) {
      throw new McpError(
        ErrorCode.Unauthorized,
        'Invalid API key'
      );
    }
  }
  
  // #region Module Management
  
  /**
   * Exports a VBA module to text
   * @param moduleName The name of the module to export
   * @param includeAttributes Whether to include module attributes
   * @returns The module content as text and its metadata
   * @throws {McpError} If the module cannot be exported
   */
  public async exportVBAModule(
    moduleName: string,
    includeAttributes = false
  ): Promise<{ content: string; metadata: { type: string; attributes: Record<string, any> } }> {
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
      
      // Get the module metadata first to determine its type
      const modules = await this.getModulesList();
      const moduleInfo = modules.find(m => m.Name === moduleName);
      
      if (!moduleInfo) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      // Convert module type to string representation
      const moduleType = ModuleType[moduleInfo.Type] || 'Unknown';
      
      // Extract module attributes if requested
      const attributes: Record<string, any> = {};
      if (includeAttributes) {
        try {
          // This is a placeholder for attribute extraction
          // Actual implementation would depend on Rubberduck's API for accessing module attributes
          // For example, if Rubberduck provides a GetAttributes method:
          // const attributesObject = sourceControl.GetAttributes(moduleName);
          // attributes = attributesObject ? JSON.parse(attributesObject) : {};
        } catch (error) {
          console.error(`Error extracting attributes for module ${moduleName}:`, error);
          // Continue without attributes rather than failing
        }
      }
      
      // Export the module content
      const moduleContent = await retryComOperation(
        async () => {
          return await withTimeout(
            Promise.resolve(sourceControl.ExportModule(moduleName, includeAttributes)),
            this.options.comTimeout
          );
        },
        this.options.maxRetries
      );
      
      if (moduleContent === null || moduleContent === undefined) {
        throw new McpError(
          ErrorCode.NotFound,
          `Failed to export module: ${moduleName}`
        );
      }
      
      return {
        content: String(moduleContent),
        metadata: {
          type: moduleType,
          attributes
        }
      };
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
   * @param createIfNotExists Whether to create the module if it doesn't exist
   * @returns Success status and any warnings
   * @throws {McpError} If the code cannot be imported
   */
  public async importVBAModule(
    moduleName: string,
    code: string,
    createIfNotExists = false
  ): Promise<{ success: boolean; warnings: string[] }> {
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
      
      // Check if the module exists
      const modules = await this.getModulesList();
      const moduleExists = modules.some(m => m.Name === moduleName);
      
      if (!moduleExists && !createIfNotExists) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}. Set createIfNotExists to true to create it.`
        );
      }
      
      // Handle module creation if needed and supported
      if (!moduleExists && createIfNotExists) {
        // This is a placeholder for module creation
        // Actual implementation would depend on Rubberduck's API
        // For example:
        // sourceControl.CreateModule(moduleName, ModuleType.Standard);
        console.log(`Creating new module: ${moduleName}`);
      }
      
      // Import the code
      const result = await retryComOperation(
        async () => {
          return await withTimeout(
            Promise.resolve(sourceControl.ImportModule(moduleName, code, createIfNotExists)),
            this.options.comTimeout
          );
        },
        this.options.maxRetries
      );
      
      // Collect any warnings (placeholder - actual implementation would depend on API)
      const warnings: string[] = [];
      // Example if Rubberduck provides a GetLastWarnings method:
      // const warningsObject = sourceControl.GetLastWarnings();
      // if (warningsObject && Array.isArray(warningsObject)) {
      //   warnings = warningsObject.map(w => String(w));
      // }
      
      return {
        success: Boolean(result),
        warnings
      };
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
   * Gets a list of all modules in the current project
   * @returns Array of module info objects
   * @throws {McpError} If the list cannot be retrieved
   */
  public async getModulesList(): Promise<IModuleInfo[]> {
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
      
      const modules = await retryComOperation(
        async () => {
          return await withTimeout(
            Promise.resolve(sourceControl.GetModulesList()),
            this.options.comTimeout
          );
        },
        this.options.maxRetries
      );
      
      if (!modules) {
        return [];
      }
      
      // Convert COM objects to plain JavaScript objects
      return Array.isArray(modules) 
        ? modules.map((module: any) => ({
            Name: String(module.Name),
            Type: Number(module.Type),
            Path: module.Path ? String(module.Path) : undefined
          }))
        : [];
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
  
  // #endregion
  
  // #region Code Analysis
  
  /**
   * Analyzes a module or the entire project for code issues
   * @param target The target module name or "project" for the entire project
   * @param rulesets Optional array of ruleset names to apply
   * @returns Array of code issues and metrics
   * @throws {McpError} If analysis fails
   */
  public async analyzeCode(
    target: string,
    rulesets?: string[]
  ): Promise<{ issues: ICodeIssue[]; metrics: Record<string, any> }> {
    try {
      this.ensureConnected();
      
      const analyzer = this.comApp.CodeAnalysis;
      if (!analyzer) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access Rubberduck Code Analysis'
        );
      }
      
      let issues: ICodeIssue[] = [];
      let metrics: Record<string, any> = {};
      
      // Analyze either a specific module or the entire project
      if (target.toLowerCase() === 'project') {
        issues = await retryComOperation(
          async () => {
            return await withTimeout(
              Promise.resolve(analyzer.AnalyzeProject(rulesets)),
              this.options.comTimeout * 2 // Double timeout for project analysis as it may take longer
            );
          },
          this.options.maxRetries
        );
        
        // Project-level metrics are not implemented in this example
        metrics = {
          totalModules: 0, // Would be populated from actual project data
          totalIssues: Array.isArray(issues) ? issues.length : 0
        };
      } else {
        // Verify the module exists
        const modules = await this.getModulesList();
        const moduleExists = modules.some(m => m.Name === target);
        
        if (!moduleExists) {
          throw new McpError(
            ErrorCode.NotFound,
            `Module not found: ${target}`
          );
        }
        
        // Analyze the specific module
        issues = await retryComOperation(
          async () => {
            return await withTimeout(
              Promise.resolve(analyzer.AnalyzeModule(target, rulesets)),
              this.options.comTimeout
            );
          },
          this.options.maxRetries
        );
        
        // Get module metrics if available
        try {
          const moduleMetrics: ICodeMetrics = await retryComOperation(
            async () => {
              return await withTimeout(
                Promise.resolve(analyzer.GetModuleMetrics(target)),
                this.options.comTimeout
              );
            },
            this.options.maxRetries
          );
          
          if (moduleMetrics) {
            metrics = {
              linesOfCode: moduleMetrics.LinesOfCode,
              cyclomaticComplexity: moduleMetrics.CyclomaticComplexity,
              methodCount: moduleMetrics.MethodCount,
              ...moduleMetrics.Metrics
            };
          }
        } catch (error) {
          console.error(`Error getting metrics for module ${target}:`, error);
          // Continue without metrics rather than failing
          metrics = {
            error: 'Failed to retrieve metrics'
          };
        }
      }
      
      // Convert COM objects to plain JavaScript objects
      return {
        issues: Array.isArray(issues) 
          ? issues.map((issue: any) => ({
              Severity: String(issue.Severity) as "Hint" | "Suggestion" | "Warning" | "Error",
              Description: String(issue.Description),
              ModuleName: String(issue.ModuleName),
              Line: Number(issue.Line),
              Column: Number(issue.Column),
              InspectionType: issue.InspectionType ? String(issue.InspectionType) : 'Unknown',
              QuickFix: issue.QuickFix ? String(issue.QuickFix) : undefined
            }))
          : [],
        metrics
      };
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error analyzing code for ${target}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }
  
  // #endregion
  
  // #region Refactoring
  
  /**
   * Executes a refactoring on a module
   * @param refactoringType The type of refactoring to execute
   * @param moduleName The module to refactor
   * @param selection Optional selection within the module
   * @param options Optional refactoring options
   * @returns Refactoring result
   * @throws {McpError} If the refactoring fails
   */
  public async executeRefactoring(
    refactoringType: string,
    moduleName: string,
    selection?: ICodeSelection,
    options?: Record<string, any>
  ): Promise<IRefactoringResult> {
    try {
      this.ensureConnected();
      
      const refactoringEngine = this.comApp.Refactorings;
      if (!refactoringEngine) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access Rubberduck Refactorings'
        );
      }
      
      // Verify the module exists
      const modules = await this.getModulesList();
      const moduleExists = modules.some(m => m.Name === moduleName);
      
      if (!moduleExists) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      // Verify the refactoring type is available
      const availableRefactorings = refactoringEngine.AvailableRefactoringTypes;
      if (availableRefactorings && !availableRefactorings.includes(refactoringType)) {
        throw new McpError(
          ErrorCode.InvalidRequest,
          `Refactoring type not available: ${refactoringType}`
        );
      }
      
      // Execute the refactoring
      const result = await retryComOperation(
        async () => {
          return await withTimeout(
            Promise.resolve(refactoringEngine.ExecuteRefactoring(
              refactoringType,
              moduleName,
              selection,
              options
            )),
            this.options.comTimeout
          );
        },
        this.options.maxRetries
      );
      
      // Convert COM object to plain JavaScript object
      return {
        Success: Boolean(result.Success),
        Description: String(result.Description),
        ErrorMessage: result.ErrorMessage ? String(result.ErrorMessage) : undefined
      };
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error executing refactoring ${refactoringType} on ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }
  
  /**
   * Gets a list of available refactoring types
   * @returns Array of refactoring type names
   * @throws {McpError} If the list cannot be retrieved
   */
  public async getAvailableRefactoringTypes(): Promise<string[]> {
    try {
      this.ensureConnected();
      
      const refactoringEngine = this.comApp.Refactorings;
      if (!refactoringEngine) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access Rubberduck Refactorings'
        );
      }
      
      const types = refactoringEngine.AvailableRefactoringTypes;
      return types ? types.map((t: any) => String(t)) : [];
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error getting available refactoring types: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }
  
  // #endregion
  
  // #region Git Integration
  
  /**
   * Commits changes to the repository
   * @param message The commit message
   * @param files Optional array of specific files to commit
   * @returns The commit hash and a summary
   * @throws {McpError} If the commit fails
   */
  public async gitCommit(
    message: string,
    files?: string[]
  ): Promise<{ commitHash: string; summary: string }> {
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
      
      // Get current status to include in summary
      const beforeStatus = await this.gitStatus();
      
      // Commit changes
      const commitHash = await retryComOperation(
        async () => {
          return await withTimeout(
            Promise.resolve(sourceControl.Commit(message, files)),
            this.options.comTimeout
          );
        },
        this.options.maxRetries
      );
      
      if (!commitHash) {
        throw new McpError(
          ErrorCode.InternalError,
          'Commit operation failed or returned empty hash'
        );
      }
      
      // Get updated status for summary
      const afterStatus = await this.gitStatus();
      
      // Create summary of what was committed
      const summary = this.createCommitSummary(beforeStatus, afterStatus, files);
      
      return {
        commitHash: String(commitHash),
        summary
      };
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
   * Creates a summary of committed changes by comparing before and after status
   * @param before Status before commit
   * @param after Status after commit
   * @param files Optional array of specific files that were committed
   * @returns A descriptive summary string
   */
  private createCommitSummary(
    before: IGitStatus,
    after: IGitStatus,
    files?: string[]
  ): string {
    // Count of staged files that were committed
    const committedCount = before.StagedChanges.length;
    
    // If specific files were provided, list them
    if (files && files.length > 0) {
      return `Committed ${files.length} specified files: ${files.join(', ')}`;
    }
    
    // Otherwise, calculate what was committed from before/after status
    return `Committed ${committedCount} changes to branch '${after.CurrentBranch}'`;
  }
  
  /**
   * Gets the current Git status
   * @returns Git status information
   * @throws {McpError} If the status cannot be retrieved
   */
  public async gitStatus(): Promise<IGitStatus> {
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
      
      const status = await retryComOperation(
        async () => {
          return await withTimeout(
            Promise.resolve(sourceControl.GetStatus()),
            this.options.comTimeout
          );
        },
        this.options.maxRetries
      );
      
      // Convert COM object to plain JavaScript object
      return {
        CurrentBranch: String(status.CurrentBranch),
        HasChanges: Boolean(status.HasChanges),
        StagedChanges: Array.isArray(status.StagedChanges) 
          ? status.StagedChanges.map((change: any) => ({
              Path: String(change.Path),
              Status: String(change.Status)
            }))
          : [],
        UnstagedChanges: Array.isArray(status.UnstagedChanges)
          ? status.UnstagedChanges.map((change: any) => ({
              Path: String(change.Path),
              Status: String(change.Status)
            }))
          : []
      };
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error getting Git status: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Creates and/or checks out a Git branch
   * @param branchName The name of the branch
   * @param create Whether to create the branch if it doesn't exist
   * @param checkout Whether to checkout the branch
   * @returns Success status and current branch
   * @throws {McpError} If the branch operation fails
   */
  public async gitBranch(
    branchName: string,
    create = false,
    checkout = true
  ): Promise<{ success: boolean; currentBranch: string }> {
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
      
      let success = true;
      
      // Create branch if requested
      if (create) {
        success = await retryComOperation(
          async () => {
            return await withTimeout(
              Promise.resolve(sourceControl.CreateBranch(branchName)),
              this.options.comTimeout
            );
          },
          this.options.maxRetries
        );
        
        if (!success) {
          throw new McpError(
            ErrorCode.InternalError,
            `Failed to create branch: ${branchName}`
          );
        }
      }
      
      // Checkout branch if requested
      if (checkout) {
        success = await retryComOperation(
          async () => {
            return await withTimeout(
              Promise.resolve(sourceControl.CheckoutBranch(branchName)),
              this.options.comTimeout
            );
          },
          this.options.maxRetries
        );
        
        if (!success) {
          throw new McpError(
            ErrorCode.InternalError,
            `Failed to checkout branch: ${branchName}`
          );
        }
      }
      
      // Get current branch
      const currentBranch = String(sourceControl.CurrentBranch);
      
      return {
        success,
        currentBranch
      };
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error in Git branch operation for ${branchName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Performs Git synchronization operations (push, pull, fetch)
   * @param operation The operation to perform
   * @param remote Optional remote name
   * @param branch Optional branch name
   * @returns Success status and details
   * @throws {McpError} If the sync operation fails
   */
  public async gitSync(
    operation: 'push' | 'pull' | 'fetch',
    remote?: string,
    branch?: string
  ): Promise<{ success: boolean; details: string }> {
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
      
      let success = false;
      
      // Perform the requested operation
      switch (operation) {
        case 'push':
          success = await retryComOperation(
            async () => {
              return await withTimeout(
                Promise.resolve(sourceControl.Push(remote, branch)),
                this.options.comTimeout * 2 // Double timeout for network operations
              );
            },
            this.options.maxRetries
          );
          break;
          
        case 'pull':
          success = await retryComOperation(
            async () => {
              return await withTimeout(
                Promise.resolve(sourceControl.Pull(remote, branch)),
                this.options.comTimeout * 2
              );
            },
            this.options.maxRetries
          );
          break;
          
        case 'fetch':
          success = await retryComOperation(
            async () => {
              return await withTimeout(
                Promise.resolve(sourceControl.Fetch(remote)),
                this.options.comTimeout * 2
              );
            },
            this.options.maxRetries
          );
          break;
          
        default:
          throw new McpError(
            ErrorCode.InvalidRequest,
            `Invalid Git sync operation: ${operation}`
          );
      }
      
      // Create details string based on operation
      let details = '';
      if (success) {
        details = `Successfully performed ${operation} operation`;
        if (remote) {
          details += ` with remote '${remote}'`;
        }
        if (branch && (operation === 'push' || operation === 'pull')) {
          details += ` on branch '${branch}'`;
        }
      } else {
        details = `Failed to perform ${operation} operation`;
      }
      
      return {
        success,
        details
      };
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error in Git sync operation ${operation}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Gets commit history for a specific module
   * @param moduleName The name of the module
   * @returns Array of commits
   * @throws {McpError} If the history cannot be retrieved
   */
  public async getModuleHistory(moduleName: string): Promise<IGitCommit[]> {
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
      
      // Verify the module exists
      const modules = await this.getModulesList();
      const moduleExists = modules.some(m => m.Name === moduleName);
      
      if (!moduleExists) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      const history = await retryComOperation(
        async () => {
          return await withTimeout(
            Promise.resolve(sourceControl.GetModuleHistory(moduleName)),
            this.options.comTimeout
          );
        },
        this.options.maxRetries
      );
      
      // Convert COM objects to plain JavaScript objects
      return Array.isArray(history) 
        ? history.map((commit: any) => ({
            Hash: String(commit.Hash),
            Message: String(commit.Message),
            Author: String(commit.Author),
            Date: String(commit.Date)
          }))
        : [];
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Error getting module history for ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }
  
  // #endregion

  /**
   * Releases COM resources
   * This should be called when the wrapper is no longer needed
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
