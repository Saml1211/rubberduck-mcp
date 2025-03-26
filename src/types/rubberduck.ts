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
  
  /** Gets the refactoring engine */
  readonly Refactorings: IRefactoringEngine;
  
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
  Commit(message: string, files?: string[]): string;
  
  /** Creates a new branch */
  CreateBranch(branchName: string): boolean;
  
  /** Switches to an existing branch */
  CheckoutBranch(branchName: string): boolean;
  
  /** Pulls changes from the remote repository */
  Pull(remote?: string, branch?: string): boolean;
  
  /** Pushes changes to the remote repository */
  Push(remote?: string, branch?: string): boolean;
  
  /** Fetches from the remote repository */
  Fetch(remote?: string): boolean;
  
  /** Exports a VBA module to text */
  ExportModule(moduleName: string, includeAttributes?: boolean): string;
  
  /** Imports code into a VBA module */
  ImportModule(moduleName: string, code: string, createIfNotExists?: boolean): boolean;
  
  /** Gets the list of modules in the current VBA project */
  GetModulesList(): IModuleInfo[];
  
  /** Gets the Git status of the repository */
  GetStatus(): IGitStatus;
  
  /** Gets the commit history for a specific module */
  GetModuleHistory(moduleName: string): IGitCommit[];
}

/**
 * Represents a module in the VBA project
 */
export interface IModuleInfo {
  /** The name of the module */
  readonly Name: string;
  
  /** The type of the module */
  readonly Type: ModuleType;
  
  /** The path to the module on disk */
  readonly Path?: string;
}

/**
 * Enumeration of module types
 */
export enum ModuleType {
  Standard = 1,
  ClassModule = 2,
  Form = 3,
  Document = 4,
  ThisWorkbook = 100
}

/**
 * Represents the Git status of a repository
 */
export interface IGitStatus {
  /** List of staged changes */
  readonly StagedChanges: IGitFileStatus[];
  
  /** List of unstaged changes */
  readonly UnstagedChanges: IGitFileStatus[];
  
  /** Current branch name */
  readonly CurrentBranch: string;
  
  /** Whether there are uncommitted changes */
  readonly HasChanges: boolean;
}

/**
 * Represents the status of a single file in Git
 */
export interface IGitFileStatus {
  /** The path of the file */
  readonly Path: string;
  
  /** The status code (e.g., Modified, Added, Deleted) */
  readonly Status: string;
}

/**
 * Represents a Git commit
 */
export interface IGitCommit {
  /** The commit hash */
  readonly Hash: string;
  
  /** The commit message */
  readonly Message: string;
  
  /** The author name */
  readonly Author: string;
  
  /** The commit date */
  readonly Date: string;
}

/**
 * Represents Rubberduck's code analysis engine
 */
export interface ICodeAnalyzer {
  /** Gets all available inspection types */
  readonly AvailableInspectionTypes: string[];
  
  /** Runs code analysis on the specified module */
  AnalyzeModule(moduleName: string, inspectionTypes?: string[]): ICodeIssue[];
  
  /** Runs code analysis on the entire project */
  AnalyzeProject(inspectionTypes?: string[]): ICodeIssue[];
  
  /** Gets the metrics for a module */
  GetModuleMetrics(moduleName: string): ICodeMetrics;
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
  
  /** The inspection type that generated this issue */
  readonly InspectionType: string;
  
  /** Suggested fix, if available */
  readonly QuickFix?: string;
}

/**
 * Represents code metrics for a module
 */
export interface ICodeMetrics {
  /** Lines of code */
  readonly LinesOfCode: number;
  
  /** Cyclomatic complexity */
  readonly CyclomaticComplexity: number;
  
  /** Number of methods */
  readonly MethodCount: number;
  
  /** Additional metrics as key-value pairs */
  readonly Metrics: Record<string, number>;
}

/**
 * Represents Rubberduck's refactoring engine
 */
export interface IRefactoringEngine {
  /** Gets all available refactoring types */
  readonly AvailableRefactoringTypes: string[];
  
  /** Executes a refactoring */
  ExecuteRefactoring(
    refactoringType: string, 
    moduleName: string, 
    selection?: ICodeSelection,
    options?: Record<string, any>
  ): IRefactoringResult;
}

/**
 * Represents a selection in the code
 */
export interface ICodeSelection {
  /** Start line of the selection */
  readonly StartLine: number;
  
  /** Start column of the selection */
  readonly StartColumn: number;
  
  /** End line of the selection */
  readonly EndLine: number;
  
  /** End column of the selection */
  readonly EndColumn: number;
}

/**
 * Represents the result of a refactoring operation
 */
export interface IRefactoringResult {
  /** Whether the refactoring was successful */
  readonly Success: boolean;
  
  /** Description of the applied refactoring */
  readonly Description: string;
  
  /** Error message if the refactoring failed */
  readonly ErrorMessage?: string;
}
