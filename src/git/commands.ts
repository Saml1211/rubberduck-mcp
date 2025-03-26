/**
 * Git command integration for Rubberduck
 */

import { RubberduckWrapper } from '../com/rubberduck.js';
import { 
  IGitStatus, 
  IGitCommit, 
  IGitFileStatus 
} from '../types/rubberduck.js';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';

/**
 * Configuration options for Git operations
 */
export interface GitCommandOptions {
  /** API key for authentication */
  apiKey?: string;
  
  /** Debug mode for additional logging */
  debug?: boolean;
}

/**
 * Class for managing Git operations through Rubberduck
 */
export class GitCommands {
  private rubberduck: RubberduckWrapper;
  private readonly options: GitCommandOptions;
  
  /**
   * Creates a new GitCommands instance
   * @param rubberduck Rubberduck wrapper instance
   * @param options Configuration options
   */
  constructor(rubberduck: RubberduckWrapper, options: GitCommandOptions = {}) {
    this.rubberduck = rubberduck;
    this.options = options;
  }
  
  /**
   * Validates authentication if API key is configured
   * @param apiKey The API key to validate
   * @throws {McpError} If authentication fails
   */
  public validateAuth(apiKey?: string): void {
    if (this.options.apiKey && apiKey !== this.options.apiKey) {
      throw new McpError(
        ErrorCode.Unauthorized,
        'Invalid API key'
      );
    }
  }
  
  /**
   * Gets the current Git status
   * @returns Git status information
   * @throws {McpError} If the status cannot be retrieved
   */
  public async getStatus(): Promise<IGitStatus> {
    return await this.rubberduck.gitStatus();
  }
  
  /**
   * Commits changes to the repository
   * @param message Commit message
   * @param files Optional specific files to commit
   * @returns Commit hash and summary
   * @throws {McpError} If the commit fails
   */
  public async commit(
    message: string,
    files?: string[]
  ): Promise<{ commitHash: string; summary: string }> {
    if (!message || message.trim() === '') {
      throw new McpError(
        ErrorCode.InvalidParams,
        'Commit message cannot be empty'
      );
    }
    
    return await this.rubberduck.gitCommit(message, files);
  }
  
  /**
   * Creates and/or checks out a Git branch
   * @param name Branch name
   * @param create Whether to create the branch if it doesn't exist
   * @param checkout Whether to checkout the branch
   * @returns Success status and current branch
   * @throws {McpError} If the branch operation fails
   */
  public async branch(
    name: string,
    create = false,
    checkout = true
  ): Promise<{ success: boolean; currentBranch: string }> {
    if (!name || name.trim() === '') {
      throw new McpError(
        ErrorCode.InvalidParams,
        'Branch name cannot be empty'
      );
    }
    
    return await this.rubberduck.gitBranch(name, create, checkout);
  }
  
  /**
   * Performs Git synchronization operations (push, pull, fetch)
   * @param operation Operation to perform
   * @param remote Optional remote name
   * @param branch Optional branch name
   * @returns Success status and details
   * @throws {McpError} If the sync operation fails
   */
  public async sync(
    operation: 'push' | 'pull' | 'fetch',
    remote?: string,
    branch?: string
  ): Promise<{ success: boolean; details: string }> {
    if (!['push', 'pull', 'fetch'].includes(operation)) {
      throw new McpError(
        ErrorCode.InvalidParams,
        `Invalid Git sync operation: ${operation}`
      );
    }
    
    return await this.rubberduck.gitSync(operation, remote, branch);
  }
  
  /**
   * Gets commit history for a specific module
   * @param moduleName Module name
   * @returns Array of commits
   * @throws {McpError} If the history cannot be retrieved
   */
  public async getModuleHistory(moduleName: string): Promise<IGitCommit[]> {
    if (!moduleName || moduleName.trim() === '') {
      throw new McpError(
        ErrorCode.InvalidParams,
        'Module name cannot be empty'
      );
    }
    
    return await this.rubberduck.getModuleHistory(moduleName);
  }
  
  /**
   * Formats Git status into a human-readable string
   * @param status Git status to format
   * @returns Formatted status string
   */
  public formatStatus(status: IGitStatus): string {
    let result = `Branch: ${status.CurrentBranch}\n\n`;
    
    if (status.StagedChanges.length > 0) {
      result += `Changes staged for commit:\n`;
      status.StagedChanges.forEach(change => {
        result += `  ${change.Status} ${change.Path}\n`;
      });
      result += '\n';
    } else {
      result += 'No staged changes\n\n';
    }
    
    if (status.UnstagedChanges.length > 0) {
      result += `Changes not staged for commit:\n`;
      status.UnstagedChanges.forEach(change => {
        result += `  ${change.Status} ${change.Path}\n`;
      });
      result += '\n';
    } else {
      result += 'No unstaged changes\n';
    }
    
    return result;
  }
  
  /**
   * Formats commit history into a human-readable string
   * @param commits Commit history to format
   * @returns Formatted history string
   */
  public formatHistory(commits: IGitCommit[]): string {
    if (commits.length === 0) {
      return 'No commit history found';
    }
    
    let result = '';
    
    commits.forEach(commit => {
      result += `${commit.Hash}\n`;
      result += `Author: ${commit.Author}\n`;
      result += `Date: ${commit.Date}\n`;
      result += `\n    ${commit.Message.split('\n').join('\n    ')}\n\n`;
    });
    
    return result;
  }
}
