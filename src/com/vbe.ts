import { win32com } from 'node-win32com';
import { 
  IVBE,
  IVBProject,
  IVBComponents,
  IVBComponent,
  ICodeModule,
  vbext_ComponentType
} from '../types/vbe.js';
import { 
  releaseComObject, 
  retryComOperation, 
  withTimeout,
  usingComObject
} from './utils.js';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';

/**
 * Configuration options for the VBE wrapper
 */
export interface VBEWrapperOptions {
  /** Timeout in milliseconds for COM operations */
  comTimeout?: number;
  
  /** Maximum retries for COM operations */
  maxRetries?: number;
  
  /** Whether to run in debug mode with additional logging */
  debug?: boolean;
}

/**
 * A wrapper around the VBE (Visual Basic Editor) COM objects with proper resource management
 */
export class VBEWrapper {
  private comVBE: any | null = null;
  private isConnected = false;
  private readonly options: Required<VBEWrapperOptions>;
  
  /**
   * Default options for the VBE wrapper
   */
  private static readonly DEFAULT_OPTIONS: Required<VBEWrapperOptions> = {
    comTimeout: 20000, // 20 seconds
    maxRetries: 2,
    debug: false
  };

  /**
   * Creates a new instance of the VBE COM wrapper
   * @param options Configuration options
   */
  constructor(options: VBEWrapperOptions = {}) {
    this.options = { ...VBEWrapper.DEFAULT_OPTIONS, ...options };
    
    if (this.options.debug) {
      console.log('Initializing VBEWrapper with options:', this.options);
    }
  }

  /**
   * Connects to the VBE application
   * @returns True if successfully connected
   * @throws {McpError} If connection fails
   */
  public async connect(): Promise<boolean> {
    try {
      // Try to get the running VBE instance
      this.comVBE = win32com.getObject('VBE');
      
      if (!this.comVBE) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to connect to VBE. Ensure VBA host application is running.'
        );
      }
      
      this.isConnected = true;
      
      if (this.options.debug) {
        console.log('Successfully connected to VBE');
        console.log('VBE version:', await this.getVersion());
      }
      
      return true;
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to connect to VBE: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Ensures the COM object is connected before performing operations
   * @throws {McpError} If not connected
   */
  private ensureConnected(): void {
    if (!this.comVBE) {
      throw new McpError(
        ErrorCode.InternalError,
        'VBE COM object is not initialized'
      );
    }
    
    if (!this.isConnected) {
      throw new McpError(
        ErrorCode.InternalError,
        'Not connected to VBE. Call connect() first.'
      );
    }
  }

  /**
   * Gets the VBE version
   * @returns The version string
   * @throws {McpError} If the version cannot be retrieved
   */
  public async getVersion(): Promise<string> {
    try {
      this.ensureConnected();
      return String(this.comVBE.Version);
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to get VBE version: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Gets the active VBE project
   * @returns The active project
   * @throws {McpError} If the active project cannot be retrieved
   */
  public async getActiveProject(): Promise<{
    name: string;
    fileName: string;
    description: string;
    componentCount: number;
  }> {
    try {
      this.ensureConnected();
      
      const activeProject = await retryComOperation(
        async () => {
          return await withTimeout(
            Promise.resolve(this.comVBE.ActiveVBProject),
            this.options.comTimeout
          );
        },
        this.options.maxRetries
      );
      
      if (!activeProject) {
        throw new McpError(
          ErrorCode.NotFound,
          'No active VB project found'
        );
      }
      
      // Extract information from the project
      const name = String(activeProject.Name);
      const fileName = String(activeProject.FileName);
      const description = activeProject.Description ? String(activeProject.Description) : '';
      const components = activeProject.VBComponents;
      const componentCount = components ? Number(components.Count) : 0;
      
      return {
        name,
        fileName,
        description,
        componentCount
      };
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to get active project: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Gets a list of all components (modules) in the active project
   * @returns Array of component info objects
   * @throws {McpError} If the list cannot be retrieved
   */
  public async getComponents(): Promise<{
    name: string;
    type: vbext_ComponentType;
    lineCount: number;
  }[]> {
    try {
      this.ensureConnected();
      
      const activeProject = this.comVBE.ActiveVBProject;
      if (!activeProject) {
        throw new McpError(
          ErrorCode.NotFound,
          'No active VB project found'
        );
      }
      
      const components = activeProject.VBComponents;
      if (!components) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access project components'
        );
      }
      
      const count = Number(components.Count);
      const result: {
        name: string;
        type: vbext_ComponentType;
        lineCount: number;
      }[] = [];
      
      // Loop through all components
      for (let i = 1; i <= count; i++) {
        const component = components.Item(i);
        
        if (component) {
          const name = String(component.Name);
          const type = Number(component.Type) as vbext_ComponentType;
          
          let lineCount = 0;
          if (component.CodeModule) {
            lineCount = Number(component.CodeModule.CountOfLines);
          }
          
          result.push({ name, type, lineCount });
          
          // Release the component COM object
          releaseComObject(component);
        }
      }
      
      return result;
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to get components: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Gets code from a specific component (module)
   * @param moduleName The name of the module
   * @returns The module code
   * @throws {McpError} If the module cannot be found or the code cannot be retrieved
   */
  public async getModuleCode(moduleName: string): Promise<string> {
    try {
      this.ensureConnected();
      
      const activeProject = this.comVBE.ActiveVBProject;
      if (!activeProject) {
        throw new McpError(
          ErrorCode.NotFound,
          'No active VB project found'
        );
      }
      
      const components = activeProject.VBComponents;
      if (!components) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access project components'
        );
      }
      
      let component;
      try {
        component = components.Item(moduleName);
      } catch (error) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      if (!component) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      let code = '';
      
      try {
        const codeModule = component.CodeModule;
        if (!codeModule) {
          throw new McpError(
            ErrorCode.InternalError,
            `Failed to access code module for ${moduleName}`
          );
        }
        
        const lineCount = Number(codeModule.CountOfLines);
        
        if (lineCount > 0) {
          code = String(codeModule.Lines(1, lineCount));
        }
      } finally {
        // Release the component COM object
        releaseComObject(component);
      }
      
      return code;
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to get module code for ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Updates code in a specific component (module)
   * @param moduleName The name of the module
   * @param code The code to update
   * @returns Success status
   * @throws {McpError} If the module cannot be found or the code cannot be updated
   */
  public async updateModuleCode(moduleName: string, code: string): Promise<boolean> {
    try {
      this.ensureConnected();
      
      const activeProject = this.comVBE.ActiveVBProject;
      if (!activeProject) {
        throw new McpError(
          ErrorCode.NotFound,
          'No active VB project found'
        );
      }
      
      const components = activeProject.VBComponents;
      if (!components) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access project components'
        );
      }
      
      let component;
      try {
        component = components.Item(moduleName);
      } catch (error) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      if (!component) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      let success = false;
      
      try {
        const codeModule = component.CodeModule;
        if (!codeModule) {
          throw new McpError(
            ErrorCode.InternalError,
            `Failed to access code module for ${moduleName}`
          );
        }
        
        // Clear existing code
        const lineCount = Number(codeModule.CountOfLines);
        if (lineCount > 0) {
          codeModule.DeleteLines(1, lineCount);
        }
        
        // Add new code
        if (code.length > 0) {
          codeModule.AddFromString(code);
        }
        
        success = true;
      } finally {
        // Release the component COM object
        releaseComObject(component);
      }
      
      return success;
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to update module code for ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Creates a new module in the active project
   * @param moduleName The name of the module
   * @param moduleType The type of module to create
   * @param initialCode Initial code for the module
   * @returns Success status
   * @throws {McpError} If the module cannot be created
   */
  public async createModule(
    moduleName: string,
    moduleType: vbext_ComponentType,
    initialCode: string = ''
  ): Promise<boolean> {
    try {
      this.ensureConnected();
      
      const activeProject = this.comVBE.ActiveVBProject;
      if (!activeProject) {
        throw new McpError(
          ErrorCode.NotFound,
          'No active VB project found'
        );
      }
      
      const components = activeProject.VBComponents;
      if (!components) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access project components'
        );
      }
      
      // Check if module already exists
      try {
        const existingComponent = components.Item(moduleName);
        if (existingComponent) {
          releaseComObject(existingComponent);
          throw new McpError(
            ErrorCode.InvalidRequest,
            `Module with name ${moduleName} already exists`
          );
        }
      } catch (error) {
        // Module not found, which is what we want
        if (!(error instanceof McpError)) {
          // Expected error when module doesn't exist, continue
        }
      }
      
      // Create the new component
      const component = components.Add(moduleType);
      if (!component) {
        throw new McpError(
          ErrorCode.InternalError,
          `Failed to create module of type ${moduleType}`
        );
      }
      
      // Set the name
      component.Name = moduleName;
      
      // Add initial code if provided
      if (initialCode.length > 0) {
        const codeModule = component.CodeModule;
        if (codeModule) {
          codeModule.AddFromString(initialCode);
        }
      }
      
      // Release the component COM object
      releaseComObject(component);
      
      return true;
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to create module ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Removes a module from the active project
   * @param moduleName The name of the module to remove
   * @returns Success status
   * @throws {McpError} If the module cannot be removed
   */
  public async removeModule(moduleName: string): Promise<boolean> {
    try {
      this.ensureConnected();
      
      const activeProject = this.comVBE.ActiveVBProject;
      if (!activeProject) {
        throw new McpError(
          ErrorCode.NotFound,
          'No active VB project found'
        );
      }
      
      const components = activeProject.VBComponents;
      if (!components) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access project components'
        );
      }
      
      // Find the component
      let component;
      try {
        component = components.Item(moduleName);
      } catch (error) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      if (!component) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      // Remove the component
      try {
        components.Remove(component);
        return true;
      } finally {
        // Release the component COM object
        releaseComObject(component);
      }
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to remove module ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Exports a module to a file
   * @param moduleName The name of the module to export
   * @param filePath The path to export to
   * @returns Success status
   * @throws {McpError} If the module cannot be exported
   */
  public async exportModule(moduleName: string, filePath: string): Promise<boolean> {
    try {
      this.ensureConnected();
      
      const activeProject = this.comVBE.ActiveVBProject;
      if (!activeProject) {
        throw new McpError(
          ErrorCode.NotFound,
          'No active VB project found'
        );
      }
      
      const components = activeProject.VBComponents;
      if (!components) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access project components'
        );
      }
      
      // Find the component
      let component;
      try {
        component = components.Item(moduleName);
      } catch (error) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      if (!component) {
        throw new McpError(
          ErrorCode.NotFound,
          `Module not found: ${moduleName}`
        );
      }
      
      // Export the component
      try {
        component.Export(filePath);
        return true;
      } catch (error) {
        throw new McpError(
          ErrorCode.InternalError,
          `Failed to export module ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
        );
      } finally {
        // Release the component COM object
        releaseComObject(component);
      }
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to export module ${moduleName}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Imports a module from a file
   * @param filePath The path to import from
   * @returns The name of the imported module
   * @throws {McpError} If the module cannot be imported
   */
  public async importModule(filePath: string): Promise<string> {
    try {
      this.ensureConnected();
      
      const activeProject = this.comVBE.ActiveVBProject;
      if (!activeProject) {
        throw new McpError(
          ErrorCode.NotFound,
          'No active VB project found'
        );
      }
      
      const components = activeProject.VBComponents;
      if (!components) {
        throw new McpError(
          ErrorCode.InternalError,
          'Failed to access project components'
        );
      }
      
      // Import the component
      const component = components.Import(filePath);
      if (!component) {
        throw new McpError(
          ErrorCode.InternalError,
          `Failed to import module from ${filePath}`
        );
      }
      
      // Get the module name
      const moduleName = String(component.Name);
      
      // Release the component COM object
      releaseComObject(component);
      
      return moduleName;
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to import module from ${filePath}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  /**
   * Releases COM resources
   * This should be called when the wrapper is no longer needed
   */
  public dispose(): void {
    try {
      if (this.comVBE) {
        // Release the COM object
        releaseComObject(this.comVBE);
        this.comVBE = null;
        this.isConnected = false;
      }
    } catch (error) {
      console.error('Error disposing VBE COM wrapper:', error);
    }
  }
}
