/**
 * TypeScript interfaces for VBE (Visual Basic Editor) COM objects
 */

/**
 * Represents the VBE application object
 */
export interface IVBE {
  /** Gets the active VBE version */
  readonly Version: string;
  
  /** Gets the active VBE project */
  readonly ActiveVBProject: IVBProject;
  
  /** Gets the collection of all projects */
  readonly VBProjects: IVBProjects;
  
  /** Gets the code window */
  readonly CodeWindow: ICodeWindow;
  
  /** Gets the active code pane */
  readonly ActiveCodePane: ICodePane;
}

/**
 * Represents a collection of VB projects
 */
export interface IVBProjects {
  /** Gets the number of projects */
  readonly Count: number;
  
  /** Gets a project by index */
  Item(index: number): IVBProject;
  
  /** Gets a project by name */
  Item(name: string): IVBProject;
}

/**
 * Represents a VB project
 */
export interface IVBProject {
  /** Gets the project name */
  readonly Name: string;
  
  /** Gets the project filename */
  readonly FileName: string;
  
  /** Gets the project description */
  readonly Description: string;
  
  /** Gets the collection of VBComponents in the project */
  readonly VBComponents: IVBComponents;
  
  /** Gets the collection of references in the project */
  readonly References: IReferences;
  
  /** Gets the project type */
  readonly Type: vbext_ProjectType;
  
  /** Gets whether the project is protected */
  readonly Protection: vbext_ProjectProtection;
  
  /** Saves the project */
  SaveAs(filename: string): void;
}

/**
 * Enumeration of project types
 */
export enum vbext_ProjectType {
  vbext_pt_HostProject = 100,
  vbext_pt_StandAlone = 101
}

/**
 * Enumeration of project protection types
 */
export enum vbext_ProjectProtection {
  vbext_pp_none = 0,
  vbext_pp_locked = 1
}

/**
 * Represents a collection of VB components (modules)
 */
export interface IVBComponents {
  /** Gets the number of components */
  readonly Count: number;
  
  /** Gets a component by index */
  Item(index: number): IVBComponent;
  
  /** Gets a component by name */
  Item(name: string): IVBComponent;
  
  /** Adds a new component of the specified type */
  Add(componentType: vbext_ComponentType): IVBComponent;
  
  /** Imports a component from a file */
  Import(path: string): IVBComponent;
  
  /** Removes a component */
  Remove(component: IVBComponent): void;
}

/**
 * Represents a single VB component (module, class, form)
 */
export interface IVBComponent {
  /** Gets or sets the component name */
  Name: string;
  
  /** Gets the component type */
  readonly Type: vbext_ComponentType;
  
  /** Gets the code module */
  readonly CodeModule: ICodeModule;
  
  /** Gets or sets the component's properties */
  readonly Properties: IProperties;
  
  /** Exports the component to a file */
  Export(path: string): void;
}

/**
 * Enumeration of component types
 */
export enum vbext_ComponentType {
  vbext_ct_StdModule = 1,
  vbext_ct_ClassModule = 2,
  vbext_ct_MSForm = 3,
  vbext_ct_ActiveXDesigner = 11,
  vbext_ct_Document = 100
}

/**
 * Represents a code module (the actual code in a component)
 */
export interface ICodeModule {
  /** Gets the parent component */
  readonly Parent: IVBComponent;
  
  /** Gets the number of lines in the module */
  readonly CountOfLines: number;
  
  /** Gets or sets the name of the module */
  Name: string;
  
  /** Gets a line of code */
  Lines(startLine: number, count: number): string;
  
  /** Replaces a line of code */
  ReplaceLine(line: number, code: string): void;
  
  /** Inserts lines of code */
  InsertLines(line: number, code: string): void;
  
  /** Deletes lines of code */
  DeleteLines(startLine: number, count: number): void;
  
  /** Adds code to the end of the module */
  AddFromString(code: string): void;
  
  /** Gets all code from the module */
  GetText(): string;
}

/**
 * Represents a collection of references
 */
export interface IReferences {
  /** Gets the number of references */
  readonly Count: number;
  
  /** Gets a reference by index */
  Item(index: number): IReference;
  
  /** Adds a reference to a type library */
  AddFromGuid(guid: string, major: number, minor: number): IReference;
  
  /** Adds a reference to a file */
  AddFromFile(path: string): IReference;
  
  /** Removes a reference */
  Remove(reference: IReference): void;
}

/**
 * Represents a single reference
 */
export interface IReference {
  /** Gets the name of the reference */
  readonly Name: string;
  
  /** Gets the path of the reference */
  readonly FullPath: string;
  
  /** Gets the GUID of the reference */
  readonly GUID: string;
  
  /** Gets the major version of the reference */
  readonly Major: number;
  
  /** Gets the minor version of the reference */
  readonly Minor: number;
  
  /** Gets whether the reference is built-in */
  readonly BuiltIn: boolean;
  
  /** Gets whether the reference is broken */
  readonly IsBroken: boolean;
  
  /** Gets the type of the reference */
  readonly Type: vbext_RefKind;
}

/**
 * Enumeration of reference types
 */
export enum vbext_RefKind {
  vbext_rk_TypeLib = 0,
  vbext_rk_Project = 1
}

/**
 * Represents a collection of properties
 */
export interface IProperties {
  /** Gets the number of properties */
  readonly Count: number;
  
  /** Gets a property by index */
  Item(index: number): IProperty;
  
  /** Gets a property by name */
  Item(name: string): IProperty;
}

/**
 * Represents a single property
 */
export interface IProperty {
  /** Gets the property name */
  readonly Name: string;
  
  /** Gets or sets the property value */
  Value: any;
  
  /** Gets the property index */
  readonly Index: number;
}

/**
 * Represents the code window
 */
export interface ICodeWindow {
  /** Gets the code pane */
  readonly CodePane: ICodePane;
  
  /** Gets the visible state of the window */
  readonly Visible: boolean;
  
  /** Gets the window state */
  readonly WindowState: vbext_WindowState;
}

/**
 * Enumeration of window states
 */
export enum vbext_WindowState {
  vbext_ws_Normal = 0,
  vbext_ws_Minimize = 1,
  vbext_ws_Maximize = 2
}

/**
 * Represents a code pane
 */
export interface ICodePane {
  /** Gets the code module */
  readonly CodeModule: ICodeModule;
  
  /** Gets the current selection */
  GetSelection(): {
    startLine: number;
    startColumn: number;
    endLine: number;
    endColumn: number;
  };
  
  /** Sets the current selection */
  SetSelection(startLine: number, startColumn: number, endLine: number, endColumn: number): void;
  
  /** Shows a specific line */
  ShowLine(line: number): void;
}
