/**
 * Type declarations for node-win32com package
 */

declare module 'node-win32com' {
  export interface Win32COM {
    /**
     * Creates a COM object for a given program ID
     * @param progId The ProgID of the COM object to create
     * @returns The created COM object
     */
    createObject(progId: string): any;
    
    /**
     * Gets a running COM object
     * @param progId The ProgID of the running COM object
     * @returns The running COM object or null if not found
     */
    getObject(progId: string): any;
    
    /**
     * Releases a COM object
     * @param comObject The COM object to release
     */
    releaseObject(comObject: any): void;
  }
  
  export const win32com: Win32COM;
}
