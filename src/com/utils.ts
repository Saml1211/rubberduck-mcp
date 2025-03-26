/**
 * COM interoperability utility functions
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';

/**
 * Default timeout for COM operations in milliseconds (20 seconds)
 */
export const DEFAULT_COM_TIMEOUT = 20000;

/**
 * Default maximum retries for COM operations
 */
export const DEFAULT_MAX_RETRIES = 2;

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
      // Alternative approach - check for gc in a more TypeScript-friendly way
      } else {
        // Try to access gc if available (may be exposed in Node.js with --expose-gc flag)
        const gc = (globalThis as any).gc;
        if (typeof gc === 'function') {
          gc();
        }
      }
    } catch (error) {
      console.error('Error releasing COM object:', error);
    }
  }
}

/**
 * Creates a wrapper function that ensures COM objects are properly released
 * 
 * @template T The return type of the wrapped function
 * @template Args The argument types of the wrapped function
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

/**
 * Executes a function with a timeout
 * 
 * @template T The return type of the promise
 * @param promise The promise to execute with a timeout
 * @param timeoutMs Timeout in milliseconds
 * @returns A promise that resolves with the result or rejects with a timeout error
 */
export async function withTimeout<T>(
  promise: Promise<T>,
  timeoutMs: number = DEFAULT_COM_TIMEOUT
): Promise<T> {
  // Create a timeout promise that rejects after timeoutMs
  const timeoutPromise = new Promise<never>((_, reject) => {
    setTimeout(() => {
      reject(new McpError(
        ErrorCode.InternalError,
        `Operation timed out after ${timeoutMs}ms`
      ));
    }, timeoutMs);
  });

  // Race the real operation against the timeout
  return Promise.race([promise, timeoutPromise]);
}

/**
 * Retries a COM operation for a specified number of times
 * 
 * @template T The return type of the operation
 * @param operation The operation to retry
 * @param maxRetries Maximum number of retries
 * @param retryDelayMs Delay between retries in milliseconds
 * @returns A promise that resolves with the operation result
 */
export async function retryComOperation<T>(
  operation: () => Promise<T>,
  maxRetries: number = DEFAULT_MAX_RETRIES,
  retryDelayMs: number = 1000
): Promise<T> {
  let lastError: Error | undefined;
  
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return await operation();
    } catch (error) {
      lastError = error instanceof Error ? error : new Error(String(error));
      
      // Don't delay on the last attempt
      if (attempt < maxRetries) {
        await new Promise(resolve => setTimeout(resolve, retryDelayMs));
      }
    }
  }
  
  // If we get here, all retries failed
  throw lastError || new McpError(
    ErrorCode.InternalError,
    'COM operation failed after multiple retries'
  );
}

/**
 * Executes a function with a COM object and ensures it's released after use
 * 
 * @template T The return type of the function
 * @param comObject The COM object to use
 * @param fn The function to execute with the COM object
 * @returns The result of the function
 */
export function usingComObject<T>(
  comObject: any,
  fn: (obj: any) => T
): T {
  try {
    return fn(comObject);
  } finally {
    releaseComObject(comObject);
  }
}

/**
 * Safely retrieves a property from a COM object, handling errors
 * 
 * @template T The expected type of the property
 * @param comObject The COM object to get the property from
 * @param propertyName The name of the property to get
 * @param defaultValue Default value to return if property access fails
 * @returns The property value or default value
 */
export function getComProperty<T>(
  comObject: any,
  propertyName: string,
  defaultValue: T
): T {
  try {
    if (!comObject) {
      return defaultValue;
    }
    
    const value = comObject[propertyName];
    return value !== undefined ? value : defaultValue;
  } catch (error) {
    console.error(`Error accessing COM property ${propertyName}:`, error);
    return defaultValue;
  }
}

/**
 * Type guard to check if an object is an Error instance
 * @param error The object to check
 * @returns True if the object is an Error instance
 */
export function isError(error: unknown): error is Error {
  return error instanceof Error;
}

/**
 * Type guard to check if an object is a McpError instance
 * @param error The object to check
 * @returns True if the object is a McpError instance
 */
export function isMcpError(error: unknown): error is McpError {
  return error instanceof McpError;
}

/**
 * Converts a generic error to an McpError with appropriate error code
 * @param error The error to convert
 * @param defaultMessage Default message to use if the error doesn't have one
 * @returns An McpError instance
 */
export function toMcpError(
  error: unknown,
  defaultMessage: string = 'An unknown error occurred'
): McpError {
  if (isMcpError(error)) {
    return error;
  }
  
  const message = isError(error) ? error.message : String(error) || defaultMessage;
  
  // Check for common error patterns to determine the appropriate error code
  const errorStr = message.toLowerCase();
  
  if (errorStr.includes('timeout') || errorStr.includes('timed out')) {
    return new McpError(ErrorCode.InternalError, message);
  }
  
  if (errorStr.includes('not found') || errorStr.includes('does not exist')) {
    return new McpError(ErrorCode.NotFound, message);
  }
  
  if (errorStr.includes('permission') || errorStr.includes('access denied') || 
      errorStr.includes('unauthorized')) {
    return new McpError(ErrorCode.Unauthorized, message);
  }
  
  if (errorStr.includes('invalid') || errorStr.includes('argument') || 
      errorStr.includes('parameter')) {
    return new McpError(ErrorCode.InvalidParams, message);
  }
  
  return new McpError(ErrorCode.InternalError, message);
}
