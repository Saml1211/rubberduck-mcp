/**
 * Type declarations for @modelcontextprotocol/sdk modules
 */

declare module '@modelcontextprotocol/sdk/types.js' {
  /**
   * Error codes for MCP errors
   */
  export enum ErrorCode {
    ParseError = -32700,
    InvalidRequest = -32600,
    MethodNotFound = -32601,
    InvalidParams = -32602,
    InternalError = -32603,
    ServerError = -32000,
    Unauthorized = 401,
    Forbidden = 403,
    NotFound = 404,
    RequestTimeout = 408,
    TooManyRequests = 429,
    ServerUnavailable = 503
  }
  
  /**
   * Standard MCP error
   */
  export class McpError extends Error {
    code: ErrorCode;
    data?: any;
    
    constructor(code: ErrorCode, message: string, data?: any);
  }
  
  /**
   * Schema for list tools request
   */
  export const ListToolsRequestSchema: {
    method: string;
    params: any;
  };
  
  /**
   * Schema for call tool request
   */
  export const CallToolRequestSchema: {
    method: string;
    params: any;
  };
  
  /**
   * Schema for list resources request
   */
  export const ListResourcesRequestSchema: {
    method: string;
    params: any;
  };
  
  /**
   * Schema for read resource request
   */
  export const ReadResourceRequestSchema: {
    method: string;
    params: any;
  };
  
  /**
   * Schema for list resource templates request
   */
  export const ListResourceTemplatesRequestSchema: {
    method: string;
    params: any;
  };
}

declare module '@modelcontextprotocol/sdk/server/index.js' {
  import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
  
  export interface ServerOptions {
    name: string;
    version: string;
  }
  
  export interface ServerCapabilities {
    capabilities: {
      tools?: any;
      resources?: any;
    };
  }
  
  export interface RequestHandler {
    (request: any): Promise<any>;
  }
  
  export interface ServerTransport {
    connect(server: Server): Promise<void>;
    close(): Promise<void>;
  }
  
  export class Server {
    constructor(options: ServerOptions, capabilities: ServerCapabilities);
    
    setRequestHandler(schema: any, handler: RequestHandler): void;
    
    connect(transport: ServerTransport): Promise<void>;
    
    close(): Promise<void>;
    
    onerror: ((error: any) => void) | null;
    
    onclose: (() => Promise<void>) | null;
  }
}

declare module '@modelcontextprotocol/sdk/server/stdio.js' {
  import { Server } from '@modelcontextprotocol/sdk/server/index.js';
  
  export class StdioServerTransport {
    connect(server: Server): Promise<void>;
    
    close(): Promise<void>;
  }
}
