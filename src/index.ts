#!/usr/bin/env node
/**
 * rubberduck-mcp - Main entry point
 * 
 * An MCP server for integrating AI assistants with Rubberduck VBA.
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { initializeTools, releaseComResources } from './server/tools.js';
import { initializeResources } from './server/resources.js';
import { parseArgs } from 'node:util';

// Version from package.json
const VERSION = '0.1.0';

/**
 * Parse command line arguments
 */
function parseCommandLineArgs() {
  const options = {
    'api-key': { type: 'string' },
    'debug': { type: 'boolean' },
    'com-timeout': { type: 'string' },
    'max-retries': { type: 'string' },
    'help': { type: 'boolean', short: 'h' },
    'version': { type: 'boolean', short: 'v' }
  };

  // Parse command line arguments
  const {
    values
  } = parseArgs({
    args: process.argv.slice(2),
    options,
    allowPositionals: false
  });

  // Show help if requested
  if (values.help) {
    console.log(
`rubberduck-mcp - Model Context Protocol server for Rubberduck VBA

Usage: rubberduck-mcp [options]

Options:
  --api-key <key>      API key for authentication (env: RUBBERDUCK_API_KEY)
  --debug              Enable debug logging (env: RUBBERDUCK_DEBUG)
  --com-timeout <ms>   Timeout for COM operations in ms (env: RUBBERDUCK_COM_TIMEOUT)
  --max-retries <num>  Maximum retries for COM operations (env: RUBBERDUCK_MAX_RETRIES)
  -h, --help           Show this help message
  -v, --version        Show version information
`
    );
    process.exit(0);
  }

  // Show version if requested
  if (values.version) {
    console.log(`rubberduck-mcp v${VERSION}`);
    process.exit(0);
  }

  // Get configuration from environment variables or command line arguments
  const config = {
    apiKey: values['api-key'] || process.env.RUBBERDUCK_API_KEY,
    debug: values.debug || process.env.RUBBERDUCK_DEBUG === 'true',
    comTimeout: values['com-timeout'] ? parseInt(values['com-timeout'], 10) 
               : process.env.RUBBERDUCK_COM_TIMEOUT ? parseInt(process.env.RUBBERDUCK_COM_TIMEOUT, 10) 
               : undefined,
    maxRetries: values['max-retries'] ? parseInt(values['max-retries'], 10) 
               : process.env.RUBBERDUCK_MAX_RETRIES ? parseInt(process.env.RUBBERDUCK_MAX_RETRIES, 10) 
               : undefined
  };

  return config;
}

/**
 * Main function
 */
async function main() {
  try {
    const config = parseCommandLineArgs();

    // Initialize the MCP server
    const server = new Server(
      {
        name: 'rubberduck-mcp',
        version: VERSION,
      },
      {
        capabilities: {
          tools: {},
          resources: {},
        },
      }
    );

    // Set up tools and resources
    initializeTools(server, {
      apiKey: config.apiKey,
      debug: config.debug,
      comTimeout: config.comTimeout,
      maxRetries: config.maxRetries
    });
    
    initializeResources(server, {
      apiKey: config.apiKey,
      debug: config.debug,
      comTimeout: config.comTimeout,
      maxRetries: config.maxRetries
    });

    // Set up error handler
    server.onerror = (error) => {
      console.error('[MCP Error]', error);
    };

    // Set up close handler for proper COM resource cleanup
    server.onclose = async () => {
      releaseComResources();
    };

    // Connect to stdio transport
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    if (config.debug) {
      console.error('Rubberduck MCP server running on stdio');
    }

    // Handle SIGINT/SIGTERM for graceful shutdown
    const handleShutdown = async () => {
      if (config.debug) {
        console.error('Shutting down...');
      }
      
      await server.close();
      process.exit(0);
    };

    process.on('SIGINT', handleShutdown);
    process.on('SIGTERM', handleShutdown);
  } catch (error) {
    console.error('Failed to start server:', error);
    process.exit(1);
  }
}

// Start the server
main().catch((error) => {
  console.error('Unhandled error:', error);
  process.exit(1);
});
