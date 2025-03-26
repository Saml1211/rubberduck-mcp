module.exports = {
  parser: '@typescript-eslint/parser',
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/recommended',
  ],
  parserOptions: {
    ecmaVersion: 2022,
    sourceType: 'module',
  },
  plugins: ['@typescript-eslint'],
  env: {
    node: true,
    es2022: true,
  },
  rules: {
    // COM interop often requires the use of 'any'
    '@typescript-eslint/no-explicit-any': 'off',
    
    // Some COM methods may not be used directly, but are part of interfaces
    '@typescript-eslint/no-unused-vars': ['warn', { 
      'argsIgnorePattern': '^_',
      'varsIgnorePattern': '^_' 
    }],
    
    // Enforce consistent error handling
    'no-throw-literal': 'error',
    
    // Require the use of === and !==
    'eqeqeq': ['error', 'always', { 'null': 'ignore' }],
    
    // Warning for TODO, FIXME, etc.
    'no-warning-comments': ['warn', { 
      'terms': ['todo', 'fixme', 'hack'], 
      'location': 'start' 
    }],
    
    // Enforce consistent use of semicolons
    'semi': ['error', 'always'],
    
    // Enforce consistent spacing inside braces
    'object-curly-spacing': ['error', 'always'],
    
    // Enforce consistent indentation
    'indent': ['error', 2, { 'SwitchCase': 1 }],
    
    // Prevent debugger statements
    'no-debugger': 'error',
    
    // Require function return types for better documentation
    '@typescript-eslint/explicit-function-return-type': ['warn', {
      'allowExpressions': true,
      'allowTypedFunctionExpressions': true
    }],
    
    // Enforce proper rejection of promises
    '@typescript-eslint/no-floating-promises': 'warn',
    
    // COM methods might require method overloads
    '@typescript-eslint/unified-signatures': 'off',
    
    // Helpful for code organization
    '@typescript-eslint/explicit-member-accessibility': ['error', { 
      'accessibility': 'explicit',
      'overrides': { 'constructors': 'no-public' }
    }],
    
    // Clean up code
    'no-console': ['warn', { 'allow': ['error', 'warn', 'info'] }],
  },
  ignorePatterns: ['dist/', 'node_modules/'],
};
