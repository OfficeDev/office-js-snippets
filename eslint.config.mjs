import stylistic from '@stylistic/eslint-plugin';
import typescript from '@typescript-eslint/eslint-plugin';
import parser from '@typescript-eslint/parser';

export default [
    {
        files: ['config/**/*.ts'],
        ignores: ['config/**/*.d.ts'],
        languageOptions: { parser },
        plugins: {
            '@stylistic': stylistic,
            '@typescript-eslint': typescript
        },
        rules: {
            '@typescript-eslint/triple-slash-reference': ['error', { path: 'never' }],
            '@typescript-eslint/no-namespace': ['error', { allowDeclarations: true }],
            '@typescript-eslint/naming-convention': [
                'error',
                { selector: ['class', 'interface'], format: ['PascalCase'] }
            ],
            '@typescript-eslint/no-redeclare': 'error',
            '@typescript-eslint/no-unused-expressions': 'error',
            'curly': 'error',
            'guard-for-in': 'error',
            'no-caller': 'error',
            'no-var': 'error',
            'no-restricted-properties': [
                'error',
                { object: 'console', property: 'time' },
                { object: 'console', property: 'timeEnd' }
            ],
            'no-new-wrappers': 'error',
            'no-debugger': 'error',
            'no-eval': 'error',
            'no-restricted-syntax': [
                'error',
                {
                    selector: 'FunctionExpression[id=null]:not(MethodDefinition > FunctionExpression):not(Property[method=true] > FunctionExpression)',
                    message: 'Use an arrow function for anonymous function expressions.'
                }
            ],
            'no-labels': ['error', { allowLoop: true, allowSwitch: true }],
            'radix': 'error',
            'eqeqeq': ['error', 'always', { null: 'ignore' }],
            '@stylistic/eol-last': ['error', 'always'],
            '@stylistic/no-tabs': 'error',
            '@stylistic/no-trailing-spaces': 'error',
            '@stylistic/brace-style': ['error', '1tbs', { allowSingleLine: true }],
            '@stylistic/space-before-blocks': 'error',
            '@stylistic/quotes': ['error', 'single', { allowTemplateLiterals: 'always' }],
            '@stylistic/semi': ['error', 'always'],
            '@stylistic/keyword-spacing': 'error',
            '@stylistic/space-infix-ops': 'error',
            '@stylistic/comma-spacing': 'error',
            '@stylistic/semi-spacing': 'error',
            '@stylistic/type-annotation-spacing': 'error'
        }
    }
];
