import resolve from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import json from '@rollup/plugin-json';

const external = [
  'exceljs',
  'path',
  'fs',
  'url',
  'node-fetch',
  'lodash',
  'lodash/cloneDeep',
  'tr46',
  'whatwg-url',
  'webidl-conversions',
  'buffer',
  'stream',
  'string_decoder',
  'events',
  'util',
  'assert',
  'crypto'
];

export default [
  // CommonJS build
  {
    input: 'src/XlsxTemplater.js',
    output: {
      file: 'dist/xlsxtemplater.cjs.js',
      format: 'cjs',
      exports: 'auto'
    },
    external,
    plugins: [
      resolve({ 
        preferBuiltins: true,
        browser: false
      }), 
      commonjs({
        ignoreDynamicRequires: true,
        transformMixedEsModules: true
      }),
      json()
    ]
  },
  // ESM build
  {
    input: 'src/XlsxTemplater.js',
    output: {
      file: 'dist/xlsxtemplater.esm.js',
      format: 'es'
    },
    external,
    plugins: [
      resolve({ 
        preferBuiltins: true,
        browser: false
      }), 
      commonjs({
        ignoreDynamicRequires: true,
        transformMixedEsModules: true
      }),
      json()
    ]
  },
  // UMD build
  {
    input: 'src/XlsxTemplater.js',
    output: {
      file: 'dist/xlsxtemplater.umd.js',
      format: 'umd',
      name: 'XlsxTemplater',
      globals: {
        exceljs: 'ExcelJS',
        'node-fetch': 'fetch',
        'lodash': '_',
        'lodash/cloneDeep': '_.cloneDeep'
      }
    },
    external,
    plugins: [
      resolve({ 
        preferBuiltins: true,
        browser: true
      }), 
      commonjs({
        ignoreDynamicRequires: true,
        transformMixedEsModules: true
      }),
      json()
    ]
  }
]; 