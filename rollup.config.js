import resolve from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';

export default [
  // ESM build
  {
    input: 'src/XlsxTemplater.js',
    output: {
      file: 'dist/xlsx-templater.esm.js',
      format: 'es'
    },
    plugins: [resolve(), commonjs()]
  },
  // CommonJS build
  {
    input: 'src/XlsxTemplater.js',
    output: {
      file: 'dist/xlsx-templater.cjs.js',
      format: 'cjs'
    },
    plugins: [resolve(), commonjs()]
  },
  // UMD build
  {
    input: 'src/XlsxTemplater.js',
    output: {
      file: 'dist/xlsx-templater.umd.js',
      format: 'umd',
      name: 'XlsxTemplater'
    },
    plugins: [resolve(), commonjs()]
  }
]; 