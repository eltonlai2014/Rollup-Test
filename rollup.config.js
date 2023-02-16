import resolve from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import { uglify } from "rollup-plugin-uglify";

export default [
    {
        input: ['src/index.js'],
        external: ['canvas', 'redux'],  // 指定不打包外部lib
        output: {
            name: 'ReportExcel',
            file: './dist/bundle.js',
            format: 'cjs'
        },
        plugins: [
            resolve(), // so Rollup can find `ms`
            commonjs(), // so Rollup can convert `ms` to an ES module
            // uglify(),
        ]
    },
    // {
    //     input: 'src/ReportExcelSite.js',
    //     output: {
    //         name: 'ReportExcel',
    //         file: './dist/bundle.js',
    //         format: 'cjs'
    //     },
    //     plugins: [
    //         resolve(), // so Rollup can find `ms`
    //         commonjs(), // so Rollup can convert `ms` to an ES module
    //         // uglify(),
    //     ]
    // },
];
