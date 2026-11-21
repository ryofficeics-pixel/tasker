#!/bin/bash
cd "$(dirname "$0")"
cat mock_head.js mock_ext.js ../Code.gs tests_main.js tests_web.js tests_end.js > run_tests.js
TZ=Asia/Jakarta node run_tests.js
