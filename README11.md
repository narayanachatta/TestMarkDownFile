# embedded-tools

## [dump-memory.py](dump-memory-script\readme.md)
This script is created to capture the data from SC4 and dump it into a binary file.
The binary file with data dumped in it is used by [create-stack-usage-report.py](stack-usage-report-script/create-stack-usage-report.py) to generate a worst-case stack usage report.

## [create-stack-usage-report.py](stack-usage-report-script/readme.md)
This script is created to identify the task stacks and generate worst-case stack usage,
based on the memory dump files generated using [dump-memory.py](dump-memory-script\dump-memory.py) script.

## [Stack_Usage_Analysis.py](..\Stack_Usage_Analysis.py)
This script is created to perform the following actions in a sequence:
- path and memory dump file name are passed to [dump-memory.py](dump-memory-script\dump-memory.py) script
- capturing of data dump from SC4 by [dump-memory.py](dump-memory-script\dump-memory.py) script and store it in the path above passed as argument
- map file path and memory dump file path are passed to [create-stack-usage-report.py](stack-usage-report-script/create-stack-usage-report.py) script
- generate worst-case stack usage report by [create-stack-usage-report.py](stack-usage-report-script/create-stack-usage-report.py) script using map files and memory dump files