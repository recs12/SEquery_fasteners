
:: Set this file for compiling the executable of the macro.

ipyc.exe /main:__main__.py ^
helper.py ^
query.py ^
Interop.SolidEdge.dll ^
/embed ^
/out:query_fasteners ^
/platform:x64 ^
/standalone ^
/target:exe ^
/win32icon:icon.ico
