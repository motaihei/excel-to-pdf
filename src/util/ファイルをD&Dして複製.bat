@echo off
set /p num="複製するファイルの数を入力してください："

for /l %%i in (1,1,%num%) do (
    for %%f in ("%~1") do (
        copy "%%~ff" "%%~dpnf_%%i%%~xf"
    )
)
