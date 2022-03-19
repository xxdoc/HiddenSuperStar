cd %TEMP%\HSSTemp
"%~dp0iconv.exe" -f utf-8 -t gbk %1_orig.json  > %1.json