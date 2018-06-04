set outputPath=D:\Projects\packages\Excel-Dna\
set Configuration=Release
set YYYYmmdd=%date:~0,4%%date:~5,2%%date:~8,2%
Set suffix=build%YYYYmmdd%

Set version-suffix=-Suffix %suffix%


cd ../src/ExcelDna.CellAddress
dotnet pack -c %Configuration% -o %outputPath%
REM --version-suffix %version-sufix%  Configuration=%Configuration%
