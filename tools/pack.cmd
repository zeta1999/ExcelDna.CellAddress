set outputPath=D:\Projects\packages\Excel-Dna\
set Configuration=Release
set YYYYmmdd=%date:~0,4%%date:~5,2%%date:~8,2%
Set suffix=build%YYYYmmdd%

Set version-suffix=--version-suffix %suffix%


cd ../src/
dotnet build ExcelDna.CellAddress  --configuration %Configuration%
dotnet pack ExcelDna.CellAddress  --configuration %Configuration%  %version-suffix%  -o %outputPath%
cd ../tools/
