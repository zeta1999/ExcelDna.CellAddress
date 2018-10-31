set Configuration=Release
cd ../src/
dotnet build ExcelDna.CellAddress  --configuration %Configuration%
cd ../tools/