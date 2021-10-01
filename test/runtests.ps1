$ErrorActionPreference = "Ignore"
del bin\result -Recurse
$ErrorActionPreference = "Stop"
dotnet test --collect:"XPlat Code Coverage" -r bin\result
reportgenerator -reports:.\bin\result\*\coverage.cobertura.xml -targetdir:.\bin\result