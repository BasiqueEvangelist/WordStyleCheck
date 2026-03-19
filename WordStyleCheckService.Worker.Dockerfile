FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /app

COPY WordStyleCheck/WordStyleCheck.csproj ./WordStyleCheck/
COPY WordStyleCheckService.Worker/WordStyleCheckService.Worker.csproj ./WordStyleCheckService.Worker/
RUN dotnet restore WordStyleCheckService.Worker

COPY WordStyleCheck/. ./WordStyleCheck/
COPY WordStyleCheckService.Worker/. ./WordStyleCheckService.Worker/
RUN dotnet publish WordStyleCheckService.Worker -c Release -o out


FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS runtime
WORKDIR /app
COPY --from=build /app/out ./
ENTRYPOINT ["dotnet", "WordStyleCheckService.Worker.dll"]
