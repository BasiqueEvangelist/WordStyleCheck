FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /app

COPY WordStyleCheck/WordStyleCheck.csproj ./WordStyleCheck/
COPY WordStyleCheckService.Worker/WordStyleCheckService.Worker.csproj ./WordStyleCheckService.Worker/
COPY WordStyleCheckService.Foreman/WordStyleCheckService.Foreman.csproj ./WordStyleCheckService.Foreman/
RUN dotnet restore WordStyleCheckService.Foreman

COPY WordStyleCheck/. ./WordStyleCheck/
COPY WordStyleCheckService.Worker/. ./WordStyleCheckService.Worker/
COPY WordStyleCheckService.Foreman/. ./WordStyleCheckService.Foreman/
RUN dotnet publish WordStyleCheckService.Foreman -c Release -o out


FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS runtime
WORKDIR /app
COPY --from=build /app/out ./
ENTRYPOINT ["dotnet", "WordStyleCheckService.Foreman.dll"]
