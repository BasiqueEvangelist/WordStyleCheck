#!/usr/bin/sh -x

rm -rf artifacts/
mkdir -p artifacts/wsc-linux-x64
mkdir -p artifacts/wsc-win-x64
mkdir -p artifacts/wsc-no-framework
mkdir -p artifacts/WordStyleCheckGui-linux-x64
mkdir -p artifacts/WordStyleCheckGui-win-x64
mkdir -p artifacts/WordStyleCheckGui-no-framework

dotnet publish wsc --configuration Release -p:PublishSingleFile=true --self-contained -o artifacts/wsc-linux-x64/ -r linux-x64
dotnet publish wsc --configuration Release -p:PublishSingleFile=true --self-contained -o artifacts/wsc-win-x64/ -r win-x64
dotnet publish wsc --configuration Release --no-self-contained -o artifacts/wsc-no-framework/
dotnet publish WordStyleCheckGui --configuration Release -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true --self-contained -o artifacts/WordStyleCheckGui-linux-x64/ -r linux-x64
dotnet publish WordStyleCheckGui --configuration Release -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true --self-contained -o artifacts/WordStyleCheckGui-win-x64/ -r win-x64
dotnet publish WordStyleCheckGui --configuration Release --no-self-contained -o artifacts/WordStyleCheckGui-no-framework/

rm artifacts/**.pdb

mv artifacts/