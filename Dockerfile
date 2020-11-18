FROM mcr.microsoft.com/dotnet/core/sdk:3.1 AS build

WORKDIR /source

# copy csproj and restore as distinct layers
#COPY *.sln .
COPY src/*.csproj ./aspnetapp/
WORKDIR /source/aspnetapp
RUN dotnet restore

# copy everything else and build app
COPY src/  /source/aspnetapp/
RUN ls -al /source/aspnetapp/

WORKDIR /source/aspnetapp
RUN dotnet publish -c release -o /app --no-restore

# final stage/image
FROM mcr.microsoft.com/dotnet/core/aspnet:3.1
WORKDIR /app
COPY --from=build /app ./

CMD ASPNETCORE_URLS=http://*:$PORT dotnet replace-text-docx.dll

