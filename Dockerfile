#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/aspnet:7.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443
ENV SimpleProperty="hello-from-BASE-dockerfile"

FROM mcr.microsoft.com/dotnet/sdk:7.0 AS build
WORKDIR /src
COPY ["OYA-Web-App/OYA_Web_Application.csproj", "OYA_Web_Application/"]
RUN dotnet restore "OYA_Web_Application/OYA_Web_Application.csproj"
COPY . .
WORKDIR "/src/LearnSmartCoding.API"
RUN dotnet build "OYA_Web_Application/OYA_Web_Application.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "OYA_Web_Application/OYA_Web_Application.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENV SimpleProperty="hello-from-FINAL-dockerfile"
ENTRYPOINT ["dotnet", "OYA_Web_Application.dll"]
