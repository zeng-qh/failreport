FROM mcr.microsoft.com/dotnet/sdk:10.0 AS publish
WORKDIR /src 
# 仅复制项目文件，还原依赖（单独成层，修改代码不触发重新下载）
COPY ["FailReport.csproj", "./"]
RUN dotnet restore "FailReport.csproj"


# FROM base AS publish
# WORKDIR /src
COPY . .
RUN dotnet build "FailReport.csproj" -c Release --no-restore \
    && echo "构建产物验证：" \
    && ls -la /src/bin/Release/net10.0/ \
    # 发布项目（不使用--no-build，让publish确保所有必要文件都被包含）
    && dotnet publish "FailReport.csproj" -c Release -o /app/publish \
       --no-build --no-restore \
    && echo "发布产物验证：" \
    && ls -la /app/publish \
    && rm -rf /src/obj /src/bin  # 清理构建残留

# 运行阶段（精简镜像，仅保留运行时依赖）
FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS final
WORKDIR /app

# 创建非root用户并切换
RUN groupadd -r appuser && useradd -r -g appuser appuser
USER appuser

COPY --from=publish /app/publish .
EXPOSE 999 
ENTRYPOINT ["dotnet", "FailReport.dll"]
