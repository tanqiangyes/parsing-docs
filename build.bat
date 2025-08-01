@echo off
setlocal enabledelayedexpansion

:: Docs Parser 构建脚本

:: 设置变量
set BINARY_NAME=docs-parser.exe
set BUILD_DIR=build
set TEST_DIR=tests
set EXAMPLES_DIR=examples

:: 版本信息
set VERSION=1.0.0
set BUILD_TIME=%date:~0,4%-%date:~5,2%-%date:~8,2%_%time:~0,2%:%time:~3,2%:%time:~6,2%
set BUILD_TIME=%BUILD_TIME: =0%

:: 构建标志
set LDFLAGS=-ldflags "-X main.Version=%VERSION% -X main.BuildTime=%BUILD_TIME%"

:: 检查参数
if "%1"=="" goto help
if "%1"=="build" goto build
if "%1"=="test" goto test
if "%1"=="clean" goto clean
if "%1"=="run" goto run
if "%1"=="install" goto install
if "%1"=="deps" goto deps
if "%1"=="fmt" goto fmt
if "%1"=="lint" goto lint
if "%1"=="help" goto help
goto help

:build
echo 构建 Docs Parser...
if not exist %BUILD_DIR% mkdir %BUILD_DIR%
go build %LDFLAGS% -o %BUILD_DIR%\%BINARY_NAME% cmd/main.go
if %errorlevel% equ 0 (
    echo 构建完成: %BUILD_DIR%\%BINARY_NAME%
) else (
    echo 构建失败
    exit /b 1
)
goto end

:test
echo 运行测试...
go test -v ./%TEST_DIR%/...
if %errorlevel% equ 0 (
    echo 测试完成
) else (
    echo 测试失败
    exit /b 1
)
goto end

:clean
echo 清理构建文件...
if exist %BUILD_DIR% rmdir /s /q %BUILD_DIR%
if exist coverage.out del coverage.out
if exist coverage.html del coverage.html
if exist *.exe del *.exe
echo 清理完成
goto end

:run
echo 运行 Docs Parser...
if exist %BUILD_DIR%\%BINARY_NAME% (
    %BUILD_DIR%\%BINARY_NAME% --help
) else (
    echo 请先构建项目: build.bat build
    exit /b 1
)
goto end

:install
echo 安装到系统...
go install %LDFLAGS% ./cmd/main.go
if %errorlevel% equ 0 (
    echo 安装完成
) else (
    echo 安装失败
    exit /b 1
)
goto end

:deps
echo 安装依赖...
go mod tidy
go mod download
if %errorlevel% equ 0 (
    echo 依赖安装完成
) else (
    echo 依赖安装失败
    exit /b 1
)
goto end

:fmt
echo 格式化代码...
go fmt ./...
if %errorlevel% equ 0 (
    echo 代码格式化完成
) else (
    echo 代码格式化失败
    exit /b 1
)
goto end

:lint
echo 检查代码...
go vet ./...
if %errorlevel% equ 0 (
    echo 代码检查完成
) else (
    echo 代码检查失败
    exit /b 1
)
goto end

:help
echo Docs Parser 构建脚本帮助
echo.
echo 可用命令:
echo   build    - 构建项目
echo   test     - 运行测试
echo   clean    - 清理构建文件
echo   run      - 运行项目
echo   install  - 安装到系统
echo   deps     - 安装依赖
echo   fmt      - 格式化代码
echo   lint     - 检查代码
echo   help     - 显示此帮助信息
echo.
echo 示例:
echo   build.bat build     # 构建项目
echo   build.bat test      # 运行测试
echo   build.bat run       # 运行项目
echo.
goto end

:end
endlocal 