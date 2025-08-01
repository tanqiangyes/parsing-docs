# Docs Parser Makefile

# 变量定义
BINARY_NAME=docs-parser
BUILD_DIR=build
TEST_DIR=tests
EXAMPLES_DIR=examples
DOCS_DIR=docs

# Go 相关变量
GO=go
GOOS?=$(shell go env GOOS)
GOARCH?=$(shell go env GOARCH)

# 版本信息
VERSION?=1.0.0
BUILD_TIME=$(shell date -u '+%Y-%m-%d_%H:%M:%S')
GIT_COMMIT=$(shell git rev-parse --short HEAD 2>/dev/null || echo "unknown")

# 构建标志
LDFLAGS=-ldflags "-X main.Version=$(VERSION) -X main.BuildTime=$(BUILD_TIME) -X main.GitCommit=$(GIT_COMMIT)"

# 默认目标
.PHONY: all
all: clean build test

# 构建项目
.PHONY: build
build:
	@echo "构建 Docs Parser..."
	@mkdir -p $(BUILD_DIR)
	$(GO) build $(LDFLAGS) -o $(BUILD_DIR)/$(BINARY_NAME) cmd/main.go
	@echo "构建完成: $(BUILD_DIR)/$(BINARY_NAME)"

# 构建 Windows 版本
.PHONY: build-windows
build-windows:
	@echo "构建 Windows 版本..."
	@mkdir -p $(BUILD_DIR)
	GOOS=windows GOARCH=amd64 $(GO) build $(LDFLAGS) -o $(BUILD_DIR)/$(BINARY_NAME).exe cmd/main.go
	@echo "构建完成: $(BUILD_DIR)/$(BINARY_NAME).exe"

# 构建 Linux 版本
.PHONY: build-linux
build-linux:
	@echo "构建 Linux 版本..."
	@mkdir -p $(BUILD_DIR)
	GOOS=linux GOARCH=amd64 $(GO) build $(LDFLAGS) -o $(BUILD_DIR)/$(BINARY_NAME)-linux cmd/main.go
	@echo "构建完成: $(BUILD_DIR)/$(BINARY_NAME)-linux"

# 构建 macOS 版本
.PHONY: build-macos
build-macos:
	@echo "构建 macOS 版本..."
	@mkdir -p $(BUILD_DIR)
	GOOS=darwin GOARCH=amd64 $(GO) build $(LDFLAGS) -o $(BUILD_DIR)/$(BINARY_NAME)-macos cmd/main.go
	@echo "构建完成: $(BUILD_DIR)/$(BINARY_NAME)-macos"

# 构建所有平台版本
.PHONY: build-all
build-all: build-windows build-linux build-macos
	@echo "所有平台版本构建完成"

# 运行测试
.PHONY: test
test:
	@echo "运行测试..."
	$(GO) test -v ./$(TEST_DIR)/...
	@echo "测试完成"

# 运行基准测试
.PHONY: benchmark
benchmark:
	@echo "运行基准测试..."
	$(GO) test -bench=. ./$(TEST_DIR)/...
	@echo "基准测试完成"

# 运行测试并生成覆盖率报告
.PHONY: test-coverage
test-coverage:
	@echo "运行测试并生成覆盖率报告..."
	$(GO) test -v -coverprofile=coverage.out ./$(TEST_DIR)/...
	$(GO) tool cover -html=coverage.out -o coverage.html
	@echo "覆盖率报告已生成: coverage.html"

# 运行示例
.PHONY: run-example
run-example:
	@echo "运行示例..."
	$(GO) run $(EXAMPLES_DIR)/basic_usage.go

# 安装依赖
.PHONY: deps
deps:
	@echo "安装依赖..."
	$(GO) mod tidy
	$(GO) mod download
	@echo "依赖安装完成"

# 代码格式化
.PHONY: fmt
fmt:
	@echo "格式化代码..."
	$(GO) fmt ./...
	@echo "代码格式化完成"

# 代码检查
.PHONY: lint
lint:
	@echo "检查代码..."
	$(GO) vet ./...
	@echo "代码检查完成"

# 清理构建文件
.PHONY: clean
clean:
	@echo "清理构建文件..."
	rm -rf $(BUILD_DIR)
	rm -f coverage.out coverage.html
	rm -f *.exe
	@echo "清理完成"

# 安装到系统
.PHONY: install
install:
	@echo "安装到系统..."
	$(GO) install $(LDFLAGS) ./cmd/main.go
	@echo "安装完成"

# 运行项目
.PHONY: run
run: build
	@echo "运行 Docs Parser..."
	./$(BUILD_DIR)/$(BINARY_NAME) --help

# 创建发布包
.PHONY: release
release: build-all
	@echo "创建发布包..."
	@mkdir -p release
	@cp $(BUILD_DIR)/* release/
	@cp README.md release/
	@cp LICENSE release/ 2>/dev/null || echo "LICENSE 文件不存在"
	@cp examples/config.json release/
	@echo "发布包已创建: release/"

# 显示帮助信息
.PHONY: help
help:
	@echo "Docs Parser Makefile 帮助"
	@echo ""
	@echo "可用目标:"
	@echo "  build          - 构建项目"
	@echo "  build-windows  - 构建 Windows 版本"
	@echo "  build-linux    - 构建 Linux 版本"
	@echo "  build-macos    - 构建 macOS 版本"
	@echo "  build-all      - 构建所有平台版本"
	@echo "  test           - 运行测试"
	@echo "  benchmark      - 运行基准测试"
	@echo "  test-coverage  - 运行测试并生成覆盖率报告"
	@echo "  run-example    - 运行示例"
	@echo "  deps           - 安装依赖"
	@echo "  fmt            - 格式化代码"
	@echo "  lint           - 检查代码"
	@echo "  clean          - 清理构建文件"
	@echo "  install        - 安装到系统"
	@echo "  run            - 运行项目"
	@echo "  release        - 创建发布包"
	@echo "  help           - 显示此帮助信息"
	@echo ""
	@echo "示例:"
	@echo "  make build     # 构建项目"
	@echo "  make test      # 运行测试"
	@echo "  make run       # 运行项目"

# 开发模式：监听文件变化并重新构建
.PHONY: dev
dev:
	@echo "开发模式：监听文件变化..."
	@if command -v air >/dev/null 2>&1; then \
		air; \
	else \
		echo "请安装 air: go install github.com/cosmtrek/air@latest"; \
		echo "或者手动运行: make build && make run"; \
	fi

# 检查代码质量
.PHONY: check
check: fmt lint test
	@echo "代码质量检查完成"

# 完整构建流程
.PHONY: full
full: clean deps check build test-coverage
	@echo "完整构建流程完成"

# 默认目标
.DEFAULT_GOAL := help 