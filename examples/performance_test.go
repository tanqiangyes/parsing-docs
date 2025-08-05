package examples

import (
	"fmt"
	"log"
	"path/filepath"
	"runtime"
	"time"

	"docs-parser/internal/core/types"
	"docs-parser/pkg/parser"
)

func PerformanceTestExample() {
	fmt.Println("=== Docs Parser 性能测试 ===")

	// 测试传统解析器
	testTraditionalParser()

	// 测试流式解析器
	testStreamingParser()

	// 测试批量处理器
	testBatchProcessor()

	// 测试内存池
	testMemoryPool()

	// 测试并发处理
	testConcurrentProcessing()
}

// testTraditionalParser 测试传统解析器
func testTraditionalParser() {
	fmt.Println("\n--- 传统解析器测试 ---")

	parser := parser.NewParser()

	// 创建测试文件列表
	testFiles := createTestFiles()

	start := time.Now()

	for i, file := range testFiles {
		doc, err := parser.ParseDocument(file)
		if err != nil {
			log.Printf("解析文件 %s 失败: %v", file, err)
			continue
		}

		fmt.Printf("文件 %d: %s - 解析成功，段落数: %d\n", i+1, filepath.Base(file), len(doc.Content.Paragraphs))
	}

	duration := time.Since(start)
	fmt.Printf("传统解析器总耗时: %v\n", duration)
}

// testStreamingParser 测试流式解析器
func testStreamingParser() {
	fmt.Println("\n--- 流式解析器测试 ---")

	streamingParser := parser.NewStreamingParser(8192, 4)

	// 创建测试文件列表
	testFiles := createTestFiles()

	start := time.Now()

	for i, file := range testFiles {
		resultChan, err := streamingParser.ParseStream(file)
		if err != nil {
			log.Printf("流式解析文件 %s 失败: %v", file, err)
			continue
		}

		for result := range resultChan {
			if result.Error != nil {
				log.Printf("解析文件 %s 失败: %v", file, result.Error)
				continue
			}

			fmt.Printf("文件 %d: %s - 流式解析成功，耗时: %v\n",
				i+1, filepath.Base(file), result.Duration)
		}
	}

	duration := time.Since(start)
	fmt.Printf("流式解析器总耗时: %v\n", duration)
}

// testBatchProcessor 测试批量处理器
func testBatchProcessor() {
	fmt.Println("\n--- 批量处理器测试 ---")

	batchProcessor := parser.NewBatchProcessor(4)
	defer batchProcessor.Close()

	// 创建测试文件列表
	testFiles := createTestFiles()

	start := time.Now()

	// 批量处理文件
	documents, errors := batchProcessor.ProcessFiles(testFiles)

	duration := time.Since(start)

	// 统计结果
	successCount := 0
	errorCount := 0

	for i, doc := range documents {
		if errors[i] != nil {
			errorCount++
			log.Printf("文件 %s 处理失败: %v", filepath.Base(testFiles[i]), errors[i])
		} else {
			successCount++
			fmt.Printf("文件 %d: %s - 批量处理成功，段落数: %d\n",
				i+1, filepath.Base(testFiles[i]), len(doc.Content.Paragraphs))
		}
	}

	fmt.Printf("批量处理完成 - 成功: %d, 失败: %d, 总耗时: %v\n",
		successCount, errorCount, duration)

	// 显示统计信息
	stats := batchProcessor.GetStats()
	fmt.Printf("处理器统计 - 处理任务: %d, 失败任务: %d, 平均时间: %v\n",
		stats.JobsProcessed, stats.JobsFailed, stats.AverageTime)
}

// testMemoryPool 测试内存池
func testMemoryPool() {
	fmt.Println("\n--- 内存池测试 ---")

	poolManager := parser.GetGlobalPoolManager()

	// 重置统计信息
	poolManager.Reset()

	// 模拟大量对象创建和复用
	for i := 0; i < 1000; i++ {
		// 获取解析器
		parser := poolManager.GetParser()

		// 获取文档对象
		doc := poolManager.GetDocument()

		// 获取缓冲区
		buffer := poolManager.GetBuffer()

		// 模拟使用
		doc.Metadata.Title = fmt.Sprintf("测试文档 %d", i)
		doc.Content.Paragraphs = append(doc.Content.Paragraphs, types.Paragraph{
			ID: fmt.Sprintf("p_%d", i),
		})

		// 归还对象到池中
		poolManager.PutParser(parser)
		poolManager.PutDocument(doc)
		poolManager.PutBuffer(buffer)
	}

	// 显示池统计信息
	stats := poolManager.GetStats()
	fmt.Println("内存池统计信息:")
	for name, stat := range stats {
		fmt.Printf("  %s - 创建: %d, 复用: %d, 丢弃: %d\n",
			name, stat.Created, stat.Reused, stat.Discarded)
	}
}

// testConcurrentProcessing 测试并发处理
func testConcurrentProcessing() {
	fmt.Println("\n--- 并发处理测试 ---")

	// 创建大量测试文件
	testFiles := createLargeTestFileList(100)

	// 测试不同并发数
	workerCounts := []int{1, 2, 4, 8}

	for _, workers := range workerCounts {
		fmt.Printf("\n测试 %d 个工作协程:\n", workers)

		batchProcessor := parser.NewBatchProcessor(workers)

		start := time.Now()

		// 使用回调函数处理结果
		processedCount := 0
		batchProcessor.ProcessFilesWithCallback(testFiles, func(index int, doc *types.Document, err error) {
			processedCount++
			if err != nil {
				log.Printf("文件 %d 处理失败: %v", index, err)
			} else {
				if processedCount%10 == 0 {
					fmt.Printf("已处理 %d 个文件...\n", processedCount)
				}
			}
		})

		duration := time.Since(start)

		// 显示统计信息
		stats := batchProcessor.GetStats()
		fmt.Printf("  - 总耗时: %v\n", duration)
		fmt.Printf("  - 处理任务: %d\n", stats.JobsProcessed)
		fmt.Printf("  - 失败任务: %d\n", stats.JobsFailed)
		fmt.Printf("  - 平均时间: %v\n", stats.AverageTime)

		batchProcessor.Close()
	}
}

// createTestFiles 创建测试文件列表
func createTestFiles() []string {
	// 这里应该创建实际的测试文件
	// 为了演示，我们使用模拟的文件路径
	return []string{
		"test_files/sample1.docx",
		"test_files/sample2.doc",
		"test_files/sample3.rtf",
		"test_files/sample4.wpd",
	}
}

// createLargeTestFileList 创建大量测试文件列表
func createLargeTestFileList(count int) []string {
	files := make([]string, count)
	for i := 0; i < count; i++ {
		files[i] = fmt.Sprintf("test_files/sample_%d.docx", i+1)
	}
	return files
}

// 性能基准测试
func benchmarkParser() {
	fmt.Println("\n--- 性能基准测试 ---")

	// 获取系统信息
	fmt.Printf("系统信息:\n")
	fmt.Printf("  - CPU核心数: %d\n", runtime.NumCPU())
	fmt.Printf("  - 内存信息: %s\n", getMemoryInfo())

	// 测试不同大小的文件
	fileSizes := []string{"small", "medium", "large"}

	for _, size := range fileSizes {
		fmt.Printf("\n测试 %s 文件:\n", size)

		// 传统解析器测试
		start := time.Now()
		// 这里应该解析实际的文件
		duration := time.Since(start)
		fmt.Printf("  - 传统解析器: %v\n", duration)

		// 流式解析器测试
		start = time.Now()
		// 这里应该进行流式解析
		duration = time.Since(start)
		fmt.Printf("  - 流式解析器: %v\n", duration)
	}
}

// getMemoryInfo 获取内存信息
func getMemoryInfo() string {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	return fmt.Sprintf("已分配: %d MB, 系统: %d MB",
		m.Alloc/1024/1024, m.Sys/1024/1024)
}

// 内存使用监控
func monitorMemoryUsage() {
	fmt.Println("\n--- 内存使用监控 ---")

	// 启动内存监控
	go func() {
		ticker := time.NewTicker(1 * time.Second)
		defer ticker.Stop()

		for range ticker.C {
			var m runtime.MemStats
			runtime.ReadMemStats(&m)

			fmt.Printf("内存使用 - 已分配: %d MB, 堆内存: %d MB, GC次数: %d\n",
				m.Alloc/1024/1024, m.HeapAlloc/1024/1024, m.NumGC)
		}
	}()

	// 运行一些测试
	time.Sleep(5 * time.Second)
}

// 性能优化建议
func performanceRecommendations() {
	fmt.Println("\n--- 性能优化建议 ---")

	fmt.Println("1. 流式处理:")
	fmt.Println("   - 对于大文件，使用流式解析器")
	fmt.Println("   - 设置合适的缓冲区大小")
	fmt.Println("   - 避免一次性加载整个文件")

	fmt.Println("\n2. 内存管理:")
	fmt.Println("   - 使用对象池减少GC压力")
	fmt.Println("   - 及时归还对象到池中")
	fmt.Println("   - 监控内存使用情况")

	fmt.Println("\n3. 并发处理:")
	fmt.Println("   - 根据CPU核心数设置工作协程数")
	fmt.Println("   - 使用批量处理器处理多个文件")
	fmt.Println("   - 避免创建过多协程")

	fmt.Println("\n4. 缓存策略:")
	fmt.Println("   - 缓存解析结果")
	fmt.Println("   - 复用解析器实例")
	fmt.Println("   - 使用连接池")
}
