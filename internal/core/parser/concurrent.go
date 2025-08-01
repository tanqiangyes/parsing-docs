package parser

import (
	"context"
	"fmt"
	"sync"
	"time"

	"docs-parser/internal/core/types"
)

// Job 处理任务
type Job struct {
	ID       string
	FilePath string
	Priority int
}

// Result 处理结果
type Result struct {
	JobID     string
	Document  *types.Document
	Error     error
	Duration  time.Duration
	StartTime time.Time
	EndTime   time.Time
}

// ConcurrentProcessor 并发处理器
type ConcurrentProcessor struct {
	workers    int
	jobQueue   chan Job
	resultChan chan Result
	ctx        context.Context
	cancel     context.CancelFunc
	wg         sync.WaitGroup
	stats      ProcessorStats
	mu         sync.RWMutex
}

// ProcessorStats 处理器统计信息
type ProcessorStats struct {
	JobsProcessed int64
	JobsFailed    int64
	TotalDuration time.Duration
	AverageTime   time.Duration
	LastJobTime   time.Time
}

// NewConcurrentProcessor 创建新的并发处理器
func NewConcurrentProcessor(workers int) *ConcurrentProcessor {
	ctx, cancel := context.WithCancel(context.Background())

	return &ConcurrentProcessor{
		workers:    workers,
		jobQueue:   make(chan Job, workers*2),
		resultChan: make(chan Result, workers*2),
		ctx:        ctx,
		cancel:     cancel,
		stats:      ProcessorStats{LastJobTime: time.Now()},
	}
}

// Start 启动并发处理器
func (cp *ConcurrentProcessor) Start() {
	for i := 0; i < cp.workers; i++ {
		cp.wg.Add(1)
		go cp.worker(i)
	}
}

// Stop 停止并发处理器
func (cp *ConcurrentProcessor) Stop() {
	cp.cancel()
	close(cp.jobQueue)
	cp.wg.Wait()
	close(cp.resultChan)
}

// SubmitJob 提交任务
func (cp *ConcurrentProcessor) SubmitJob(job Job) error {
	select {
	case cp.jobQueue <- job:
		return nil
	case <-cp.ctx.Done():
		return fmt.Errorf("processor is stopped")
	default:
		return fmt.Errorf("job queue is full")
	}
}

// GetResults 获取结果通道
func (cp *ConcurrentProcessor) GetResults() <-chan Result {
	return cp.resultChan
}

// ProcessBatch 批量处理文件
func (cp *ConcurrentProcessor) ProcessBatch(files []string) []Result {
	// 启动处理器
	cp.Start()
	defer cp.Stop()

	// 提交所有任务
	for i, file := range files {
		job := Job{
			ID:       fmt.Sprintf("job_%d", i),
			FilePath: file,
			Priority: 1,
		}
		cp.SubmitJob(job)
	}

	// 收集结果
	var results []Result
	for range files {
		select {
		case result := <-cp.resultChan:
			results = append(results, result)
		case <-time.After(30 * time.Second):
			// 超时处理
			break
		}
	}

	return results
}

// worker 工作协程
func (cp *ConcurrentProcessor) worker(id int) {
	defer cp.wg.Done()

	// 获取全局池管理器
	poolManager := GetGlobalPoolManager()

	for {
		select {
		case job := <-cp.jobQueue:
			cp.processJob(job, poolManager)
		case <-cp.ctx.Done():
			return
		}
	}
}

// processJob 处理单个任务
func (cp *ConcurrentProcessor) processJob(job Job, poolManager *GlobalPoolManager) {
	startTime := time.Now()

	// 从池中获取解析器
	parser := poolManager.GetParser()
	defer poolManager.PutParser(parser)

	// 从池中获取文档对象
	doc := poolManager.GetDocument()
	defer poolManager.PutDocument(doc)

	// 解析文档
	err := cp.parseDocument(parser, job.FilePath, doc)

	endTime := time.Now()
	duration := endTime.Sub(startTime)

	// 更新统计信息
	cp.updateStats(duration, err)

	// 发送结果
	result := Result{
		JobID:     job.ID,
		Document:  doc,
		Error:     err,
		Duration:  duration,
		StartTime: startTime,
		EndTime:   endTime,
	}

	select {
	case cp.resultChan <- result:
	case <-cp.ctx.Done():
		return
	}
}

// parseDocument 解析文档
func (cp *ConcurrentProcessor) parseDocument(parser Parser, filePath string, doc *types.Document) error {
	// 使用解析器解析文档
	parsedDoc, err := parser.ParseDocument(filePath)
	if err != nil {
		return fmt.Errorf("failed to parse document: %w", err)
	}

	// 复制解析结果到池中的文档对象
	*doc = *parsedDoc

	return nil
}

// updateStats 更新统计信息
func (cp *ConcurrentProcessor) updateStats(duration time.Duration, err error) {
	cp.mu.Lock()
	defer cp.mu.Unlock()

	cp.stats.JobsProcessed++
	cp.stats.TotalDuration += duration
	cp.stats.AverageTime = cp.stats.TotalDuration / time.Duration(cp.stats.JobsProcessed)
	cp.stats.LastJobTime = time.Now()

	if err != nil {
		cp.stats.JobsFailed++
	}
}

// GetStats 获取统计信息
func (cp *ConcurrentProcessor) GetStats() ProcessorStats {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.stats
}

// Reset 重置统计信息
func (cp *ConcurrentProcessor) Reset() {
	cp.mu.Lock()
	cp.stats = ProcessorStats{LastJobTime: time.Now()}
	cp.mu.Unlock()
}

// BatchProcessor 批量处理器
type BatchProcessor struct {
	processor *ConcurrentProcessor
	pool      *GlobalPoolManager
}

// NewBatchProcessor 创建新的批量处理器
func NewBatchProcessor(workers int) *BatchProcessor {
	return &BatchProcessor{
		processor: NewConcurrentProcessor(workers),
		pool:      GetGlobalPoolManager(),
	}
}

// ProcessFiles 批量处理文件
func (bp *BatchProcessor) ProcessFiles(files []string) ([]*types.Document, []error) {
	// 启动处理器
	bp.processor.Start()
	defer bp.processor.Stop()

	// 提交所有任务
	for i, file := range files {
		job := Job{
			ID:       fmt.Sprintf("batch_%d", i),
			FilePath: file,
			Priority: 1,
		}
		bp.processor.SubmitJob(job)
	}

	// 收集结果
	var documents []*types.Document
	var errors []error

	for range files {
		select {
		case result := <-bp.processor.GetResults():
			if result.Error != nil {
				errors = append(errors, result.Error)
				documents = append(documents, nil)
			} else {
				documents = append(documents, result.Document)
				errors = append(errors, nil)
			}
		case <-time.After(60 * time.Second):
			// 超时处理
			break
		}
	}

	return documents, errors
}

// ProcessFilesWithCallback 带回调的批量处理
func (bp *BatchProcessor) ProcessFilesWithCallback(files []string, callback func(int, *types.Document, error)) {
	// 启动处理器
	bp.processor.Start()
	defer bp.processor.Stop()

	// 提交所有任务
	for i, file := range files {
		job := Job{
			ID:       fmt.Sprintf("callback_%d", i),
			FilePath: file,
			Priority: 1,
		}
		bp.processor.SubmitJob(job)
	}

	// 处理结果
	for i := range files {
		select {
		case result := <-bp.processor.GetResults():
			callback(i, result.Document, result.Error)
		case <-time.After(60 * time.Second):
			// 超时处理
			callback(i, nil, fmt.Errorf("processing timeout"))
		}
	}
}

// GetStats 获取统计信息
func (bp *BatchProcessor) GetStats() ProcessorStats {
	return bp.processor.GetStats()
}

// GetPoolStats 获取池统计信息
func (bp *BatchProcessor) GetPoolStats() map[string]PoolStats {
	return bp.pool.GetStats()
}

// Reset 重置统计信息
func (bp *BatchProcessor) Reset() {
	bp.processor.Reset()
	bp.pool.Reset()
}

// Close 关闭处理器
func (bp *BatchProcessor) Close() {
	bp.processor.Stop()
	bp.pool.Close()
}

// StreamingBatchProcessor 流式批量处理器
type StreamingBatchProcessor struct {
	processor *ConcurrentProcessor
	pool      *GlobalPoolManager
}

// NewStreamingBatchProcessor 创建新的流式批量处理器
func NewStreamingBatchProcessor(workers int) *StreamingBatchProcessor {
	return &StreamingBatchProcessor{
		processor: NewConcurrentProcessor(workers),
		pool:      GetGlobalPoolManager(),
	}
}

// ProcessFilesStream 流式批量处理文件
func (sbp *StreamingBatchProcessor) ProcessFilesStream(files []string) (<-chan Result, error) {
	resultChan := make(chan Result, len(files))

	// 启动处理器
	sbp.processor.Start()

	// 在后台处理结果
	go func() {
		defer close(resultChan)
		defer sbp.processor.Stop()

		// 提交所有任务
		for i, file := range files {
			job := Job{
				ID:       fmt.Sprintf("stream_%d", i),
				FilePath: file,
				Priority: 1,
			}
			sbp.processor.SubmitJob(job)
		}

		// 流式返回结果
		for range files {
			select {
			case result := <-sbp.processor.GetResults():
				resultChan <- result
			case <-time.After(60 * time.Second):
				// 超时处理
				break
			}
		}
	}()

	return resultChan, nil
}

// GetStats 获取统计信息
func (sbp *StreamingBatchProcessor) GetStats() ProcessorStats {
	return sbp.processor.GetStats()
}

// GetPoolStats 获取池统计信息
func (sbp *StreamingBatchProcessor) GetPoolStats() map[string]PoolStats {
	return sbp.pool.GetStats()
}

// Reset 重置统计信息
func (sbp *StreamingBatchProcessor) Reset() {
	sbp.processor.Reset()
	sbp.pool.Reset()
}

// Close 关闭处理器
func (sbp *StreamingBatchProcessor) Close() {
	sbp.processor.Stop()
	sbp.pool.Close()
}
