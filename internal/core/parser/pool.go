package parser

import (
	"sync"
	"time"

	"docs-parser/internal/core/types"
)

// ParserPool 解析器对象池，用于减少内存分配和GC压力
type ParserPool struct {
	parsers chan Parser
	maxSize int
	mu      sync.Mutex
	stats   PoolStats
}

// PoolStats 池统计信息
type PoolStats struct {
	Created    int64
	Reused     int64
	Discarded  int64
	LastAccess time.Time
}

// NewParserPool 创建新的解析器池
func NewParserPool(maxSize int) *ParserPool {
	return &ParserPool{
		parsers: make(chan Parser, maxSize),
		maxSize: maxSize,
		stats:   PoolStats{LastAccess: time.Now()},
	}
}

// Get 从池中获取解析器
func (pp *ParserPool) Get() Parser {
	pp.mu.Lock()
	pp.stats.LastAccess = time.Now()
	pp.mu.Unlock()

	select {
	case parser := <-pp.parsers:
		pp.mu.Lock()
		pp.stats.Reused++
		pp.mu.Unlock()
		return parser
	default:
		// 池为空，创建新的解析器
		pp.mu.Lock()
		pp.stats.Created++
		pp.mu.Unlock()
		// 创建一个简单的解析器实现
		factory := NewParserFactory()
		// 注册默认解析器
		factory.RegisterParser("docx", &DefaultParser{})
		return factory
	}
}

// Put 归还解析器到池中
func (pp *ParserPool) Put(parser Parser) {
	// 重置解析器状态
	if resetter, ok := parser.(interface{ Reset() }); ok {
		resetter.Reset()
	}

	select {
	case pp.parsers <- parser:
		// 成功归还到池中
	default:
		// 池已满，丢弃解析器
		pp.mu.Lock()
		pp.stats.Discarded++
		pp.mu.Unlock()
	}
}

// GetStats 获取池统计信息
func (pp *ParserPool) GetStats() PoolStats {
	pp.mu.Lock()
	defer pp.mu.Unlock()
	return pp.stats
}

// Reset 重置池统计信息
func (pp *ParserPool) Reset() {
	pp.mu.Lock()
	pp.stats = PoolStats{LastAccess: time.Now()}
	pp.mu.Unlock()
}

// Close 关闭池
func (pp *ParserPool) Close() {
	close(pp.parsers)
}

// DocumentPool 文档对象池
type DocumentPool struct {
	documents chan *types.Document
	maxSize   int
	mu        sync.Mutex
	stats     PoolStats
}

// NewDocumentPool 创建新的文档池
func NewDocumentPool(maxSize int) *DocumentPool {
	return &DocumentPool{
		documents: make(chan *types.Document, maxSize),
		maxSize:   maxSize,
		stats:     PoolStats{LastAccess: time.Now()},
	}
}

// Get 从池中获取文档对象
func (dp *DocumentPool) Get() *types.Document {
	dp.mu.Lock()
	dp.stats.LastAccess = time.Now()
	dp.mu.Unlock()

	select {
	case doc := <-dp.documents:
		dp.mu.Lock()
		dp.stats.Reused++
		dp.mu.Unlock()
		// 重置文档对象
		*doc = types.Document{}
		return doc
	default:
		// 池为空，创建新的文档对象
		dp.mu.Lock()
		dp.stats.Created++
		dp.mu.Unlock()
		return &types.Document{}
	}
}

// Put 归还文档对象到池中
func (dp *DocumentPool) Put(doc *types.Document) {
	// 重置文档对象
	*doc = types.Document{}

	select {
	case dp.documents <- doc:
		// 成功归还到池中
	default:
		// 池已满，丢弃文档对象
		dp.mu.Lock()
		dp.stats.Discarded++
		dp.mu.Unlock()
	}
}

// GetStats 获取池统计信息
func (dp *DocumentPool) GetStats() PoolStats {
	dp.mu.Lock()
	defer dp.mu.Unlock()
	return dp.stats
}

// Reset 重置池统计信息
func (dp *DocumentPool) Reset() {
	dp.mu.Lock()
	dp.stats = PoolStats{LastAccess: time.Now()}
	dp.mu.Unlock()
}

// Close 关闭池
func (dp *DocumentPool) Close() {
	close(dp.documents)
}

// BufferPool 缓冲区池
type BufferPool struct {
	buffers chan []byte
	size    int
	maxSize int
	mu      sync.Mutex
	stats   PoolStats
}

// NewBufferPool 创建新的缓冲区池
func NewBufferPool(size, maxSize int) *BufferPool {
	return &BufferPool{
		buffers: make(chan []byte, maxSize),
		size:    size,
		maxSize: maxSize,
		stats:   PoolStats{LastAccess: time.Now()},
	}
}

// Get 从池中获取缓冲区
func (bp *BufferPool) Get() []byte {
	bp.mu.Lock()
	bp.stats.LastAccess = time.Now()
	bp.mu.Unlock()

	select {
	case buffer := <-bp.buffers:
		bp.mu.Lock()
		bp.stats.Reused++
		bp.mu.Unlock()
		// 清空缓冲区
		for i := range buffer {
			buffer[i] = 0
		}
		return buffer
	default:
		// 池为空，创建新的缓冲区
		bp.mu.Lock()
		bp.stats.Created++
		bp.mu.Unlock()
		return make([]byte, bp.size)
	}
}

// Put 归还缓冲区到池中
func (bp *BufferPool) Put(buffer []byte) {
	// 检查缓冲区大小是否匹配
	if len(buffer) != bp.size {
		// 大小不匹配，丢弃
		bp.mu.Lock()
		bp.stats.Discarded++
		bp.mu.Unlock()
		return
	}

	select {
	case bp.buffers <- buffer:
		// 成功归还到池中
	default:
		// 池已满，丢弃缓冲区
		bp.mu.Lock()
		bp.stats.Discarded++
		bp.mu.Unlock()
	}
}

// GetStats 获取池统计信息
func (bp *BufferPool) GetStats() PoolStats {
	bp.mu.Lock()
	defer bp.mu.Unlock()
	return bp.stats
}

// Reset 重置池统计信息
func (bp *BufferPool) Reset() {
	bp.mu.Lock()
	bp.stats = PoolStats{LastAccess: time.Now()}
	bp.mu.Unlock()
}

// Close 关闭池
func (bp *BufferPool) Close() {
	close(bp.buffers)
}

// GlobalPoolManager 全局池管理器
type GlobalPoolManager struct {
	parserPool   *ParserPool
	documentPool *DocumentPool
	bufferPool   *BufferPool
	mu           sync.RWMutex
}

// NewGlobalPoolManager 创建全局池管理器
func NewGlobalPoolManager() *GlobalPoolManager {
	return &GlobalPoolManager{
		parserPool:   NewParserPool(10),
		documentPool: NewDocumentPool(20),
		bufferPool:   NewBufferPool(8192, 50), // 8KB缓冲区，最多50个
	}
}

// GetParser 获取解析器
func (gpm *GlobalPoolManager) GetParser() Parser {
	return gpm.parserPool.Get()
}

// PutParser 归还解析器
func (gpm *GlobalPoolManager) PutParser(parser Parser) {
	gpm.parserPool.Put(parser)
}

// GetDocument 获取文档对象
func (gpm *GlobalPoolManager) GetDocument() *types.Document {
	return gpm.documentPool.Get()
}

// PutDocument 归还文档对象
func (gpm *GlobalPoolManager) PutDocument(doc *types.Document) {
	gpm.documentPool.Put(doc)
}

// GetBuffer 获取缓冲区
func (gpm *GlobalPoolManager) GetBuffer() []byte {
	return gpm.bufferPool.Get()
}

// PutBuffer 归还缓冲区
func (gpm *GlobalPoolManager) PutBuffer(buffer []byte) {
	gpm.bufferPool.Put(buffer)
}

// GetStats 获取所有池的统计信息
func (gpm *GlobalPoolManager) GetStats() map[string]PoolStats {
	gpm.mu.RLock()
	defer gpm.mu.RUnlock()

	return map[string]PoolStats{
		"parser":   gpm.parserPool.GetStats(),
		"document": gpm.documentPool.GetStats(),
		"buffer":   gpm.bufferPool.GetStats(),
	}
}

// Reset 重置所有池的统计信息
func (gpm *GlobalPoolManager) Reset() {
	gpm.mu.Lock()
	defer gpm.mu.Unlock()

	gpm.parserPool.Reset()
	gpm.documentPool.Reset()
	gpm.bufferPool.Reset()
}

// Close 关闭所有池
func (gpm *GlobalPoolManager) Close() {
	gpm.mu.Lock()
	defer gpm.mu.Unlock()

	gpm.parserPool.Close()
	gpm.documentPool.Close()
	gpm.bufferPool.Close()
}

// 全局池管理器实例
var globalPoolManager *GlobalPoolManager
var once sync.Once

// GetGlobalPoolManager 获取全局池管理器实例
func GetGlobalPoolManager() *GlobalPoolManager {
	once.Do(func() {
		globalPoolManager = NewGlobalPoolManager()
	})
	return globalPoolManager
}
