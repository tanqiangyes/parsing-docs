package utils

import (
	"fmt"
	"time"
)

// PerformanceMonitor 性能监控器
type PerformanceMonitor struct {
	startTime time.Time
	steps     map[string]time.Duration
}

// NewPerformanceMonitor 创建新的性能监控器
func NewPerformanceMonitor() *PerformanceMonitor {
	return &PerformanceMonitor{
		startTime: time.Now(),
		steps:     make(map[string]time.Duration),
	}
}

// StartStep 开始监控一个步骤
func (pm *PerformanceMonitor) StartStep(stepName string) func() {
	start := time.Now()
	return func() {
		pm.steps[stepName] = time.Since(start)
	}
}

// GetTotalTime 获取总耗时
func (pm *PerformanceMonitor) GetTotalTime() time.Duration {
	return time.Since(pm.startTime)
}

// PrintReport 打印性能报告
func (pm *PerformanceMonitor) PrintReport() {
	fmt.Printf("\n=== 性能监控报告 ===\n")
	fmt.Printf("总耗时: %v\n", pm.GetTotalTime())
	
	if len(pm.steps) > 0 {
		fmt.Printf("各步骤耗时:\n")
		for step, duration := range pm.steps {
			fmt.Printf("  - %s: %v\n", step, duration)
		}
	}
	fmt.Printf("==================\n")
}

// GetStepTime 获取特定步骤的耗时
func (pm *PerformanceMonitor) GetStepTime(stepName string) time.Duration {
	return pm.steps[stepName]
}

// Reset 重置监控器
func (pm *PerformanceMonitor) Reset() {
	pm.startTime = time.Now()
	pm.steps = make(map[string]time.Duration)
} 