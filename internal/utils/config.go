package utils

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
)

// Config 配置结构
type Config struct {
	// 解析选项
	ParseOptions struct {
		EnablePerformanceMonitoring bool `json:"enable_performance_monitoring"`
		EnableDetailedOutput       bool `json:"enable_detailed_output"`
		MaxParagraphsToShow        int  `json:"max_paragraphs_to_show"`
		EnableStyleParsing         bool `json:"enable_style_parsing"`
		EnableTableParsing         bool `json:"enable_table_parsing"`
		EnableImageParsing         bool `json:"enable_image_parsing"`
	} `json:"parse_options"`

	// 比较选项
	CompareOptions struct {
		StrictMode           bool `json:"strict_mode"`
		IgnoreCase           bool `json:"ignore_case"`
		EnableDetailedReport bool `json:"enable_detailed_report"`
		MaxIssuesToShow      int  `json:"max_issues_to_show"`
	} `json:"compare_options"`

	// 输出选项
	OutputOptions struct {
		EnableColorOutput bool `json:"enable_color_output"`
		EnableJSONOutput  bool `json:"enable_json_output"`
		OutputDirectory   string `json:"output_directory"`
	} `json:"output_options"`

	// 性能选项
	PerformanceOptions struct {
		EnableCaching     bool `json:"enable_caching"`
		CacheSize         int  `json:"cache_size"`
		EnableConcurrency bool `json:"enable_concurrency"`
		MaxWorkers        int  `json:"max_workers"`
	} `json:"performance_options"`
}

// DefaultConfig 返回默认配置
func DefaultConfig() *Config {
	config := &Config{}
	
	// 解析选项默认值
	config.ParseOptions.EnablePerformanceMonitoring = true
	config.ParseOptions.EnableDetailedOutput = true
	config.ParseOptions.MaxParagraphsToShow = 3
	config.ParseOptions.EnableStyleParsing = true
	config.ParseOptions.EnableTableParsing = true
	config.ParseOptions.EnableImageParsing = false

	// 比较选项默认值
	config.CompareOptions.StrictMode = false
	config.CompareOptions.IgnoreCase = true
	config.CompareOptions.EnableDetailedReport = true
	config.CompareOptions.MaxIssuesToShow = 10

	// 输出选项默认值
	config.OutputOptions.EnableColorOutput = true
	config.OutputOptions.EnableJSONOutput = false
	config.OutputOptions.OutputDirectory = "./output"

	// 性能选项默认值
	config.PerformanceOptions.EnableCaching = true
	config.PerformanceOptions.CacheSize = 100
	config.PerformanceOptions.EnableConcurrency = false
	config.PerformanceOptions.MaxWorkers = 4

	return config
}

// LoadConfig 从文件加载配置
func LoadConfig(configPath string) (*Config, error) {
	// 如果配置文件不存在，创建默认配置
	if _, err := os.Stat(configPath); os.IsNotExist(err) {
		config := DefaultConfig()
		if err := SaveConfig(configPath, config); err != nil {
			return nil, fmt.Errorf("failed to create default config: %w", err)
		}
		return config, nil
	}

	// 读取配置文件
	data, err := os.ReadFile(configPath)
	if err != nil {
		return nil, fmt.Errorf("failed to read config file: %w", err)
	}

	var config Config
	if err := json.Unmarshal(data, &config); err != nil {
		return nil, fmt.Errorf("failed to parse config file: %w", err)
	}

	return &config, nil
}

// SaveConfig 保存配置到文件
func SaveConfig(configPath string, config *Config) error {
	// 确保目录存在
	dir := filepath.Dir(configPath)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return fmt.Errorf("failed to create config directory: %w", err)
	}

	// 序列化配置
	data, err := json.MarshalIndent(config, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to marshal config: %w", err)
	}

	// 写入文件
	if err := os.WriteFile(configPath, data, 0644); err != nil {
		return fmt.Errorf("failed to write config file: %w", err)
	}

	return nil
}

// ValidateConfig 验证配置
func ValidateConfig(config *Config) error {
	if config.ParseOptions.MaxParagraphsToShow < 0 {
		return fmt.Errorf("max_paragraphs_to_show must be non-negative")
	}

	if config.CompareOptions.MaxIssuesToShow < 0 {
		return fmt.Errorf("max_issues_to_show must be non-negative")
	}

	if config.PerformanceOptions.CacheSize < 0 {
		return fmt.Errorf("cache_size must be non-negative")
	}

	if config.PerformanceOptions.MaxWorkers < 1 {
		return fmt.Errorf("max_workers must be at least 1")
	}

	return nil
}

// GetConfigPath 获取配置文件路径
func GetConfigPath() string {
	homeDir, err := os.UserHomeDir()
	if err != nil {
		return "./docs-parser-config.json"
	}
	return filepath.Join(homeDir, ".docs-parser", "config.json")
} 