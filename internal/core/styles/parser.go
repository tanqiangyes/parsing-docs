package styles

import (
	"fmt"
	"time"

	"docs-parser/internal/core/types"
)

// Parser 样式解析器接口
type Parser interface {
	// ParseStyles 解析文档样式
	ParseStyles(filePath string) (*types.StyleManager, error)
	
	// ParseStyleInheritance 解析样式继承关系
	ParseStyleInheritance(filePath string) (map[string]types.StyleInheritance, error)
	
	// ParseThemeStyles 解析主题样式
	ParseThemeStyles(filePath string) (map[string]*types.ThemeStyle, error)
	
	// ParseConditionalStyles 解析条件样式
	ParseConditionalStyles(filePath string) ([]types.ConditionalStyle, error)
	
	// ValidateStyles 验证样式
	ValidateStyles(styles *types.StyleManager) (*types.StyleValidation, error)
	
	// ResolveConflicts 解决样式冲突
	ResolveConflicts(styles *types.StyleManager) ([]types.StyleConflict, error)
	
	// GetSupportedStyleTypes 获取支持的样式类型
	GetSupportedStyleTypes() []types.StyleType
}

// DefaultParser 默认样式解析器
type DefaultParser struct{}

// NewParser 创建新的样式解析器
func NewParser() Parser {
	return &DefaultParser{}
}

// ParseStyles 解析文档样式
func (dp *DefaultParser) ParseStyles(filePath string) (*types.StyleManager, error) {
	// 基础实现，返回空的样式管理器
	manager := &types.StyleManager{
		Styles:          make(map[string]*types.AdvancedStyle),
		InheritanceTree: make(map[string][]string),
		ThemeStyles:     make(map[string]*types.ThemeStyle),
		Conflicts:       []types.StyleConflict{},
		Validation: types.StyleValidation{
			Valid:       true,
			Errors:      []string{},
			Warnings:    []string{},
			Suggestions: []string{},
			LastChecked: time.Now(),
		},
	}
	
	return manager, nil
}

// ParseStyleInheritance 解析样式继承关系
func (dp *DefaultParser) ParseStyleInheritance(filePath string) (map[string]types.StyleInheritance, error) {
	// 基础实现，返回空的继承关系
	return make(map[string]types.StyleInheritance), nil
}

// ParseThemeStyles 解析主题样式
func (dp *DefaultParser) ParseThemeStyles(filePath string) (map[string]*types.ThemeStyle, error) {
	// 基础实现，返回空的主题样式
	return make(map[string]*types.ThemeStyle), nil
}

// ParseConditionalStyles 解析条件样式
func (dp *DefaultParser) ParseConditionalStyles(filePath string) ([]types.ConditionalStyle, error) {
	// 基础实现，返回空的条件样式
	return []types.ConditionalStyle{}, nil
}

// ValidateStyles 验证样式
func (dp *DefaultParser) ValidateStyles(styles *types.StyleManager) (*types.StyleValidation, error) {
	validation := &types.StyleValidation{
		Valid:       true,
		Errors:      []string{},
		Warnings:    []string{},
		Suggestions: []string{},
		LastChecked: time.Now(),
	}
	
	// 基础验证逻辑
	if styles == nil {
		validation.Valid = false
		validation.Errors = append(validation.Errors, "样式管理器为空")
		return validation, nil
	}
	
	// 验证样式完整性
	for id, style := range styles.Styles {
		if style == nil {
			validation.Valid = false
			validation.Errors = append(validation.Errors, fmt.Sprintf("样式 %s 为空", id))
			continue
		}
		
		if style.Name == "" {
			validation.Warnings = append(validation.Warnings, fmt.Sprintf("样式 %s 缺少名称", id))
		}
		
		if style.Type == "" {
			validation.Warnings = append(validation.Warnings, fmt.Sprintf("样式 %s 缺少类型", id))
		}
	}
	
	return validation, nil
}

// ResolveConflicts 解决样式冲突
func (dp *DefaultParser) ResolveConflicts(styles *types.StyleManager) ([]types.StyleConflict, error) {
	// 基础实现，返回空的冲突列表
	return []types.StyleConflict{}, nil
}

// GetSupportedStyleTypes 获取支持的样式类型
func (dp *DefaultParser) GetSupportedStyleTypes() []types.StyleType {
	return []types.StyleType{
		types.StyleTypeParagraph,
		types.StyleTypeCharacter,
		types.StyleTypeTable,
		types.StyleTypeList,
		types.StyleTypePage,
		types.StyleTypeSection,
		types.StyleTypeTheme,
		types.StyleTypeCondition,
	}
}

// ParserFactory 样式解析器工厂
type ParserFactory struct {
	parsers map[string]Parser
}

// NewParserFactory 创建新的样式解析器工厂
func NewParserFactory() *ParserFactory {
	factory := &ParserFactory{
		parsers: make(map[string]Parser),
	}
	
	// 注册默认解析器
	factory.RegisterParser("default", NewParser())
	
	return factory
}

// RegisterParser 注册样式解析器
func (pf *ParserFactory) RegisterParser(name string, parser Parser) {
	pf.parsers[name] = parser
}

// GetParser 获取样式解析器
func (pf *ParserFactory) GetParser(name string) (Parser, error) {
	parser, exists := pf.parsers[name]
	if !exists {
		return nil, fmt.Errorf("样式解析器 %s 不存在", name)
	}
	return parser, nil
}

// GetDefaultParser 获取默认样式解析器
func (pf *ParserFactory) GetDefaultParser() Parser {
	parser, err := pf.GetParser("default")
	if err != nil {
		// 如果默认解析器不存在，创建一个新的
		return NewParser()
	}
	return parser
}

// 错误定义
var (
	ErrUnsupportedStyleType = fmt.Errorf("不支持的样式类型")
	ErrInvalidStyleData     = fmt.Errorf("无效的样式数据")
	ErrStyleConflict        = fmt.Errorf("样式冲突")
	ErrStyleValidation      = fmt.Errorf("样式验证失败")
) 