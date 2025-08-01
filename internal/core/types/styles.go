package types

import (
	"time"
)

// StyleType 样式类型
type StyleType string

const (
	StyleTypeParagraph StyleType = "paragraph" // 段落样式
	StyleTypeCharacter StyleType = "character" // 字符样式
	StyleTypeTable     StyleType = "table"     // 表格样式
	StyleTypeList      StyleType = "list"      // 列表样式
	StyleTypePage      StyleType = "page"      // 页面样式
	StyleTypeSection   StyleType = "section"   // 节样式
	StyleTypeTheme     StyleType = "theme"     // 主题样式
	StyleTypeCondition StyleType = "condition" // 条件样式
)

// StyleInheritance 样式继承关系
type StyleInheritance struct {
	BasedOn     string   `json:"based_on"`      // 基于样式
	Next        string   `json:"next"`          // 下一样式
	Linked      string   `json:"linked"`        // 链接样式
	Parent      string   `json:"parent"`        // 父样式
	Children    []string `json:"children"`      // 子样式
	Priority    int      `json:"priority"`      // 优先级
	Hidden      bool     `json:"hidden"`        // 是否隐藏
	QuickFormat bool     `json:"quick_format"`  // 是否快速格式
}

// ThemeStyle 主题样式
type ThemeStyle struct {
	ThemeName   string `json:"theme_name"`    // 主题名称
	ColorScheme string `json:"color_scheme"`  // 颜色方案
	FontScheme  string `json:"font_scheme"`   // 字体方案
	Effects     string `json:"effects"`       // 效果方案
	Version     string `json:"version"`       // 主题版本
}

// ConditionalStyle 条件样式
type ConditionalStyle struct {
	ID          string `json:"id"`
	Condition   string `json:"condition"`     // 条件表达式
	Style       AdvancedStyle `json:"style"`  // 应用样式
	Priority    int    `json:"priority"`      // 优先级
	Active      bool   `json:"active"`        // 是否激活
	Description string `json:"description"`   // 条件描述
}

// StyleConflict 样式冲突
type StyleConflict struct {
	ID          string   `json:"id"`
	Conflicting []string `json:"conflicting"` // 冲突的样式ID
	Resolution  string   `json:"resolution"`  // 解决方案
	Priority    int      `json:"priority"`    // 优先级
	Resolved    bool     `json:"resolved"`    // 是否已解决
}

// StyleValidation 样式验证
type StyleValidation struct {
	Valid       bool     `json:"valid"`        // 是否有效
	Errors      []string `json:"errors"`       // 错误信息
	Warnings    []string `json:"warnings"`     // 警告信息
	Suggestions []string `json:"suggestions"`  // 建议
	LastChecked time.Time `json:"last_checked"` // 最后检查时间
}

// StyleProperties 样式属性
type StyleProperties struct {
	// 基础属性
	Name        string  `json:"name"`
	Description string  `json:"description"`
	Category    string  `json:"category"`
	
	// 字体属性
	Font        Font    `json:"font"`
	Size        float64 `json:"size"`
	Color       Color   `json:"color"`
	Bold        bool    `json:"bold"`
	Italic      bool    `json:"italic"`
	Underline   Underline `json:"underline"`
	Highlight   Highlight `json:"highlight"`
	
	// 段落属性
	Alignment   Alignment   `json:"alignment"`
	Indentation Indentation `json:"indentation"`
	Spacing     Spacing     `json:"spacing"`
	Borders     Borders     `json:"borders"`
	Shading     Shading     `json:"shading"`
	
	// 高级属性
	Position    Position    `json:"position"`
	Rotation    float64     `json:"rotation"`
	Scale       float64     `json:"scale"`
	Opacity     float64     `json:"opacity"`
	Effects     []Effect    `json:"effects"`
	
	// 列表属性
	ListType    ListType    `json:"list_type"`
	ListLevel   int         `json:"list_level"`
	Numbering   Numbering   `json:"numbering"`
	
	// 表格属性
	TableBorders TableBorders `json:"table_borders"`
	TableShading TableShading `json:"table_shading"`
	CellPadding  CellPadding  `json:"cell_padding"`
	
	// 页面属性
	PageSize    PageSize    `json:"page_size"`
	PageMargins PageMargins `json:"page_margins"`
	Columns     Columns     `json:"columns"`
	
	// 节属性
	SectionType SectionType `json:"section_type"`
	HeaderFooter HeaderFooter `json:"header_footer"`
}

// Effect 效果
type Effect struct {
	Type        string  `json:"type"`
	Value       string  `json:"value"`
	Intensity   float64 `json:"intensity"`
	Color       Color   `json:"color"`
	Direction   string  `json:"direction"`
}

// ListType 列表类型
type ListType string

const (
	ListTypeNone     ListType = "none"
	ListTypeBullet   ListType = "bullet"
	ListTypeNumber   ListType = "number"
	ListTypeOutline  ListType = "outline"
	ListTypeCustom   ListType = "custom"
)

// Numbering 编号
type Numbering struct {
	Type        string `json:"type"`
	Format      string `json:"format"`
	Start       int    `json:"start"`
	Increment   int    `json:"increment"`
	Restart     bool   `json:"restart"`
	Level       int    `json:"level"`
	Text        string `json:"text"`
	Alignment   string `json:"alignment"`
}

// CellPadding 单元格内边距
type CellPadding struct {
	Top    float64 `json:"top"`
	Bottom float64 `json:"bottom"`
	Left   float64 `json:"left"`
	Right  float64 `json:"right"`
}

// SectionType 节类型
type SectionType string

const (
	SectionTypeContinuous SectionType = "continuous"
	SectionTypeNewColumn  SectionType = "new_column"
	SectionTypeNewPage    SectionType = "new_page"
	SectionTypeEvenPage   SectionType = "even_page"
	SectionTypeOddPage    SectionType = "odd_page"
)

// HeaderFooter 页眉页脚
type HeaderFooter struct {
	Header      bool   `json:"header"`
	Footer      bool   `json:"footer"`
	FirstPage   bool   `json:"first_page"`
	EvenPage    bool   `json:"even_page"`
	OddPage     bool   `json:"odd_page"`
	Different   bool   `json:"different"`
	LinkToPrev  bool   `json:"link_to_prev"`
}

// AdvancedStyle 高级样式
type AdvancedStyle struct {
	ID              string            `json:"id"`
	Name            string            `json:"name"`
	Type            StyleType         `json:"type"`
	Properties      StyleProperties   `json:"properties"`
	Inheritance     StyleInheritance `json:"inheritance"`
	Theme           ThemeStyle        `json:"theme"`
	Conditions      []ConditionalStyle `json:"conditions"`
	Conflicts       []StyleConflict   `json:"conflicts"`
	Validation      StyleValidation   `json:"validation"`
	Created         time.Time         `json:"created"`
	Modified        time.Time         `json:"modified"`
	Version         string            `json:"version"`
}

// StyleManager 样式管理器
type StyleManager struct {
	Styles          map[string]*AdvancedStyle `json:"styles"`
	InheritanceTree map[string][]string      `json:"inheritance_tree"`
	ThemeStyles     map[string]*ThemeStyle   `json:"theme_styles"`
	Conflicts       []StyleConflict          `json:"conflicts"`
	Validation      StyleValidation          `json:"validation"`
}

// StyleParser 样式解析器接口
type StyleParser interface {
	// ParseStyles 解析文档样式
	ParseStyles(filePath string) (*StyleManager, error)
	
	// ParseStyleInheritance 解析样式继承关系
	ParseStyleInheritance(filePath string) (map[string]StyleInheritance, error)
	
	// ParseThemeStyles 解析主题样式
	ParseThemeStyles(filePath string) (map[string]*ThemeStyle, error)
	
	// ParseConditionalStyles 解析条件样式
	ParseConditionalStyles(filePath string) ([]ConditionalStyle, error)
	
	// ValidateStyles 验证样式
	ValidateStyles(styles *StyleManager) (*StyleValidation, error)
	
	// ResolveConflicts 解决样式冲突
	ResolveConflicts(styles *StyleManager) ([]StyleConflict, error)
	
	// GetSupportedStyleTypes 获取支持的样式类型
	GetSupportedStyleTypes() []StyleType
}

// StyleComparator 样式比较器接口
type StyleComparator interface {
	// CompareStyles 比较样式
	CompareStyles(style1, style2 *AdvancedStyle) (*StyleComparison, error)
	
	// CompareStyleManagers 比较样式管理器
	CompareStyleManagers(manager1, manager2 *StyleManager) (*StyleManagerComparison, error)
	
	// FindStyleDifferences 查找样式差异
	FindStyleDifferences(manager1, manager2 *StyleManager) ([]StyleDifference, error)
}

// StyleComparison 样式比较结果
type StyleComparison struct {
	ID              string            `json:"id"`
	Differences     []StyleDifference `json:"differences"`
	Similarity      float64           `json:"similarity"`
	Compatibility   bool              `json:"compatibility"`
	Recommendations []string          `json:"recommendations"`
}

// StyleManagerComparison 样式管理器比较结果
type StyleManagerComparison struct {
	TotalStyles     int                `json:"total_styles"`
	MatchingStyles  int                `json:"matching_styles"`
	DifferentStyles int                `json:"different_styles"`
	MissingStyles   int                `json:"missing_styles"`
	ExtraStyles     int                `json:"extra_styles"`
	Comparisons     []StyleComparison  `json:"comparisons"`
	OverallSimilarity float64          `json:"overall_similarity"`
}

// StyleDifference 样式差异
type StyleDifference struct {
	ID          string `json:"id"`
	Property    string `json:"property"`
	Value1      string `json:"value1"`
	Value2      string `json:"value2"`
	Type        string `json:"type"`
	Severity    string `json:"severity"`
	Description string `json:"description"`
	Fix         string `json:"fix"`
} 