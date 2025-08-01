package comparator

import (
	"docs-parser/internal/core/types"
	"fmt"
)

// Comparator 文档对比器接口
type Comparator interface {
	// CompareWithTemplate 与模板进行对比
	CompareWithTemplate(docPath, templatePath string) (*ComparisonReport, error)
	
	// CompareDocuments 对比两个文档
	CompareDocuments(doc1Path, doc2Path string) (*ComparisonReport, error)
	
	// CompareFormatRules 对比格式规则
	CompareFormatRules(docRules, templateRules *types.FormatRules) (*FormatComparison, error)
	
	// CompareContent 对比内容
	CompareContent(docContent, templateContent *types.DocumentContent) (*ContentComparison, error)
	
	// CompareStyles 对比样式
	CompareStyles(docStyles, templateStyles *types.DocumentStyles) (*StyleComparison, error)
}

// ComparisonReport 对比报告
type ComparisonReport struct {
	DocumentPath    string                `json:"document_path"`
	TemplatePath    string                `json:"template_path"`
	OverallScore    float64               `json:"overall_score"`
	ComplianceRate  float64               `json:"compliance_rate"`
	Issues          []FormatIssue         `json:"issues"`
	FormatComparison *FormatComparison    `json:"format_comparison"`
	ContentComparison *ContentComparison  `json:"content_comparison"`
	StyleComparison *StyleComparison      `json:"style_comparison"`
	Recommendations []Recommendation      `json:"recommendations"`
	Summary         ComparisonSummary     `json:"summary"`
}

// FormatIssue 格式问题
type FormatIssue struct {
	ID          string      `json:"id"`
	Type        IssueType   `json:"type"`
	Severity    Severity    `json:"severity"`
	Location    string      `json:"location"`
	Description string      `json:"description"`
	Current     interface{} `json:"current"`
	Expected    interface{} `json:"expected"`
	Rule        string      `json:"rule"`
	Suggestions []string    `json:"suggestions"`
}

// IssueType 问题类型
type IssueType string
const (
	IssueFont        IssueType = "font"
	IssueParagraph   IssueType = "paragraph"
	IssueTable       IssueType = "table"
	IssuePage        IssueType = "page"
	IssueStyle       IssueType = "style"
	IssueContent     IssueType = "content"
	IssueStructure   IssueType = "structure"
)

// Severity 严重程度
type Severity string
const (
	SeverityLow    Severity = "low"
	SeverityMedium Severity = "medium"
	SeverityHigh   Severity = "high"
	SeverityCritical Severity = "critical"
)

// FormatComparison 格式对比
type FormatComparison struct {
	FontRules      []RuleComparison `json:"font_rules"`
	ParagraphRules []RuleComparison `json:"paragraph_rules"`
	TableRules     []RuleComparison `json:"table_rules"`
	PageRules      []RuleComparison `json:"page_rules"`
	StyleRules     []RuleComparison `json:"style_rules"`
	Score          float64          `json:"score"`
	Issues         []FormatIssue    `json:"issues"`
}

// ContentComparison 内容对比
type ContentComparison struct {
	Paragraphs    []ElementComparison `json:"paragraphs"`
	Tables        []ElementComparison `json:"tables"`
	Headers       []ElementComparison `json:"headers"`
	Footers       []ElementComparison `json:"footers"`
	Images        []ElementComparison `json:"images"`
	Score         float64             `json:"score"`
	Issues        []FormatIssue       `json:"issues"`
}

// StyleComparison 样式对比
type StyleComparison struct {
	ParagraphStyles []StyleElementComparison `json:"paragraph_styles"`
	CharacterStyles []StyleElementComparison `json:"character_styles"`
	TableStyles     []StyleElementComparison `json:"table_styles"`
	Score           float64                  `json:"score"`
	Issues          []FormatIssue            `json:"issues"`
}

// RuleComparison 规则对比
type RuleComparison struct {
	RuleID       string      `json:"rule_id"`
	RuleName     string      `json:"rule_name"`
	RuleType     string      `json:"rule_type"`
	Compliant    bool        `json:"compliant"`
	Score        float64     `json:"score"`
	Differences  []Difference `json:"differences"`
	Issues       []FormatIssue `json:"issues"`
}

// ElementComparison 元素对比
type ElementComparison struct {
	ElementID    string      `json:"element_id"`
	ElementType  string      `json:"element_type"`
	Compliant    bool        `json:"compliant"`
	Score        float64     `json:"score"`
	Differences  []Difference `json:"differences"`
	Issues       []FormatIssue `json:"issues"`
}

// StyleElementComparison 样式元素对比
type StyleElementComparison struct {
	StyleID      string      `json:"style_id"`
	StyleName    string      `json:"style_name"`
	StyleType    string      `json:"style_type"`
	Compliant    bool        `json:"compliant"`
	Score        float64     `json:"score"`
	Differences  []Difference `json:"differences"`
	Issues       []FormatIssue `json:"issues"`
}

// Difference 差异
type Difference struct {
	Field        string      `json:"field"`
	Current      interface{} `json:"current"`
	Expected     interface{} `json:"expected"`
	Description  string      `json:"description"`
	Impact       string      `json:"impact"`
}

// Recommendation 建议
type Recommendation struct {
	ID          string   `json:"id"`
	Type        string   `json:"type"`
	Priority    Priority `json:"priority"`
	Description string   `json:"description"`
	Actions     []Action `json:"actions"`
	Impact      string   `json:"impact"`
}

// Priority 优先级
type Priority string
const (
	PriorityLow    Priority = "low"
	PriorityMedium Priority = "medium"
	PriorityHigh   Priority = "high"
	PriorityUrgent Priority = "urgent"
)

// Action 操作
type Action struct {
	Type        string `json:"type"`
	Description string `json:"description"`
	Steps       []Step `json:"steps"`
}

// Step 步骤
type Step struct {
	Order       int    `json:"order"`
	Description string `json:"description"`
	Details     string `json:"details"`
}

// ComparisonSummary 对比摘要
type ComparisonSummary struct {
	TotalIssues       int     `json:"total_issues"`
	CriticalIssues    int     `json:"critical_issues"`
	HighIssues        int     `json:"high_issues"`
	MediumIssues      int     `json:"medium_issues"`
	LowIssues         int     `json:"low_issues"`
	CompliantRules    int     `json:"compliant_rules"`
	NonCompliantRules int     `json:"non_compliant_rules"`
	OverallScore      float64 `json:"overall_score"`
	Recommendations   int     `json:"recommendations"`
}

// ComparatorFactory 对比器工厂
type ComparatorFactory struct {
	comparators map[string]Comparator
}

// NewComparatorFactory 创建对比器工厂
func NewComparatorFactory() *ComparatorFactory {
	return &ComparatorFactory{
		comparators: make(map[string]Comparator),
	}
}

// RegisterComparator 注册对比器
func (cf *ComparatorFactory) RegisterComparator(name string, comparator Comparator) {
	cf.comparators[name] = comparator
}

// GetComparator 获取对比器
func (cf *ComparatorFactory) GetComparator(name string) (Comparator, error) {
	comparator, exists := cf.comparators[name]
	if !exists {
		return nil, ErrComparatorNotFound
	}
	return comparator, nil
}

// 错误定义
var (
	ErrComparatorNotFound = fmt.Errorf("comparator not found")
	ErrTemplateNotFound   = fmt.Errorf("template not found")
	ErrInvalidTemplate    = fmt.Errorf("invalid template")
	ErrComparisonFailed   = fmt.Errorf("comparison failed")
) 