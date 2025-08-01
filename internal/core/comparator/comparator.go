package comparator

import (
	"fmt"

	"docs-parser/internal/core/types"
	"docs-parser/internal/formats"
	"docs-parser/internal/templates"
	"docs-parser/internal/utils"
)

// DocumentComparator 文档比较器实现
type DocumentComparator struct {
	wordParser      *formats.WordParser
	templateManager *templates.TemplateManager
}

// NewDocumentComparator 创建文档比较器
func NewDocumentComparator() *DocumentComparator {
	return &DocumentComparator{
		wordParser:      formats.NewWordParser(),
		templateManager: templates.NewTemplateManager(""),
	}
}

// CompareWithTemplate 与Word文档模板进行对比
func (dc *DocumentComparator) CompareWithTemplate(docPath, templatePath string) (*ComparisonReport, error) {
	// 验证文档文件
	if err := utils.ValidateFile(docPath); err != nil {
		return nil, fmt.Errorf("document validation failed: %w", err)
	}

	// 验证模板文件
	if err := utils.ValidateFile(templatePath); err != nil {
		return nil, fmt.Errorf("template validation failed: %w", err)
	}

	// 解析文档
	doc, err := dc.wordParser.ParseDocument(docPath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse document: %w", err)
	}

	// 解析模板文档
	templateDoc, err := dc.wordParser.ParseDocument(templatePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse template document: %w", err)
	}

	// 比较格式规则
	formatComparison, err := dc.CompareFormatRules(&doc.FormatRules, &templateDoc.FormatRules)
	if err != nil {
		return nil, fmt.Errorf("format comparison failed: %w", err)
	}

	// 比较内容
	contentComparison, err := dc.CompareContent(&doc.Content, &templateDoc.Content)
	if err != nil {
		return nil, fmt.Errorf("content comparison failed: %w", err)
	}

	// 比较样式
	styleComparison, err := dc.CompareStyles(&doc.Styles, &templateDoc.Styles)
	if err != nil {
		return nil, fmt.Errorf("style comparison failed: %w", err)
	}

	// 生成建议
	recommendations := dc.generateRecommendations(formatComparison, contentComparison, styleComparison)

	// 生成摘要
	summary := dc.generateSummary(formatComparison, contentComparison, styleComparison)

	// 计算总体分数
	overallScore := (formatComparison.Score + contentComparison.Score + styleComparison.Score) / 3.0
	complianceRate := dc.calculateComplianceRate(formatComparison, contentComparison, styleComparison)

	// 收集所有问题
	var allIssues []FormatIssue
	allIssues = append(allIssues, formatComparison.Issues...)
	allIssues = append(allIssues, contentComparison.Issues...)
	allIssues = append(allIssues, styleComparison.Issues...)

	return &ComparisonReport{
		DocumentPath:      docPath,
		TemplatePath:      templatePath,
		OverallScore:      overallScore,
		ComplianceRate:    complianceRate,
		Issues:            allIssues,
		FormatComparison:  formatComparison,
		ContentComparison: contentComparison,
		StyleComparison:   styleComparison,
		Recommendations:   recommendations,
		Summary:           summary,
	}, nil
}

// CompareDocuments 对比两个文档
func (dc *DocumentComparator) CompareDocuments(doc1Path, doc2Path string) (*ComparisonReport, error) {
	// 解析两个文档
	doc1, err := dc.wordParser.ParseDocument(doc1Path)
	if err != nil {
		return nil, fmt.Errorf("failed to parse first document: %w", err)
	}

	doc2, err := dc.wordParser.ParseDocument(doc2Path)
	if err != nil {
		return nil, fmt.Errorf("failed to parse second document: %w", err)
	}

	// 比较格式规则
	formatComparison, err := dc.CompareFormatRules(&doc1.FormatRules, &doc2.FormatRules)
	if err != nil {
		return nil, fmt.Errorf("format comparison failed: %w", err)
	}

	// 比较内容
	contentComparison, err := dc.CompareContent(&doc1.Content, &doc2.Content)
	if err != nil {
		return nil, fmt.Errorf("content comparison failed: %w", err)
	}

	// 比较样式
	styleComparison, err := dc.CompareStyles(&doc1.Styles, &doc2.Styles)
	if err != nil {
		return nil, fmt.Errorf("style comparison failed: %w", err)
	}

	// 生成建议
	recommendations := dc.generateRecommendations(formatComparison, contentComparison, styleComparison)

	// 生成摘要
	summary := dc.generateSummary(formatComparison, contentComparison, styleComparison)

	// 计算总体分数
	overallScore := (formatComparison.Score + contentComparison.Score + styleComparison.Score) / 3.0
	complianceRate := dc.calculateComplianceRate(formatComparison, contentComparison, styleComparison)

	// 收集所有问题
	var allIssues []FormatIssue
	allIssues = append(allIssues, formatComparison.Issues...)
	allIssues = append(allIssues, contentComparison.Issues...)
	allIssues = append(allIssues, styleComparison.Issues...)

	return &ComparisonReport{
		DocumentPath:      doc1Path,
		TemplatePath:      doc2Path,
		OverallScore:      overallScore,
		ComplianceRate:    complianceRate,
		Issues:            allIssues,
		FormatComparison:  formatComparison,
		ContentComparison: contentComparison,
		StyleComparison:   styleComparison,
		Recommendations:   recommendations,
		Summary:           summary,
	}, nil
}

// CompareFormatRules 对比格式规则
func (dc *DocumentComparator) CompareFormatRules(docRules, templateRules *types.FormatRules) (*FormatComparison, error) {
	comparison := &FormatComparison{
		FontRules:      []RuleComparison{},
		ParagraphRules: []RuleComparison{},
		TableRules:     []RuleComparison{},
		PageRules:      []RuleComparison{},
		StyleRules:     []RuleComparison{},
		Issues:         []FormatIssue{},
	}

	// 比较字体规则
	fontComparisons := dc.compareFontRules(docRules.FontRules, templateRules.FontRules)
	comparison.FontRules = fontComparisons

	// 比较段落规则
	paragraphComparisons := dc.compareParagraphRules(docRules.ParagraphRules, templateRules.ParagraphRules)
	comparison.ParagraphRules = paragraphComparisons

	// 比较表格规则
	tableComparisons := dc.compareTableRules(docRules.TableRules, templateRules.TableRules)
	comparison.TableRules = tableComparisons

	// 比较页面规则
	pageComparisons := dc.comparePageRules(docRules.PageRules, templateRules.PageRules)
	comparison.PageRules = pageComparisons

	// 计算分数
	comparison.Score = dc.calculateFormatScore(comparison)

	// 收集问题
	comparison.Issues = dc.collectFormatIssues(comparison)

	return comparison, nil
}

// CompareContent 对比内容
func (dc *DocumentComparator) CompareContent(docContent, templateContent *types.DocumentContent) (*ContentComparison, error) {
	comparison := &ContentComparison{
		Paragraphs: []ElementComparison{},
		Tables:     []ElementComparison{},
		Headers:    []ElementComparison{},
		Footers:    []ElementComparison{},
		Images:     []ElementComparison{},
		Issues:     []FormatIssue{},
	}

	// 比较段落
	paragraphComparisons := dc.compareParagraphs(docContent.Paragraphs, templateContent.Paragraphs)
	comparison.Paragraphs = paragraphComparisons

	// 比较表格
	tableComparisons := dc.compareTables(docContent.Tables, templateContent.Tables)
	comparison.Tables = tableComparisons

	// 计算分数
	comparison.Score = dc.calculateContentScore(comparison)

	// 收集问题
	comparison.Issues = dc.collectContentIssues(comparison)

	return comparison, nil
}

// CompareStyles 对比样式
func (dc *DocumentComparator) CompareStyles(docStyles, templateStyles *types.DocumentStyles) (*StyleComparison, error) {
	comparison := &StyleComparison{
		ParagraphStyles: []StyleElementComparison{},
		CharacterStyles: []StyleElementComparison{},
		TableStyles:     []StyleElementComparison{},
		Issues:          []FormatIssue{},
	}

	// 比较段落样式
	paragraphStyleComparisons := dc.compareParagraphStyles(docStyles.ParagraphStyles, templateStyles.ParagraphStyles)
	comparison.ParagraphStyles = paragraphStyleComparisons

	// 比较字符样式
	characterStyleComparisons := dc.compareCharacterStyles(docStyles.CharacterStyles, templateStyles.CharacterStyles)
	comparison.CharacterStyles = characterStyleComparisons

	// 比较表格样式
	tableStyleComparisons := dc.compareTableStyles(docStyles.TableStyles, templateStyles.TableStyles)
	comparison.TableStyles = tableStyleComparisons

	// 计算分数
	comparison.Score = dc.calculateStyleScore(comparison)

	// 收集问题
	comparison.Issues = dc.collectStyleIssues(comparison)

	return comparison, nil
}

// 辅助方法实现
func (dc *DocumentComparator) compareFontRules(docFonts, templateFonts []types.FontRule) []RuleComparison {
	var comparisons []RuleComparison

	for _, templateFont := range templateFonts {
		comparison := RuleComparison{
			RuleID:   templateFont.ID,
			RuleName: templateFont.Name,
			RuleType: "font",
		}

		// 查找匹配的文档字体
		found := false
		for _, docFont := range docFonts {
			if docFont.ID == templateFont.ID {
				found = true
				comparison.Compliant = dc.isFontCompliant(docFont, templateFont)
				comparison.Score = dc.calculateFontScore(docFont, templateFont)
				comparison.Differences = dc.getFontDifferences(docFont, templateFont)
				break
			}
		}

		if !found {
			comparison.Compliant = false
			comparison.Score = 0.0
			comparison.Issues = []FormatIssue{{
				ID:          fmt.Sprintf("font_missing_%s", templateFont.ID),
				Type:        IssueFont,
				Severity:    SeverityHigh,
				Location:    "document",
				Description: fmt.Sprintf("Missing required font: %s", templateFont.Name),
				Expected:    templateFont,
				Current:     nil,
			}}
		}

		comparisons = append(comparisons, comparison)
	}

	return comparisons
}

func (dc *DocumentComparator) compareParagraphRules(docParagraphs, templateParagraphs []types.ParagraphRule) []RuleComparison {
	var comparisons []RuleComparison

	for _, templatePara := range templateParagraphs {
		comparison := RuleComparison{
			RuleID:   templatePara.ID,
			RuleName: templatePara.Name,
			RuleType: "paragraph",
		}

		// 查找匹配的文档段落规则
		found := false
		for _, docPara := range docParagraphs {
			if docPara.ID == templatePara.ID {
				found = true
				comparison.Compliant = dc.isParagraphCompliant(docPara, templatePara)
				comparison.Score = dc.calculateParagraphScore(docPara, templatePara)
				comparison.Differences = dc.getParagraphDifferences(docPara, templatePara)
				break
			}
		}

		if !found {
			comparison.Compliant = false
			comparison.Score = 0.0
			comparison.Issues = []FormatIssue{{
				ID:          fmt.Sprintf("paragraph_missing_%s", templatePara.ID),
				Type:        IssueParagraph,
				Severity:    SeverityHigh,
				Location:    "document",
				Description: fmt.Sprintf("Missing required paragraph rule: %s", templatePara.Name),
				Expected:    templatePara,
				Current:     nil,
			}}
		}

		comparisons = append(comparisons, comparison)
	}

	return comparisons
}

func (dc *DocumentComparator) compareTableRules(docTables, templateTables []types.TableRule) []RuleComparison {
	var comparisons []RuleComparison

	for _, templateTable := range templateTables {
		comparison := RuleComparison{
			RuleID:   templateTable.ID,
			RuleName: templateTable.Name,
			RuleType: "table",
		}

		// 查找匹配的文档表格规则
		found := false
		for _, docTable := range docTables {
			if docTable.ID == templateTable.ID {
				found = true
				comparison.Compliant = dc.isTableCompliant(docTable, templateTable)
				comparison.Score = dc.calculateTableScore(docTable, templateTable)
				comparison.Differences = dc.getTableDifferences(docTable, templateTable)
				break
			}
		}

		if !found {
			comparison.Compliant = false
			comparison.Score = 0.0
			comparison.Issues = []FormatIssue{{
				ID:          fmt.Sprintf("table_missing_%s", templateTable.ID),
				Type:        IssueTable,
				Severity:    SeverityHigh,
				Location:    "document",
				Description: fmt.Sprintf("Missing required table rule: %s", templateTable.Name),
				Expected:    templateTable,
				Current:     nil,
			}}
		}

		comparisons = append(comparisons, comparison)
	}

	return comparisons
}

func (dc *DocumentComparator) comparePageRules(docPages, templatePages []types.PageRule) []RuleComparison {
	var comparisons []RuleComparison

	for _, templatePage := range templatePages {
		comparison := RuleComparison{
			RuleID:   templatePage.ID,
			RuleName: templatePage.Name,
			RuleType: "page",
		}

		// 查找匹配的文档页面规则
		found := false
		for _, docPage := range docPages {
			if docPage.ID == templatePage.ID {
				found = true
				comparison.Compliant = dc.isPageCompliant(docPage, templatePage)
				comparison.Score = dc.calculatePageScore(docPage, templatePage)
				comparison.Differences = dc.getPageDifferences(docPage, templatePage)
				break
			}
		}

		if !found {
			comparison.Compliant = false
			comparison.Score = 0.0
			comparison.Issues = []FormatIssue{{
				ID:          fmt.Sprintf("page_missing_%s", templatePage.ID),
				Type:        IssuePage,
				Severity:    SeverityHigh,
				Location:    "document",
				Description: fmt.Sprintf("Missing required page rule: %s", templatePage.Name),
				Expected:    templatePage,
				Current:     nil,
			}}
		}

		comparisons = append(comparisons, comparison)
	}

	return comparisons
}

// 比较方法实现
func (dc *DocumentComparator) compareParagraphs(docParagraphs, templateParagraphs []types.Paragraph) []ElementComparison {
	var comparisons []ElementComparison

	for _, docPara := range docParagraphs {
		comparison := ElementComparison{
			ElementID:   docPara.ID,
			ElementType: "paragraph",
		}

		// 查找匹配的模板段落
		found := false
		for _, templatePara := range templateParagraphs {
			if templatePara.ID == docPara.ID {
				found = true
				comparison.Compliant = dc.isParagraphElementCompliant(docPara, templatePara)
				comparison.Score = dc.calculateParagraphElementScore(docPara, templatePara)
				comparison.Differences = dc.getParagraphElementDifferences(docPara, templatePara)
				break
			}
		}

		if !found {
			comparison.Compliant = false
			comparison.Score = 0.0
		}

		comparisons = append(comparisons, comparison)
	}

	return comparisons
}

func (dc *DocumentComparator) compareTables(docTables, templateTables []types.Table) []ElementComparison {
	var comparisons []ElementComparison

	for _, docTable := range docTables {
		comparison := ElementComparison{
			ElementID:   docTable.ID,
			ElementType: "table",
		}

		// 查找匹配的模板表格
		found := false
		for _, templateTable := range templateTables {
			if templateTable.ID == docTable.ID {
				found = true
				comparison.Compliant = dc.isTableElementCompliant(docTable, templateTable)
				comparison.Score = dc.calculateTableElementScore(docTable, templateTable)
				comparison.Differences = dc.getTableElementDifferences(docTable, templateTable)
				break
			}
		}

		if !found {
			comparison.Compliant = false
			comparison.Score = 0.0
		}

		comparisons = append(comparisons, comparison)
	}

	return comparisons
}

// 样式比较方法
func (dc *DocumentComparator) compareParagraphStyles(docStyles, templateStyles []types.ParagraphStyle) []StyleElementComparison {
	var comparisons []StyleElementComparison

	for _, templateStyle := range templateStyles {
		comparison := StyleElementComparison{
			StyleID:   templateStyle.ID,
			StyleName: templateStyle.Name,
			StyleType: "paragraph",
		}

		// 查找匹配的文档样式
		found := false
		for _, docStyle := range docStyles {
			if docStyle.ID == templateStyle.ID {
				found = true
				comparison.Compliant = dc.isParagraphStyleCompliant(docStyle, templateStyle)
				comparison.Score = dc.calculateParagraphStyleScore(docStyle, templateStyle)
				comparison.Differences = dc.getParagraphStyleDifferences(docStyle, templateStyle)
				break
			}
		}

		if !found {
			comparison.Compliant = false
			comparison.Score = 0.0
		}

		comparisons = append(comparisons, comparison)
	}

	return comparisons
}

func (dc *DocumentComparator) compareCharacterStyles(docStyles, templateStyles []types.CharacterStyle) []StyleElementComparison {
	var comparisons []StyleElementComparison

	for _, templateStyle := range templateStyles {
		comparison := StyleElementComparison{
			StyleID:   templateStyle.ID,
			StyleName: templateStyle.Name,
			StyleType: "character",
		}

		// 查找匹配的文档样式
		found := false
		for _, docStyle := range docStyles {
			if docStyle.ID == templateStyle.ID {
				found = true
				comparison.Compliant = dc.isCharacterStyleCompliant(docStyle, templateStyle)
				comparison.Score = dc.calculateCharacterStyleScore(docStyle, templateStyle)
				comparison.Differences = dc.getCharacterStyleDifferences(docStyle, templateStyle)
				break
			}
		}

		if !found {
			comparison.Compliant = false
			comparison.Score = 0.0
		}

		comparisons = append(comparisons, comparison)
	}

	return comparisons
}

func (dc *DocumentComparator) compareTableStyles(docStyles, templateStyles []types.TableStyle) []StyleElementComparison {
	var comparisons []StyleElementComparison

	for _, templateStyle := range templateStyles {
		comparison := StyleElementComparison{
			StyleID:   templateStyle.ID,
			StyleName: templateStyle.Name,
			StyleType: "table",
		}

		// 查找匹配的文档样式
		found := false
		for _, docStyle := range docStyles {
			if docStyle.ID == templateStyle.ID {
				found = true
				comparison.Compliant = dc.isTableStyleCompliant(docStyle, templateStyle)
				comparison.Score = dc.calculateTableStyleScore(docStyle, templateStyle)
				comparison.Differences = dc.getTableStyleDifferences(docStyle, templateStyle)
				break
			}
		}

		if !found {
			comparison.Compliant = false
			comparison.Score = 0.0
		}

		comparisons = append(comparisons, comparison)
	}

	return comparisons
}

// 合规性检查方法
func (dc *DocumentComparator) isFontCompliant(docFont, templateFont types.FontRule) bool {
	return docFont.Name == templateFont.Name && docFont.Size == templateFont.Size
}

func (dc *DocumentComparator) isParagraphCompliant(docPara, templatePara types.ParagraphRule) bool {
	return docPara.Name == templatePara.Name && docPara.Alignment == templatePara.Alignment
}

func (dc *DocumentComparator) isTableCompliant(docTable, templateTable types.TableRule) bool {
	return docTable.Name == templateTable.Name && docTable.Width == templateTable.Width
}

func (dc *DocumentComparator) isPageCompliant(docPage, templatePage types.PageRule) bool {
	return docPage.Name == templatePage.Name &&
		docPage.PageSize.Width == templatePage.PageSize.Width &&
		docPage.PageSize.Height == templatePage.PageSize.Height
}

// 分数计算方法
func (dc *DocumentComparator) calculateFontScore(docFont, templateFont types.FontRule) float64 {
	score := 0.0
	if docFont.Name == templateFont.Name {
		score += 50.0
	}
	if docFont.Size == templateFont.Size {
		score += 50.0
	}
	return score
}

func (dc *DocumentComparator) calculateParagraphScore(docPara, templatePara types.ParagraphRule) float64 {
	score := 0.0
	if docPara.Name == templatePara.Name {
		score += 50.0
	}
	if docPara.Alignment == templatePara.Alignment {
		score += 50.0
	}
	return score
}

func (dc *DocumentComparator) calculateTableScore(docTable, templateTable types.TableRule) float64 {
	score := 0.0
	if docTable.Name == templateTable.Name {
		score += 50.0
	}
	if docTable.Width == templateTable.Width {
		score += 50.0
	}
	return score
}

func (dc *DocumentComparator) calculatePageScore(docPage, templatePage types.PageRule) float64 {
	score := 0.0
	if docPage.Name == templatePage.Name {
		score += 25.0
	}
	if docPage.PageSize.Width == templatePage.PageSize.Width {
		score += 37.5
	}
	if docPage.PageSize.Height == templatePage.PageSize.Height {
		score += 37.5
	}
	return score
}

// 差异获取方法
func (dc *DocumentComparator) getFontDifferences(docFont, templateFont types.FontRule) []Difference {
	var differences []Difference

	if docFont.Name != templateFont.Name {
		differences = append(differences, Difference{
			Field:       "name",
			Current:     docFont.Name,
			Expected:    templateFont.Name,
			Description: "Font name mismatch",
			Impact:      "medium",
		})
	}

	if docFont.Size != templateFont.Size {
		differences = append(differences, Difference{
			Field:       "size",
			Current:     docFont.Size,
			Expected:    templateFont.Size,
			Description: "Font size mismatch",
			Impact:      "medium",
		})
	}

	return differences
}

func (dc *DocumentComparator) getParagraphDifferences(docPara, templatePara types.ParagraphRule) []Difference {
	var differences []Difference

	if docPara.Name != templatePara.Name {
		differences = append(differences, Difference{
			Field:       "name",
			Current:     docPara.Name,
			Expected:    templatePara.Name,
			Description: "Paragraph rule name mismatch",
			Impact:      "medium",
		})
	}

	if docPara.Alignment != templatePara.Alignment {
		differences = append(differences, Difference{
			Field:       "alignment",
			Current:     docPara.Alignment,
			Expected:    templatePara.Alignment,
			Description: "Paragraph alignment mismatch",
			Impact:      "medium",
		})
	}

	return differences
}

func (dc *DocumentComparator) getTableDifferences(docTable, templateTable types.TableRule) []Difference {
	var differences []Difference

	if docTable.Name != templateTable.Name {
		differences = append(differences, Difference{
			Field:       "name",
			Current:     docTable.Name,
			Expected:    templateTable.Name,
			Description: "Table rule name mismatch",
			Impact:      "medium",
		})
	}

	if docTable.Width != templateTable.Width {
		differences = append(differences, Difference{
			Field:       "width",
			Current:     docTable.Width,
			Expected:    templateTable.Width,
			Description: "Table width mismatch",
			Impact:      "medium",
		})
	}

	return differences
}

func (dc *DocumentComparator) getPageDifferences(docPage, templatePage types.PageRule) []Difference {
	var differences []Difference

	if docPage.Name != templatePage.Name {
		differences = append(differences, Difference{
			Field:       "name",
			Current:     docPage.Name,
			Expected:    templatePage.Name,
			Description: "Page rule name mismatch",
			Impact:      "medium",
		})
	}

	if docPage.PageSize.Width != templatePage.PageSize.Width {
		differences = append(differences, Difference{
			Field:       "page_width",
			Current:     docPage.PageSize.Width,
			Expected:    templatePage.PageSize.Width,
			Description: "Page width mismatch",
			Impact:      "medium",
		})
	}

	if docPage.PageSize.Height != templatePage.PageSize.Height {
		differences = append(differences, Difference{
			Field:       "page_height",
			Current:     docPage.PageSize.Height,
			Expected:    templatePage.PageSize.Height,
			Description: "Page height mismatch",
			Impact:      "medium",
		})
	}

	return differences
}

// 分数计算方法
func (dc *DocumentComparator) calculateFormatScore(comparison *FormatComparison) float64 {
	if len(comparison.FontRules) == 0 && len(comparison.ParagraphRules) == 0 &&
		len(comparison.TableRules) == 0 && len(comparison.PageRules) == 0 {
		return 100.0
	}

	totalScore := 0.0
	totalRules := 0

	for _, rule := range comparison.FontRules {
		totalScore += rule.Score
		totalRules++
	}

	for _, rule := range comparison.ParagraphRules {
		totalScore += rule.Score
		totalRules++
	}

	for _, rule := range comparison.TableRules {
		totalScore += rule.Score
		totalRules++
	}

	for _, rule := range comparison.PageRules {
		totalScore += rule.Score
		totalRules++
	}

	if totalRules == 0 {
		return 100.0
	}

	return totalScore / float64(totalRules)
}

func (dc *DocumentComparator) calculateContentScore(comparison *ContentComparison) float64 {
	if len(comparison.Paragraphs) == 0 && len(comparison.Tables) == 0 {
		return 100.0
	}

	totalScore := 0.0
	totalElements := 0

	for _, element := range comparison.Paragraphs {
		totalScore += element.Score
		totalElements++
	}

	for _, element := range comparison.Tables {
		totalScore += element.Score
		totalElements++
	}

	if totalElements == 0 {
		return 100.0
	}

	return totalScore / float64(totalElements)
}

func (dc *DocumentComparator) calculateStyleScore(comparison *StyleComparison) float64 {
	if len(comparison.ParagraphStyles) == 0 && len(comparison.CharacterStyles) == 0 &&
		len(comparison.TableStyles) == 0 {
		return 100.0
	}

	totalScore := 0.0
	totalStyles := 0

	for _, style := range comparison.ParagraphStyles {
		totalScore += style.Score
		totalStyles++
	}

	for _, style := range comparison.CharacterStyles {
		totalScore += style.Score
		totalStyles++
	}

	for _, style := range comparison.TableStyles {
		totalScore += style.Score
		totalStyles++
	}

	if totalStyles == 0 {
		return 100.0
	}

	return totalScore / float64(totalStyles)
}

// 问题收集方法
func (dc *DocumentComparator) collectFormatIssues(comparison *FormatComparison) []FormatIssue {
	var issues []FormatIssue

	for _, rule := range comparison.FontRules {
		issues = append(issues, rule.Issues...)
	}

	for _, rule := range comparison.ParagraphRules {
		issues = append(issues, rule.Issues...)
	}

	for _, rule := range comparison.TableRules {
		issues = append(issues, rule.Issues...)
	}

	for _, rule := range comparison.PageRules {
		issues = append(issues, rule.Issues...)
	}

	return issues
}

func (dc *DocumentComparator) collectContentIssues(comparison *ContentComparison) []FormatIssue {
	var issues []FormatIssue

	for _, element := range comparison.Paragraphs {
		issues = append(issues, element.Issues...)
	}

	for _, element := range comparison.Tables {
		issues = append(issues, element.Issues...)
	}

	return issues
}

func (dc *DocumentComparator) collectStyleIssues(comparison *StyleComparison) []FormatIssue {
	var issues []FormatIssue

	for _, style := range comparison.ParagraphStyles {
		issues = append(issues, style.Issues...)
	}

	for _, style := range comparison.CharacterStyles {
		issues = append(issues, style.Issues...)
	}

	for _, style := range comparison.TableStyles {
		issues = append(issues, style.Issues...)
	}

	return issues
}

// 建议生成方法
func (dc *DocumentComparator) generateRecommendations(formatComparison *FormatComparison, contentComparison *ContentComparison, styleComparison *StyleComparison) []Recommendation {
	var recommendations []Recommendation

	// 基于格式比较生成建议
	for _, issue := range formatComparison.Issues {
		recommendation := Recommendation{
			ID:          fmt.Sprintf("format_%s", issue.ID),
			Type:        "format",
			Priority:    dc.getPriorityFromSeverity(issue.Severity),
			Description: issue.Description,
			Actions: []Action{{
				Type:        "fix",
				Description: "Fix format issue",
				Steps: []Step{{
					Order:       1,
					Description: "Apply suggested changes",
					Details:     issue.Description,
				}},
			}},
			Impact: "Format compliance",
		}
		recommendations = append(recommendations, recommendation)
	}

	// 基于内容比较生成建议
	for _, issue := range contentComparison.Issues {
		recommendation := Recommendation{
			ID:          fmt.Sprintf("content_%s", issue.ID),
			Type:        "content",
			Priority:    dc.getPriorityFromSeverity(issue.Severity),
			Description: issue.Description,
			Actions: []Action{{
				Type:        "fix",
				Description: "Fix content issue",
				Steps: []Step{{
					Order:       1,
					Description: "Apply suggested changes",
					Details:     issue.Description,
				}},
			}},
			Impact: "Content compliance",
		}
		recommendations = append(recommendations, recommendation)
	}

	// 基于样式比较生成建议
	for _, issue := range styleComparison.Issues {
		recommendation := Recommendation{
			ID:          fmt.Sprintf("style_%s", issue.ID),
			Type:        "style",
			Priority:    dc.getPriorityFromSeverity(issue.Severity),
			Description: issue.Description,
			Actions: []Action{{
				Type:        "fix",
				Description: "Fix style issue",
				Steps: []Step{{
					Order:       1,
					Description: "Apply suggested changes",
					Details:     issue.Description,
				}},
			}},
			Impact: "Style compliance",
		}
		recommendations = append(recommendations, recommendation)
	}

	return recommendations
}

// 摘要生成方法
func (dc *DocumentComparator) generateSummary(formatComparison *FormatComparison, contentComparison *ContentComparison, styleComparison *StyleComparison) ComparisonSummary {
	summary := ComparisonSummary{}

	// 统计问题
	allIssues := append(formatComparison.Issues, contentComparison.Issues...)
	allIssues = append(allIssues, styleComparison.Issues...)

	for _, issue := range allIssues {
		summary.TotalIssues++
		switch issue.Severity {
		case SeverityCritical:
			summary.CriticalIssues++
		case SeverityHigh:
			summary.HighIssues++
		case SeverityMedium:
			summary.MediumIssues++
		case SeverityLow:
			summary.LowIssues++
		}
	}

	// 统计合规规则
	summary.CompliantRules = dc.countCompliantRules(formatComparison, contentComparison, styleComparison)
	summary.NonCompliantRules = dc.countNonCompliantRules(formatComparison, contentComparison, styleComparison)

	// 计算总体分数
	summary.OverallScore = (formatComparison.Score + contentComparison.Score + styleComparison.Score) / 3.0

	// 统计建议数量
	summary.Recommendations = len(dc.generateRecommendations(formatComparison, contentComparison, styleComparison))

	return summary
}

// 合规率计算方法
func (dc *DocumentComparator) calculateComplianceRate(formatComparison *FormatComparison, contentComparison *ContentComparison, styleComparison *StyleComparison) float64 {
	totalRules := dc.countCompliantRules(formatComparison, contentComparison, styleComparison) +
		dc.countNonCompliantRules(formatComparison, contentComparison, styleComparison)

	if totalRules == 0 {
		return 100.0
	}

	compliantRules := dc.countCompliantRules(formatComparison, contentComparison, styleComparison)
	return float64(compliantRules) / float64(totalRules) * 100.0
}

// 辅助方法
func (dc *DocumentComparator) countCompliantRules(formatComparison *FormatComparison, contentComparison *ContentComparison, styleComparison *StyleComparison) int {
	count := 0

	for _, rule := range formatComparison.FontRules {
		if rule.Compliant {
			count++
		}
	}

	for _, rule := range formatComparison.ParagraphRules {
		if rule.Compliant {
			count++
		}
	}

	for _, rule := range formatComparison.TableRules {
		if rule.Compliant {
			count++
		}
	}

	for _, rule := range formatComparison.PageRules {
		if rule.Compliant {
			count++
		}
	}

	return count
}

func (dc *DocumentComparator) countNonCompliantRules(formatComparison *FormatComparison, contentComparison *ContentComparison, styleComparison *StyleComparison) int {
	count := 0

	for _, rule := range formatComparison.FontRules {
		if !rule.Compliant {
			count++
		}
	}

	for _, rule := range formatComparison.ParagraphRules {
		if !rule.Compliant {
			count++
		}
	}

	for _, rule := range formatComparison.TableRules {
		if !rule.Compliant {
			count++
		}
	}

	for _, rule := range formatComparison.PageRules {
		if !rule.Compliant {
			count++
		}
	}

	return count
}

func (dc *DocumentComparator) getPriorityFromSeverity(severity Severity) Priority {
	switch severity {
	case SeverityCritical:
		return PriorityUrgent
	case SeverityHigh:
		return PriorityHigh
	case SeverityMedium:
		return PriorityMedium
	case SeverityLow:
		return PriorityLow
	default:
		return PriorityMedium
	}
}

// 占位符方法（需要根据具体需求实现）
func (dc *DocumentComparator) isParagraphElementCompliant(docPara, templatePara types.Paragraph) bool {
	return true // 占位符实现
}

func (dc *DocumentComparator) calculateParagraphElementScore(docPara, templatePara types.Paragraph) float64 {
	return 100.0 // 占位符实现
}

func (dc *DocumentComparator) getParagraphElementDifferences(docPara, templatePara types.Paragraph) []Difference {
	return []Difference{} // 占位符实现
}

func (dc *DocumentComparator) isTableElementCompliant(docTable, templateTable types.Table) bool {
	return true // 占位符实现
}

func (dc *DocumentComparator) calculateTableElementScore(docTable, templateTable types.Table) float64 {
	return 100.0 // 占位符实现
}

func (dc *DocumentComparator) getTableElementDifferences(docTable, templateTable types.Table) []Difference {
	return []Difference{} // 占位符实现
}

func (dc *DocumentComparator) isParagraphStyleCompliant(docStyle, templateStyle types.ParagraphStyle) bool {
	return true // 占位符实现
}

func (dc *DocumentComparator) calculateParagraphStyleScore(docStyle, templateStyle types.ParagraphStyle) float64 {
	return 100.0 // 占位符实现
}

func (dc *DocumentComparator) getParagraphStyleDifferences(docStyle, templateStyle types.ParagraphStyle) []Difference {
	return []Difference{} // 占位符实现
}

func (dc *DocumentComparator) isCharacterStyleCompliant(docStyle, templateStyle types.CharacterStyle) bool {
	return true // 占位符实现
}

func (dc *DocumentComparator) calculateCharacterStyleScore(docStyle, templateStyle types.CharacterStyle) float64 {
	return 100.0 // 占位符实现
}

func (dc *DocumentComparator) getCharacterStyleDifferences(docStyle, templateStyle types.CharacterStyle) []Difference {
	return []Difference{} // 占位符实现
}

func (dc *DocumentComparator) isTableStyleCompliant(docStyle, templateStyle types.TableStyle) bool {
	return true // 占位符实现
}

func (dc *DocumentComparator) calculateTableStyleScore(docStyle, templateStyle types.TableStyle) float64 {
	return 100.0 // 占位符实现
}

func (dc *DocumentComparator) getTableStyleDifferences(docStyle, templateStyle types.TableStyle) []Difference {
	return []Difference{} // 占位符实现
}
