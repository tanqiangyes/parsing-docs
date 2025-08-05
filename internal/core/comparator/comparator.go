package comparator

import (
	"fmt"
	"strings"

	"docs-parser/internal/core/annotator"
	"docs-parser/internal/core/types"
	"docs-parser/internal/formats"
	"docs-parser/internal/templates"
)

// DocumentComparator 文档对比器
type DocumentComparator struct {
	annotator       *annotator.Annotator
	wordParser      *formats.WordParser
	templateManager *templates.TemplateManager
}

// NewDocumentComparator 创建新的文档对比器
func NewDocumentComparator() *DocumentComparator {
	return &DocumentComparator{
		annotator:       annotator.NewAnnotator(),
		wordParser:      formats.NewWordParser(),
		templateManager: templates.NewTemplateManager(""),
	}
}

// CompareWithTemplate 与模板进行对比
func (dc *DocumentComparator) CompareWithTemplate(docPath, templatePath string) (*ComparisonReport, error) {
	// 解析文档
	doc, err := dc.wordParser.ParseDocument(docPath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse document: %w", err)
	}

	// 解析模板
	template, err := dc.wordParser.ParseDocument(templatePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse template: %w", err)
	}

	// 只对比格式规则（文本运行级别的字体比较）
	formatComparison, err := dc.CompareFormatRules(&doc.FormatRules, &template.FormatRules, doc, template)
	if err != nil {
		return nil, fmt.Errorf("format comparison failed: %w", err)
	}

	// 创建对比报告
	report := &ComparisonReport{
		DocumentPath:     docPath,
		TemplatePath:     templatePath,
		Issues:           formatComparison.Issues,
		FormatComparison: formatComparison,
		ContentComparison: &ContentComparison{Issues: []types.FormatIssue{}},
		StyleComparison:  &StyleComparison{Issues: []types.FormatIssue{}},
	}

	// 如果有格式问题，自动生成标注文档
	fmt.Printf("DEBUG: 发现 %d 个问题，准备生成标注文档\n", len(formatComparison.Issues))
	if len(formatComparison.Issues) > 0 {
		fmt.Printf("DEBUG: 开始生成标注文档...\n")
		annotatedPath, err := dc.annotator.AnnotateDocumentWithIssues(docPath, formatComparison.Issues)
		if err != nil {
			fmt.Printf("警告: 生成标注文档失败: %v\n", err)
		} else {
			report.AnnotatedDocumentPath = annotatedPath
			fmt.Printf("已生成标注文档: %s\n", annotatedPath)
		}
	}

	return report, nil
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
	formatComparison, err := dc.CompareFormatRules(&doc1.FormatRules, &doc2.FormatRules, doc1, doc2)
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

	// 合并所有问题
	var allIssues []types.FormatIssue
	allIssues = append(allIssues, formatComparison.Issues...)
	allIssues = append(allIssues, contentComparison.Issues...)
	allIssues = append(allIssues, styleComparison.Issues...)

	// 创建对比报告
	report := &ComparisonReport{
		DocumentPath:     doc1Path,
		TemplatePath:     doc2Path,
		Issues:           allIssues,
		FormatComparison: formatComparison,
		ContentComparison: contentComparison,
		StyleComparison:  styleComparison,
	}

	// 如果有格式问题，自动生成标注文档
	if len(allIssues) > 0 {
		annotatedPath, err := dc.annotator.AnnotateDocumentWithIssues(doc1Path, allIssues)
		if err != nil {
			fmt.Printf("警告: 生成标注文档失败: %v\n", err)
		} else {
			report.AnnotatedDocumentPath = annotatedPath
			fmt.Printf("已生成标注文档: %s\n", annotatedPath)
		}
	}

	return report, nil
}

// CompareFormatRules 对比格式规则
func (dc *DocumentComparator) CompareFormatRules(docRules, templateRules *types.FormatRules, doc, template *types.Document) (*FormatComparison, error) {
	comparison := &FormatComparison{
		FontRules:      []RuleComparison{},
		ParagraphRules: []RuleComparison{},
		TableRules:     []RuleComparison{},
		PageRules:      []RuleComparison{},
		StyleRules:     []RuleComparison{},
		Score:          0.0,
		Issues:         []types.FormatIssue{},
	}

	// 对比段落格式（对齐、缩进、间距等）
	fmt.Printf("DEBUG: 开始对比段落格式，文档段落数: %d, 模板段落数: %d\n", len(docRules.ParagraphRules), len(templateRules.ParagraphRules))
	dc.compareParagraphFormats(docRules.ParagraphRules, templateRules.ParagraphRules, &comparison.Issues)

	// 对比文本运行级别的字体信息（合并同一文本的多个问题）
	if doc != nil && template != nil {
		fmt.Printf("DEBUG: 开始对比内容字体，文档段落数: %d, 模板段落数: %d\n", len(doc.Content.Paragraphs), len(template.Content.Paragraphs))
		dc.compareContentFonts(&doc.Content, &template.Content, &comparison.Issues)
	}

	fmt.Printf("DEBUG: 格式对比问题数量: %d\n", len(comparison.Issues))
	return comparison, nil
}

// compareParagraphFormats 对比段落格式
func (dc *DocumentComparator) compareParagraphFormats(docRules, templateRules []types.ParagraphRule, issues *[]types.FormatIssue) {
	fmt.Printf("DEBUG: 开始对比段落格式，文档段落数: %d, 模板段落数: %d\n", len(docRules), len(templateRules))
	
	// 为每个段落生成具体的格式对比
	for i, templateRule := range templateRules {
		if i < len(docRules) {
			docRule := docRules[i]
			
			fmt.Printf("DEBUG: 对比段落 %d: 文档对齐=%s, 模板对齐=%s\n", i+1, docRule.Alignment, templateRule.Alignment)
			
			// 检查对齐方式
			if docRule.Alignment != templateRule.Alignment {
				currentFormat := map[string]interface{}{
					"alignment": docRule.Alignment,
					"spacing":   docRule.Spacing,
				}
				expectedFormat := map[string]interface{}{
					"alignment": templateRule.Alignment,
					"spacing":   templateRule.Spacing,
				}
				
				*issues = append(*issues, types.FormatIssue{
					ID:          fmt.Sprintf("paragraph_format_%d", i),
					Type:        "paragraph",
					Severity:    "medium",
					Location:    fmt.Sprintf("第%d段", i+1),
					Description: fmt.Sprintf("第%d段对齐方式不符合模板要求", i+1),
					Current:     currentFormat,
					Expected:    expectedFormat,
					Rule:        "paragraph_format",
					Suggestions: []string{"调整段落对齐方式以匹配模板"},
				})
				fmt.Printf("DEBUG: 发现段落对齐问题\n")
			}
			
			// 检查间距
			if docRule.Spacing.Before != templateRule.Spacing.Before || 
			   docRule.Spacing.After != templateRule.Spacing.After {
				currentFormat := map[string]interface{}{
					"spacingBefore": docRule.Spacing.Before,
					"spacingAfter":  docRule.Spacing.After,
				}
				expectedFormat := map[string]interface{}{
					"spacingBefore": templateRule.Spacing.Before,
					"spacingAfter":  templateRule.Spacing.After,
				}
				
				*issues = append(*issues, types.FormatIssue{
					ID:          fmt.Sprintf("paragraph_spacing_%d", i),
					Type:        "paragraph",
					Severity:    "low",
					Location:    fmt.Sprintf("第%d段", i+1),
					Description: fmt.Sprintf("第%d段间距不符合模板要求", i+1),
					Current:     currentFormat,
					Expected:    expectedFormat,
					Rule:        "paragraph_format",
					Suggestions: []string{"调整段落间距以匹配模板"},
				})
				fmt.Printf("DEBUG: 发现段落间距问题\n")
			}
		}
	}
	
	fmt.Printf("DEBUG: 段落格式对比完成，发现问题数: %d\n", len(*issues))
}

// compareFontFormats 对比字体格式
func (dc *DocumentComparator) compareFontFormats(docRules, templateRules []types.FontRule, issues *[]types.FormatIssue) {
	// 为每个字体规则生成具体的格式对比
	for i, templateRule := range templateRules {
		if i < len(docRules) {
			docRule := docRules[i]
			
			// 检查字体名称
			if docRule.Name != templateRule.Name {
				currentFormat := map[string]interface{}{
					"fontName": docRule.Name,
					"fontSize": docRule.Size,
				}
				expectedFormat := map[string]interface{}{
					"fontName": templateRule.Name,
					"fontSize": templateRule.Size,
				}
				
				*issues = append(*issues, types.FormatIssue{
					ID:          fmt.Sprintf("font_name_%d", i),
					Type:        "font",
					Severity:    "medium",
					Location:    fmt.Sprintf("第%d个字体规则", i+1),
					Description: fmt.Sprintf("字体名称不符合模板要求"),
					Current:     currentFormat,
					Expected:    expectedFormat,
					Rule:        "font_format",
					Suggestions: []string{"调整字体名称以匹配模板"},
				})
			}
			
			// 检查字体大小
			if docRule.Size != templateRule.Size {
				currentFormat := map[string]interface{}{
					"fontName": docRule.Name,
					"fontSize": docRule.Size,
				}
				expectedFormat := map[string]interface{}{
					"fontName": templateRule.Name,
					"fontSize": templateRule.Size,
				}
				
				*issues = append(*issues, types.FormatIssue{
					ID:          fmt.Sprintf("font_size_%d", i),
					Type:        "font",
					Severity:    "medium",
					Location:    fmt.Sprintf("第%d个字体规则", i+1),
					Description: fmt.Sprintf("字体大小不符合模板要求"),
					Current:     currentFormat,
					Expected:    expectedFormat,
					Rule:        "font_format",
					Suggestions: []string{"调整字体大小以匹配模板"},
				})
			}
			
			// 检查字体颜色
			if docRule.Color.RGB != templateRule.Color.RGB {
				currentFormat := map[string]interface{}{
					"fontName": docRule.Name,
					"fontColor": docRule.Color.RGB,
				}
				expectedFormat := map[string]interface{}{
					"fontName": templateRule.Name,
					"fontColor": templateRule.Color.RGB,
				}
				
				*issues = append(*issues, types.FormatIssue{
					ID:          fmt.Sprintf("font_color_%d", i),
					Type:        "font",
					Severity:    "low",
					Location:    fmt.Sprintf("第%d个字体规则", i+1),
					Description: fmt.Sprintf("字体颜色不符合模板要求"),
					Current:     currentFormat,
					Expected:    expectedFormat,
					Rule:        "font_format",
					Suggestions: []string{"调整字体颜色以匹配模板"},
				})
			}
		}
	}
}

// CompareContent 对比内容
func (dc *DocumentComparator) CompareContent(docContent, templateContent *types.DocumentContent) (*ContentComparison, error) {
	comparison := &ContentComparison{
		Paragraphs:    []ElementComparison{},
		Tables:        []ElementComparison{},
		Headers:       []ElementComparison{},
		Footers:       []ElementComparison{},
		Images:        []ElementComparison{},
		Score:         0.0,
		Issues:        []types.FormatIssue{},
	}

	// 简化的内容对比
	if len(docContent.Paragraphs) != len(templateContent.Paragraphs) {
		comparison.Issues = append(comparison.Issues, types.FormatIssue{
			ID:          "content_paragraph_count_mismatch",
			Type:        "content",
			Severity:    "medium",
			Location:    "document",
			Description: fmt.Sprintf("内容段落数量不匹配: 文档有 %d 个段落，模板有 %d 个段落", len(docContent.Paragraphs), len(templateContent.Paragraphs)),
			Current:     len(docContent.Paragraphs),
			Expected:    len(templateContent.Paragraphs),
			Rule:        "content_paragraph_count",
			Suggestions: []string{"调整内容段落数量以匹配模板"},
		})
	}

	return comparison, nil
}

// CompareStyles 对比样式
func (dc *DocumentComparator) CompareStyles(docStyles, templateStyles *types.DocumentStyles) (*StyleComparison, error) {
	comparison := &StyleComparison{
		ParagraphStyles: []StyleElementComparison{},
		CharacterStyles: []StyleElementComparison{},
		TableStyles:     []StyleElementComparison{},
		Score:           0.0,
		Issues:          []types.FormatIssue{},
	}

	// 对比段落样式
	dc.compareParagraphStyles(docStyles.ParagraphStyles, templateStyles.ParagraphStyles, &comparison.Issues)
	
	// 对比字符样式
	dc.compareCharacterStyles(docStyles.CharacterStyles, templateStyles.CharacterStyles, &comparison.Issues)
	
	// 对比表格样式
	dc.compareTableStyles(docStyles.TableStyles, templateStyles.TableStyles, &comparison.Issues)

	return comparison, nil
}

// compareParagraphStyles 对比段落样式
func (dc *DocumentComparator) compareParagraphStyles(docStyles, templateStyles []types.ParagraphStyle, issues *[]types.FormatIssue) {
	// 创建样式名称映射
	docStyleMap := make(map[string]bool)
	templateStyleMap := make(map[string]bool)
	
	for _, style := range docStyles {
		docStyleMap[style.Name] = true
	}
	
	for _, style := range templateStyles {
		templateStyleMap[style.Name] = true
	}
	
	// 找出缺少的样式
	var missingStyles []string
	for _, templateStyle := range templateStyles {
		if !docStyleMap[templateStyle.Name] {
			missingStyles = append(missingStyles, templateStyle.Name)
		}
	}
	
	// 找出多余的样式
	var extraStyles []string
	for _, docStyle := range docStyles {
		if !templateStyleMap[docStyle.Name] {
			extraStyles = append(extraStyles, docStyle.Name)
		}
	}
	
	// 生成问题报告
	if len(missingStyles) > 0 {
		*issues = append(*issues, types.FormatIssue{
			ID:          "missing_paragraph_styles",
			Type:        "style",
			Severity:    "medium",
			Location:    "document",
			Description: fmt.Sprintf("缺少段落样式: %s", strings.Join(missingStyles, ", ")),
			Current:     fmt.Sprintf("文档包含 %d 个段落样式", len(docStyles)),
			Expected:    fmt.Sprintf("模板包含 %d 个段落样式", len(templateStyles)),
			Rule:        "missing_paragraph_styles",
			Suggestions: []string{fmt.Sprintf("添加缺少的段落样式: %s", strings.Join(missingStyles, ", "))},
		})
	}
	
	if len(extraStyles) > 0 {
		*issues = append(*issues, types.FormatIssue{
			ID:          "extra_paragraph_styles",
			Type:        "style",
			Severity:    "low",
			Location:    "document",
			Description: fmt.Sprintf("多余的段落样式: %s", strings.Join(extraStyles, ", ")),
			Current:     fmt.Sprintf("文档包含 %d 个段落样式", len(docStyles)),
			Expected:    fmt.Sprintf("模板包含 %d 个段落样式", len(templateStyles)),
			Rule:        "extra_paragraph_styles",
			Suggestions: []string{fmt.Sprintf("移除多余的段落样式: %s", strings.Join(extraStyles, ", "))},
		})
	}
}

// compareCharacterStyles 对比字符样式
func (dc *DocumentComparator) compareCharacterStyles(docStyles, templateStyles []types.CharacterStyle, issues *[]types.FormatIssue) {
	// 创建样式名称映射
	docStyleMap := make(map[string]bool)
	templateStyleMap := make(map[string]bool)
	
	for _, style := range docStyles {
		docStyleMap[style.Name] = true
	}
	
	for _, style := range templateStyles {
		templateStyleMap[style.Name] = true
	}
	
	// 找出缺少的样式
	var missingStyles []string
	for _, templateStyle := range templateStyles {
		if !docStyleMap[templateStyle.Name] {
			missingStyles = append(missingStyles, templateStyle.Name)
		}
	}
	
	// 找出多余的样式
	var extraStyles []string
	for _, docStyle := range docStyles {
		if !templateStyleMap[docStyle.Name] {
			extraStyles = append(extraStyles, docStyle.Name)
		}
	}
	
	// 生成问题报告
	if len(missingStyles) > 0 {
		*issues = append(*issues, types.FormatIssue{
			ID:          "missing_character_styles",
			Type:        "style",
			Severity:    "medium",
			Location:    "document",
			Description: fmt.Sprintf("缺少字符样式: %s", strings.Join(missingStyles, ", ")),
			Current:     fmt.Sprintf("文档包含 %d 个字符样式", len(docStyles)),
			Expected:    fmt.Sprintf("模板包含 %d 个字符样式", len(templateStyles)),
			Rule:        "missing_character_styles",
			Suggestions: []string{fmt.Sprintf("添加缺少的字符样式: %s", strings.Join(missingStyles, ", "))},
		})
	}
	
	if len(extraStyles) > 0 {
		*issues = append(*issues, types.FormatIssue{
			ID:          "extra_character_styles",
			Type:        "style",
			Severity:    "low",
			Location:    "document",
			Description: fmt.Sprintf("多余的字符样式: %s", strings.Join(extraStyles, ", ")),
			Current:     fmt.Sprintf("文档包含 %d 个字符样式", len(docStyles)),
			Expected:    fmt.Sprintf("模板包含 %d 个字符样式", len(templateStyles)),
			Rule:        "extra_character_styles",
			Suggestions: []string{fmt.Sprintf("移除多余的字符样式: %s", strings.Join(extraStyles, ", "))},
		})
	}
}

// compareContentFonts 对比文档内容中的实际字体信息
func (dc *DocumentComparator) compareContentFonts(docContent, templateContent *types.DocumentContent, issues *[]types.FormatIssue) {
	fmt.Printf("DEBUG: 开始对比内容字体，文档段落数: %d, 模板段落数: %d\n", len(docContent.Paragraphs), len(templateContent.Paragraphs))
	
	// 为每个段落比较字体信息
	for i, templatePara := range templateContent.Paragraphs {
		if i < len(docContent.Paragraphs) {
			docPara := docContent.Paragraphs[i]
			
			// 比较段落中的文本运行
			for j, templateRun := range templatePara.Runs {
				if j < len(docPara.Runs) {
					docRun := docPara.Runs[j]
					
					// 收集这个文本运行的所有字体问题
					var fontIssues []string
					var currentFormat map[string]interface{}
					var expectedFormat map[string]interface{}
					
					// 检查字体名称
					if docRun.Font.Name != templateRun.Font.Name {
						fontIssues = append(fontIssues, fmt.Sprintf("字体名称: 文档=%s, 模板=%s", docRun.Font.Name, templateRun.Font.Name))
						fmt.Printf("DEBUG: 发现字体名称问题: 文档=%s, 模板=%s\n", docRun.Font.Name, templateRun.Font.Name)
					}
					
					// 检查字体大小
					if docRun.Font.Size != templateRun.Font.Size {
						fontIssues = append(fontIssues, fmt.Sprintf("字体大小: 文档=%.1f, 模板=%.1f", docRun.Font.Size, templateRun.Font.Size))
						fmt.Printf("DEBUG: 发现字体大小问题: 文档=%.1f, 模板=%.1f\n", docRun.Font.Size, templateRun.Font.Size)
					}
					
					// 检查字体颜色
					if docRun.Font.Color.RGB != templateRun.Font.Color.RGB {
						fontIssues = append(fontIssues, fmt.Sprintf("字体颜色: 文档=%s, 模板=%s", docRun.Font.Color.RGB, templateRun.Font.Color.RGB))
					}
					
					// 检查粗体
					if docRun.Font.Bold != templateRun.Font.Bold {
						fontIssues = append(fontIssues, fmt.Sprintf("粗体: 文档=%v, 模板=%v", docRun.Font.Bold, templateRun.Font.Bold))
					}
					
					// 检查斜体
					if docRun.Font.Italic != templateRun.Font.Italic {
						fontIssues = append(fontIssues, fmt.Sprintf("斜体: 文档=%v, 模板=%v", docRun.Font.Italic, templateRun.Font.Italic))
					}
					
					// 如果有字体问题，创建一个合并的问题
					if len(fontIssues) > 0 {
						currentFormat = map[string]interface{}{
							"fontName":  docRun.Font.Name,
							"fontSize":  docRun.Font.Size,
							"fontColor": docRun.Font.Color.RGB,
							"bold":      docRun.Font.Bold,
							"italic":    docRun.Font.Italic,
						}
						expectedFormat = map[string]interface{}{
							"fontName":  templateRun.Font.Name,
							"fontSize":  templateRun.Font.Size,
							"fontColor": templateRun.Font.Color.RGB,
							"bold":      templateRun.Font.Bold,
							"italic":    templateRun.Font.Italic,
						}
						
						*issues = append(*issues, types.FormatIssue{
							ID:          fmt.Sprintf("font_format_%d_%d", i, j),
							Type:        "font",
							Severity:    "medium",
							Location:    fmt.Sprintf("第%d段第%d个文本", i+1, j+1),
							Description: fmt.Sprintf("第%d段第%d个文本的字体格式不符合模板要求", i+1, j+1),
							Current:     currentFormat,
							Expected:    expectedFormat,
							Rule:        "font_format",
							Suggestions: []string{fmt.Sprintf("调整字体格式: %s", strings.Join(fontIssues, "; "))},
						})
					}
				}
			}
		}
	}
}

// compareTableStyles 对比表格样式
func (dc *DocumentComparator) compareTableStyles(docStyles, templateStyles []types.TableStyle, issues *[]types.FormatIssue) {
	// 创建样式名称映射
	docStyleMap := make(map[string]bool)
	templateStyleMap := make(map[string]bool)
	
	for _, style := range docStyles {
		docStyleMap[style.Name] = true
	}
	
	for _, style := range templateStyles {
		templateStyleMap[style.Name] = true
	}
	
	// 找出缺少的样式
	var missingStyles []string
	for _, templateStyle := range templateStyles {
		if !docStyleMap[templateStyle.Name] {
			missingStyles = append(missingStyles, templateStyle.Name)
		}
	}
	
	// 找出多余的样式
	var extraStyles []string
	for _, docStyle := range docStyles {
		if !templateStyleMap[docStyle.Name] {
			extraStyles = append(extraStyles, docStyle.Name)
		}
	}
	
	// 生成问题报告
	if len(missingStyles) > 0 {
		*issues = append(*issues, types.FormatIssue{
			ID:          "missing_table_styles",
			Type:        "style",
			Severity:    "medium",
			Location:    "document",
			Description: fmt.Sprintf("缺少表格样式: %s", strings.Join(missingStyles, ", ")),
			Current:     fmt.Sprintf("文档包含 %d 个表格样式", len(docStyles)),
			Expected:    fmt.Sprintf("模板包含 %d 个表格样式", len(templateStyles)),
			Rule:        "missing_table_styles",
			Suggestions: []string{fmt.Sprintf("添加缺少的表格样式: %s", strings.Join(missingStyles, ", "))},
		})
	}
	
	if len(extraStyles) > 0 {
		*issues = append(*issues, types.FormatIssue{
			ID:          "extra_table_styles",
			Type:        "style",
			Severity:    "low",
			Location:    "document",
			Description: fmt.Sprintf("多余的表格样式: %s", strings.Join(extraStyles, ", ")),
			Current:     fmt.Sprintf("文档包含 %d 个表格样式", len(docStyles)),
			Expected:    fmt.Sprintf("模板包含 %d 个表格样式", len(templateStyles)),
			Rule:        "extra_table_styles",
			Suggestions: []string{fmt.Sprintf("移除多余的表格样式: %s", strings.Join(extraStyles, ", "))},
		})
	}
}
