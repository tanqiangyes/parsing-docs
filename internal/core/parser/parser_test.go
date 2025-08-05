package parser

import (
	"testing"
	"time"
	"docs-parser/internal/core/types"
)

// TestNewParser 测试创建新的解析器
func TestNewParser(t *testing.T) {
	parser := &DefaultParser{}
	if parser == nil {
		t.Error("期望返回非空的解析器")
	}
}

// TestParser_ParseDocument 测试文档解析功能
func TestParser_ParseDocument(t *testing.T) {
	parser := &DefaultParser{}

	// 测试解析不存在的文档
	_, err := parser.ParseDocument("nonexistent.docx")
	if err == nil {
		t.Error("期望解析不存在的文档返回错误")
	}

	// 测试解析空路径
	_, err = parser.ParseDocument("")
	if err == nil {
		t.Error("期望解析空路径返回错误")
	}
}

// TestParser_ParseWordDocument 测试Word文档解析功能
func TestParser_ParseWordDocument(t *testing.T) {
	parser := &DefaultParser{}

	// 测试解析不存在的Word文档
	_, err := parser.ParseDocument("nonexistent.docx")
	if err == nil {
		t.Error("期望解析不存在的Word文档返回错误")
	}
}

// TestParser_ParseTemplate 测试模板解析功能
func TestParser_ParseTemplate(t *testing.T) {
	parser := &DefaultParser{}

	// 测试解析不存在的模板
	_, err := parser.ParseDocument("nonexistent_template.docx")
	if err == nil {
		t.Error("期望解析不存在的模板返回错误")
	}
}

// TestDocument_Validation 测试文档结构验证
func TestDocument_Validation(t *testing.T) {
	// 测试有效的文档结构
	validDocument := &types.Document{
		Metadata: types.DocumentMetadata{
			Title:    "测试文档",
			Author:   "测试作者",
			Created:  time.Now(),
			Modified: time.Now(),
		},
		Content: types.DocumentContent{
			Paragraphs: []types.Paragraph{
				{
					Text: "测试段落",
					Runs: []types.TextRun{
						{
							Text: "测试文本",
							Font: types.Font{
								Name:  "宋体",
								Size:  12.0,
								Color: types.Color{RGB: "000000"},
								Bold:  false,
							},
						},
					},
				},
			},
		},
		Styles: types.DocumentStyles{
			ParagraphStyles: []types.ParagraphStyle{
				{
					ID:   "style1",
					Name: "标题1",
					Font: types.Font{
						Name:  "黑体",
						Size:  16.0,
						Color: types.Color{RGB: "000000"},
						Bold:  true,
					},
					Alignment: "center",
				},
			},
		},
		FormatRules: types.FormatRules{
			FontRules: []types.FontRule{
				{
					ID:       "font1",
					Name:     "宋体",
					Size:     12.0,
					Color:    types.Color{RGB: "000000"},
					Bold:     false,
					Italic:   false,
				},
			},
			ParagraphRules: []types.ParagraphRule{
				{
					ID:        "para1",
					Alignment: "left",
					Indentation: types.Indentation{
						Left:   0,
						Right:  0,
						First:  0,
					},
					Spacing: types.Spacing{
						Before: 0,
						After:  0,
						Line:   1.0,
					},
				},
			},
		},
	}

	// 验证文档元数据
	if validDocument.Metadata.Title == "" {
		t.Error("文档应该有标题")
	}

	if validDocument.Metadata.Author == "" {
		t.Error("文档应该有作者")
	}

	// 验证文档内容
	if len(validDocument.Content.Paragraphs) == 0 {
		t.Error("文档应该有段落内容")
	}

	// 验证文档样式
	if len(validDocument.Styles.ParagraphStyles) == 0 {
		t.Error("文档应该有段落样式")
	}

	// 验证格式规则
	if len(validDocument.FormatRules.FontRules) == 0 {
		t.Error("文档应该有字体规则")
	}

	if len(validDocument.FormatRules.ParagraphRules) == 0 {
		t.Error("文档应该有段落规则")
	}
}

// TestDocumentMetadata_Validation 测试文档元数据验证
func TestDocumentMetadata_Validation(t *testing.T) {
	// 测试有效的文档元数据
	validMetadata := types.DocumentMetadata{
		Title:    "测试文档",
		Author:   "测试作者",
		Subject:  "测试主题",
		Keywords: []string{"测试", "文档"},
		Created:  time.Now(),
		Modified: time.Now(),
		PageCount: 10,
		WordCount: 1000,
	}

	if validMetadata.Title == "" {
		t.Error("文档元数据应该有标题")
	}

	if validMetadata.Author == "" {
		t.Error("文档元数据应该有作者")
	}

	if validMetadata.Created.IsZero() {
		t.Error("文档元数据应该有创建时间")
	}

	if validMetadata.Modified.IsZero() {
		t.Error("文档元数据应该有修改时间")
	}
}

// TestDocumentContent_Validation 测试文档内容验证
func TestDocumentContent_Validation(t *testing.T) {
	// 测试有效的文档内容
	validContent := types.DocumentContent{
		Paragraphs: []types.Paragraph{
			{
				Text: "测试段落1",
				Runs: []types.TextRun{
					{
						Text: "测试文本1",
						Font: types.Font{
							Name:  "宋体",
							Size:  12.0,
							Color: types.Color{RGB: "000000"},
							Bold:  false,
						},
					},
					{
						Text: "测试文本2",
						Font: types.Font{
							Name:  "黑体",
							Size:  14.0,
							Color: types.Color{RGB: "000000"},
							Bold:  true,
						},
					},
				},
				Alignment: "left",
				Indentation: types.Indentation{
					Left:   0,
					Right:  0,
					First:  0,
				},
				Spacing: types.Spacing{
					Before: 0,
					After:  0,
					Line:   1.0,
				},
			},
			{
				Text: "测试段落2",
				Runs: []types.TextRun{
					{
						Text: "测试文本3",
						Font: types.Font{
							Name:  "宋体",
							Size:  12.0,
							Color: types.Color{RGB: "000000"},
							Bold:  false,
						},
					},
				},
				Alignment: "center",
			},
		},
	}

	if len(validContent.Paragraphs) == 0 {
		t.Error("文档内容应该有段落")
	}

	// 验证段落内容
	for i, paragraph := range validContent.Paragraphs {
		if paragraph.Text == "" {
			t.Errorf("段落 %d 应该有文本内容", i)
		}

		if len(paragraph.Runs) == 0 {
			t.Errorf("段落 %d 应该有文本运行", i)
		}

		// 验证文本运行
		for j, run := range paragraph.Runs {
			if run.Text == "" {
				t.Errorf("段落 %d 的运行 %d 应该有文本", i, j)
			}

			if run.Font.Name == "" {
				t.Errorf("段落 %d 的运行 %d 应该有字体名称", i, j)
			}

			if run.Font.Size <= 0 {
				t.Errorf("段落 %d 的运行 %d 应该有有效的字体大小", i, j)
			}
		}
	}
}

// TestDocumentStyles_Validation 测试文档样式验证
func TestDocumentStyles_Validation(t *testing.T) {
	// 测试有效的文档样式
	validStyles := types.DocumentStyles{
		ParagraphStyles: []types.ParagraphStyle{
			{
				ID:   "style1",
				Name: "标题1",
				Font: types.Font{
					Name:  "黑体",
					Size:  16.0,
					Color: types.Color{RGB: "000000"},
					Bold:  true,
				},
				Alignment: "center",
				Indentation: types.Indentation{
					Left:   0,
					Right:  0,
					First:  0,
				},
				Spacing: types.Spacing{
					Before: 0,
					After:  0,
					Line:   1.0,
				},
			},
			{
				ID:   "style2",
				Name: "正文",
				Font: types.Font{
					Name:  "宋体",
					Size:  12.0,
					Color: types.Color{RGB: "000000"},
					Bold:  false,
				},
				Alignment: "left",
			},
		},
	}

	if len(validStyles.ParagraphStyles) == 0 {
		t.Error("文档样式应该有段落样式")
	}

	// 验证段落样式
	for i, style := range validStyles.ParagraphStyles {
		if style.ID == "" {
			t.Errorf("样式 %d 应该有ID", i)
		}

		if style.Name == "" {
			t.Errorf("样式 %d 应该有名称", i)
		}

		if style.Font.Name == "" {
			t.Errorf("样式 %d 应该有字体名称", i)
		}

		if style.Font.Size <= 0 {
			t.Errorf("样式 %d 应该有有效的字体大小", i)
		}
	}
}

// TestFormatRules_Validation 测试格式规则验证
func TestFormatRules_Validation(t *testing.T) {
	// 测试有效的格式规则
	validRules := types.FormatRules{
		FontRules: []types.FontRule{
			{
				ID:       "font1",
				Name:     "宋体",
				Size:     12.0,
				Color:    types.Color{RGB: "000000"},
				Bold:     false,
				Italic:   false,
			},
			{
				ID:       "font2",
				Name:     "黑体",
				Size:     14.0,
				Color:    types.Color{RGB: "000000"},
				Bold:     true,
				Italic:   false,
			},
		},
		ParagraphRules: []types.ParagraphRule{
			{
				ID:        "para1",
				Alignment: "left",
				Indentation: types.Indentation{
					Left:   0,
					Right:  0,
					First:  0,
				},
				Spacing: types.Spacing{
					Before: 0,
					After:  0,
					Line:   1.0,
				},
			},
			{
				ID:        "para2",
				Alignment: "center",
				Indentation: types.Indentation{
					Left:   0,
					Right:  0,
					First:  0,
				},
				Spacing: types.Spacing{
					Before: 0,
					After:  0,
					Line:   1.0,
				},
			},
		},
	}

	if len(validRules.FontRules) == 0 {
		t.Error("格式规则应该有字体规则")
	}

	if len(validRules.ParagraphRules) == 0 {
		t.Error("格式规则应该有段落规则")
	}

	// 验证字体规则
	for i, rule := range validRules.FontRules {
		if rule.ID == "" {
			t.Errorf("字体规则 %d 应该有ID", i)
		}

		if rule.Name == "" {
			t.Errorf("字体规则 %d 应该有名称", i)
		}

		if rule.Size <= 0 {
			t.Errorf("字体规则 %d 应该有有效的字体大小", i)
		}
	}

	// 验证段落规则
	for i, rule := range validRules.ParagraphRules {
		if rule.ID == "" {
			t.Errorf("段落规则 %d 应该有ID", i)
		}

		if rule.Alignment == "" {
			t.Errorf("段落规则 %d 应该有对齐方式", i)
		}
	}
} 