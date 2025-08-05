package styles

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
	"time"

	"docs-parser/internal/core/types"
)

// DocxStyleParser DOCX样式解析器
type DocxStyleParser struct {
	baseParser Parser
}

// NewDocxStyleParser 创建新的DOCX样式解析器
func NewDocxStyleParser() Parser {
	return &DocxStyleParser{
		baseParser: NewParser(),
	}
}

// ParseStyles 解析DOCX文档样式
func (dsp *DocxStyleParser) ParseStyles(filePath string) (*types.StyleManager, error) {
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

	// 打开DOCX文件
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, fmt.Errorf("无法打开DOCX文件: %w", err)
	}
	defer reader.Close()

	// 解析样式文件
	for _, file := range reader.File {
		switch file.Name {
		case "word/styles.xml":
			if err := dsp.parseStylesXML(file, manager); err != nil {
				return nil, fmt.Errorf("解析样式XML失败: %w", err)
			}
		case "word/theme/theme1.xml":
			if err := dsp.parseThemeXML(file, manager); err != nil {
				return nil, fmt.Errorf("解析主题XML失败: %w", err)
			}
		case "word/numbering.xml":
			if err := dsp.parseNumberingXML(file, manager); err != nil {
				return nil, fmt.Errorf("解析编号XML失败: %w", err)
			}
		}
	}

	// 解析继承关系
	inheritance, err := dsp.ParseStyleInheritance(filePath)
	if err != nil {
		return nil, fmt.Errorf("解析样式继承关系失败: %w", err)
	}

	// 构建继承树
	for styleID, inheritanceInfo := range inheritance {
		if inheritanceInfo.BasedOn != "" {
			manager.InheritanceTree[inheritanceInfo.BasedOn] = append(
				manager.InheritanceTree[inheritanceInfo.BasedOn], styleID)
		}
	}

	// 验证样式
	validation, err := dsp.ValidateStyles(manager)
	if err != nil {
		return nil, fmt.Errorf("验证样式失败: %w", err)
	}
	manager.Validation = *validation

	return manager, nil
}

// parseStylesXML 解析样式XML文件
func (dsp *DocxStyleParser) parseStylesXML(file *zip.File, manager *types.StyleManager) error {
	rc, err := file.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)
	var currentStyle *types.AdvancedStyle

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			continue
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "style":
				currentStyle = &types.AdvancedStyle{
					ID:          dsp.getAttribute(t, "styleId"),
					Name:        dsp.getAttribute(t, "name"),
					Type:        types.StyleType(dsp.getAttribute(t, "type")),
					Properties:  types.StyleProperties{},
					Inheritance: types.StyleInheritance{},
					Theme:       types.ThemeStyle{},
					Conditions:  []types.ConditionalStyle{},
					Conflicts:   []types.StyleConflict{},
					Validation:  types.StyleValidation{},
					Created:     time.Now(),
					Modified:    time.Now(),
					Version:     "1.0",
				}

			case "w:basedOn":
				if currentStyle != nil {
					currentStyle.Inheritance.BasedOn = dsp.getAttribute(t, "val")
				}

			case "w:next":
				if currentStyle != nil {
					currentStyle.Inheritance.Next = dsp.getAttribute(t, "val")
				}

			case "w:link":
				if currentStyle != nil {
					currentStyle.Inheritance.Linked = dsp.getAttribute(t, "val")
				}

			case "w:qFormat":
				if currentStyle != nil {
					currentStyle.Inheritance.QuickFormat = true
				}

			case "w:uiPriority":
				if currentStyle != nil {
					if priority, err := strconv.Atoi(dsp.getAttribute(t, "val")); err == nil {
						currentStyle.Inheritance.Priority = priority
					}
				}

			case "w:rsid":
				// 处理样式ID

			case "w:pPr":
				if currentStyle != nil {
					dsp.parseParagraphProperties(t, &currentStyle.Properties)
				}

			case "w:rPr":
				if currentStyle != nil {
					dsp.parseRunProperties(t, &currentStyle.Properties)
				}

			case "w:tblPr":
				if currentStyle != nil {
					dsp.parseTableProperties(t, &currentStyle.Properties)
				}

			case "w:tcPr":
				if currentStyle != nil {
					dsp.parseTableCellProperties(t, &currentStyle.Properties)
				}
			}

		case xml.EndElement:
			if t.Name.Local == "style" && currentStyle != nil {
				// 保存当前样式
				manager.Styles[currentStyle.ID] = currentStyle
				currentStyle = nil
			}
		}
	}

	return nil
}

// parseThemeXML 解析主题XML文件
func (dsp *DocxStyleParser) parseThemeXML(file *zip.File, manager *types.StyleManager) error {
	rc, err := file.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)
	var currentTheme *types.ThemeStyle

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			continue
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "a:theme":
				currentTheme = &types.ThemeStyle{
					ThemeName:   "Theme1",
					ColorScheme: "Office",
					FontScheme:  "Office",
					Effects:     "Office",
					Version:     "1.0",
				}

			case "a:clrScheme":
				if currentTheme != nil {
					currentTheme.ColorScheme = dsp.getAttribute(t, "name")
				}

			case "a:fontScheme":
				if currentTheme != nil {
					currentTheme.FontScheme = dsp.getAttribute(t, "name")
				}

			case "a:fmtScheme":
				if currentTheme != nil {
					currentTheme.Effects = dsp.getAttribute(t, "name")
				}
			}

		case xml.EndElement:
			if t.Name.Local == "a:theme" && currentTheme != nil {
				manager.ThemeStyles["Theme1"] = currentTheme
				currentTheme = nil
			}
		}
	}

	return nil
}

// parseNumberingXML 解析编号XML文件
func (dsp *DocxStyleParser) parseNumberingXML(file *zip.File, manager *types.StyleManager) error {
	rc, err := file.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)
	var currentNumbering types.Numbering

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			continue
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "w:num":
				currentNumbering = types.Numbering{
					Type:      "decimal",
					Format:    "1, 2, 3...",
					Start:     1,
					Increment: 1,
					Restart:   false,
					Level:     0,
					Text:      "",
					Alignment: "left",
				}

			case "w:numFmt":
				currentNumbering.Format = dsp.getAttribute(t, "val")

			case "w:numStart":
				if start, err := strconv.Atoi(dsp.getAttribute(t, "val")); err == nil {
					currentNumbering.Start = start
				}

			case "w:numRestart":
				currentNumbering.Restart = true
			}

		case xml.EndElement:
			if t.Name.Local == "w:num" {
				// 这里可以将编号信息应用到相应的样式中
				// 暂时跳过，因为需要与样式关联
			}
		}
	}

	return nil
}

// parseParagraphProperties 解析段落属性
func (dsp *DocxStyleParser) parseParagraphProperties(element xml.StartElement, props *types.StyleProperties) {
	for _, attr := range element.Attr {
		switch attr.Name.Local {
		case "jc":
			props.Alignment = types.Alignment(attr.Value)
		case "ind":
			// 处理缩进
			if left, err := strconv.ParseFloat(attr.Value, 64); err == nil {
				props.Indentation.Left = left
			}
		case "spacing":
			// 处理间距
			if before, err := strconv.ParseFloat(attr.Value, 64); err == nil {
				props.Spacing.Before = before
			}
		}
	}
}

// parseRunProperties 解析运行属性
func (dsp *DocxStyleParser) parseRunProperties(element xml.StartElement, props *types.StyleProperties) {
	for _, attr := range element.Attr {
		switch attr.Name.Local {
		case "b":
			props.Bold = attr.Value == "true"
		case "i":
			props.Italic = attr.Value == "true"
		case "sz":
			if size, err := strconv.ParseFloat(attr.Value, 64); err == nil {
				props.Size = size
			}
		case "color":
			props.Color.RGB = attr.Value
		}
	}
}

// parseTableProperties 解析表格属性
func (dsp *DocxStyleParser) parseTableProperties(element xml.StartElement, props *types.StyleProperties) {
	// 解析表格属性
	for _, attr := range element.Attr {
		switch attr.Name.Local {
		case "w":
			if width, err := strconv.ParseFloat(attr.Value, 64); err == nil {
				props.TableBorders.Top.Width = width
			}
		}
	}
}

// parseTableCellProperties 解析表格单元格属性
func (dsp *DocxStyleParser) parseTableCellProperties(element xml.StartElement, props *types.StyleProperties) {
	// 解析表格单元格属性
	for _, attr := range element.Attr {
		switch attr.Name.Local {
		case "w":
			if width, err := strconv.ParseFloat(attr.Value, 64); err == nil {
				props.CellPadding.Left = width
			}
		}
	}
}

// getAttribute 获取XML属性值
func (dsp *DocxStyleParser) getAttribute(element xml.StartElement, name string) string {
	for _, attr := range element.Attr {
		if attr.Name.Local == name {
			return attr.Value
		}
	}
	return ""
}

// ParseStyleInheritance 解析样式继承关系
func (dsp *DocxStyleParser) ParseStyleInheritance(filePath string) (map[string]types.StyleInheritance, error) {
	inheritance := make(map[string]types.StyleInheritance)

	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	for _, file := range reader.File {
		if file.Name == "word/styles.xml" {
			rc, err := file.Open()
			if err != nil {
				continue
			}
			defer rc.Close()

			decoder := xml.NewDecoder(rc)
			var currentInheritance types.StyleInheritance
			var currentStyleID string

			for {
				token, err := decoder.Token()
				if err == io.EOF {
					break
				}
				if err != nil {
					continue
				}

				switch t := token.(type) {
				case xml.StartElement:
					switch t.Name.Local {
					case "style":
						currentStyleID = dsp.getAttribute(t, "styleId")
						currentInheritance = types.StyleInheritance{}

					case "w:basedOn":
						currentInheritance.BasedOn = dsp.getAttribute(t, "val")

					case "w:next":
						currentInheritance.Next = dsp.getAttribute(t, "val")

					case "w:link":
						currentInheritance.Linked = dsp.getAttribute(t, "val")

					case "w:qFormat":
						currentInheritance.QuickFormat = true

					case "w:uiPriority":
						if priority, err := strconv.Atoi(dsp.getAttribute(t, "val")); err == nil {
							currentInheritance.Priority = priority
						}
					}

				case xml.EndElement:
					if t.Name.Local == "style" && currentStyleID != "" {
						inheritance[currentStyleID] = currentInheritance
						currentStyleID = ""
					}
				}
			}
		}
	}

	return inheritance, nil
}

// ParseThemeStyles 解析主题样式
func (dsp *DocxStyleParser) ParseThemeStyles(filePath string) (map[string]*types.ThemeStyle, error) {
	themeStyles := make(map[string]*types.ThemeStyle)

	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	for _, file := range reader.File {
		if strings.HasPrefix(file.Name, "word/theme/") && strings.HasSuffix(file.Name, ".xml") {
			rc, err := file.Open()
			if err != nil {
				continue
			}
			defer rc.Close()

			decoder := xml.NewDecoder(rc)
			var currentTheme *types.ThemeStyle

			for {
				token, err := decoder.Token()
				if err == io.EOF {
					break
				}
				if err != nil {
					continue
				}

				switch t := token.(type) {
				case xml.StartElement:
					switch t.Name.Local {
					case "a:theme":
						themeName := strings.TrimSuffix(strings.TrimPrefix(file.Name, "word/theme/"), ".xml")
						currentTheme = &types.ThemeStyle{
							ThemeName:   themeName,
							ColorScheme: "Office",
							FontScheme:  "Office",
							Effects:     "Office",
							Version:     "1.0",
						}

					case "a:clrScheme":
						if currentTheme != nil {
							currentTheme.ColorScheme = dsp.getAttribute(t, "name")
						}

					case "a:fontScheme":
						if currentTheme != nil {
							currentTheme.FontScheme = dsp.getAttribute(t, "name")
						}

					case "a:fmtScheme":
						if currentTheme != nil {
							currentTheme.Effects = dsp.getAttribute(t, "name")
						}
					}

				case xml.EndElement:
					if t.Name.Local == "a:theme" && currentTheme != nil {
						themeStyles[currentTheme.ThemeName] = currentTheme
						currentTheme = nil
					}
				}
			}
		}
	}

	return themeStyles, nil
}

// ParseConditionalStyles 解析条件样式
func (dsp *DocxStyleParser) ParseConditionalStyles(filePath string) ([]types.ConditionalStyle, error) {
	// DOCX格式通常不直接支持条件样式，这里返回空列表
	return []types.ConditionalStyle{}, nil
}

// ValidateStyles 验证样式
func (dsp *DocxStyleParser) ValidateStyles(styles *types.StyleManager) (*types.StyleValidation, error) {
	return dsp.baseParser.ValidateStyles(styles)
}

// ResolveConflicts 解决样式冲突
func (dsp *DocxStyleParser) ResolveConflicts(styles *types.StyleManager) ([]types.StyleConflict, error) {
	conflicts := []types.StyleConflict{}

	// 检查样式冲突
	for styleID, style := range styles.Styles {
		// 检查继承冲突
		if style.Inheritance.BasedOn != "" {
			if _, exists := styles.Styles[style.Inheritance.BasedOn]; !exists {
				conflicts = append(conflicts, types.StyleConflict{
					ID:          fmt.Sprintf("conflict_%s", styleID),
					Conflicting: []string{styleID, style.Inheritance.BasedOn},
					Resolution:  fmt.Sprintf("样式 %s 基于不存在的样式 %s", styleID, style.Inheritance.BasedOn),
					Priority:    1,
					Resolved:    false,
				})
			}
		}

		// 检查下一样式冲突
		if style.Inheritance.Next != "" {
			if _, exists := styles.Styles[style.Inheritance.Next]; !exists {
				conflicts = append(conflicts, types.StyleConflict{
					ID:          fmt.Sprintf("conflict_%s", styleID),
					Conflicting: []string{styleID, style.Inheritance.Next},
					Resolution:  fmt.Sprintf("样式 %s 的下一样式 %s 不存在", styleID, style.Inheritance.Next),
					Priority:    2,
					Resolved:    false,
				})
			}
		}
	}

	styles.Conflicts = conflicts
	return conflicts, nil
}

// GetSupportedStyleTypes 获取支持的样式类型
func (dsp *DocxStyleParser) GetSupportedStyleTypes() []types.StyleType {
	return []types.StyleType{
		types.StyleTypeParagraph,
		types.StyleTypeCharacter,
		types.StyleTypeTable,
		types.StyleTypeList,
		types.StyleTypePage,
		types.StyleTypeSection,
		types.StyleTypeTheme,
	}
}
