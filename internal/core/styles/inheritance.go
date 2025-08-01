package styles

import (
	"fmt"

	"docs-parser/internal/core/types"
)

// InheritanceProcessor 样式继承处理器
type InheritanceProcessor struct{}

// NewInheritanceProcessor 创建新的继承处理器
func NewInheritanceProcessor() *InheritanceProcessor {
	return &InheritanceProcessor{}
}

// BuildInheritanceTree 构建样式继承树
func (ip *InheritanceProcessor) BuildInheritanceTree(styles map[string]*types.AdvancedStyle) map[string][]string {
	tree := make(map[string][]string)
	
	// 构建继承关系
	for styleID, style := range styles {
		if style.Inheritance.BasedOn != "" {
			parentID := style.Inheritance.BasedOn
			tree[parentID] = append(tree[parentID], styleID)
		}
	}
	
	return tree
}

// ResolveInheritance 解析样式继承
func (ip *InheritanceProcessor) ResolveInheritance(styles map[string]*types.AdvancedStyle) error {
	// 创建样式副本以避免循环引用
	resolvedStyles := make(map[string]*types.AdvancedStyle)
	
	// 首先复制所有样式
	for id, style := range styles {
		resolvedStyles[id] = ip.cloneStyle(style)
	}
	
	// 解析继承关系
	for styleID, style := range resolvedStyles {
		if err := ip.resolveStyleInheritance(styleID, style, resolvedStyles); err != nil {
			return fmt.Errorf("解析样式 %s 的继承关系失败: %w", styleID, err)
		}
	}
	
	// 更新原始样式
	for id, resolvedStyle := range resolvedStyles {
		styles[id] = resolvedStyle
	}
	
	return nil
}

// resolveStyleInheritance 解析单个样式的继承关系
func (ip *InheritanceProcessor) resolveStyleInheritance(styleID string, style *types.AdvancedStyle, allStyles map[string]*types.AdvancedStyle) error {
	// 如果已经解析过，直接返回
	if style.Properties.Name != "" && style.Properties.Name != style.Name {
		return nil
	}
	
	// 解析基于样式
	if style.Inheritance.BasedOn != "" {
		parentStyle, exists := allStyles[style.Inheritance.BasedOn]
		if !exists {
			return fmt.Errorf("基于样式 %s 不存在", style.Inheritance.BasedOn)
		}
		
		// 递归解析父样式
		if err := ip.resolveStyleInheritance(style.Inheritance.BasedOn, parentStyle, allStyles); err != nil {
			return err
		}
		
		// 合并父样式属性
		ip.mergeStyleProperties(&style.Properties, &parentStyle.Properties)
	}
	
	// 解析链接样式
	if style.Inheritance.Linked != "" {
		linkedStyle, exists := allStyles[style.Inheritance.Linked]
		if exists {
			// 递归解析链接样式
			if err := ip.resolveStyleInheritance(style.Inheritance.Linked, linkedStyle, allStyles); err != nil {
				return err
			}
			
			// 合并链接样式属性
			ip.mergeStyleProperties(&style.Properties, &linkedStyle.Properties)
		}
	}
	
	return nil
}

// mergeStyleProperties 合并样式属性
func (ip *InheritanceProcessor) mergeStyleProperties(target, source *types.StyleProperties) {
	// 合并字体属性
	if target.Font.Name == "" && source.Font.Name != "" {
		target.Font.Name = source.Font.Name
	}
	if target.Size == 0 && source.Size > 0 {
		target.Size = source.Size
	}
	if target.Color.RGB == "" && source.Color.RGB != "" {
		target.Color = source.Color
	}
	if !target.Bold && source.Bold {
		target.Bold = source.Bold
	}
	if !target.Italic && source.Italic {
		target.Italic = source.Italic
	}
	if target.Underline == "" && source.Underline != "" {
		target.Underline = source.Underline
	}
	if target.Highlight == "" && source.Highlight != "" {
		target.Highlight = source.Highlight
	}
	
	// 合并段落属性
	if target.Alignment == "" && source.Alignment != "" {
		target.Alignment = source.Alignment
	}
	if target.Indentation.Left == 0 && source.Indentation.Left > 0 {
		target.Indentation.Left = source.Indentation.Left
	}
	if target.Indentation.Right == 0 && source.Indentation.Right > 0 {
		target.Indentation.Right = source.Indentation.Right
	}
	if target.Indentation.First == 0 && source.Indentation.First > 0 {
		target.Indentation.First = source.Indentation.First
	}
	if target.Indentation.Hanging == 0 && source.Indentation.Hanging > 0 {
		target.Indentation.Hanging = source.Indentation.Hanging
	}
	if target.Spacing.Before == 0 && source.Spacing.Before > 0 {
		target.Spacing.Before = source.Spacing.Before
	}
	if target.Spacing.After == 0 && source.Spacing.After > 0 {
		target.Spacing.After = source.Spacing.After
	}
	if target.Spacing.Line == 0 && source.Spacing.Line > 0 {
		target.Spacing.Line = source.Spacing.Line
	}
	
	// 合并边框属性
	ip.mergeBorders(&target.Borders, &source.Borders)
	
	// 合并底纹属性
	if target.Shading.Fill.RGB == "" && source.Shading.Fill.RGB != "" {
		target.Shading.Fill = source.Shading.Fill
	}
	if target.Shading.Pattern == "" && source.Shading.Pattern != "" {
		target.Shading.Pattern = source.Shading.Pattern
	}
	
	// 合并高级属性
	if target.Position == "" && source.Position != "" {
		target.Position = source.Position
	}
	if target.Rotation == 0 && source.Rotation > 0 {
		target.Rotation = source.Rotation
	}
	if target.Scale == 0 && source.Scale > 0 {
		target.Scale = source.Scale
	}
	if target.Opacity == 0 && source.Opacity > 0 {
		target.Opacity = source.Opacity
	}
	
	// 合并列表属性
	if target.ListType == "" && source.ListType != "" {
		target.ListType = source.ListType
	}
	if target.ListLevel == 0 && source.ListLevel > 0 {
		target.ListLevel = source.ListLevel
	}
	
	// 合并表格属性
	ip.mergeTableBorders(&target.TableBorders, &source.TableBorders)
	ip.mergeTableShading(&target.TableShading, &source.TableShading)
	ip.mergeCellPadding(&target.CellPadding, &source.CellPadding)
	
	// 合并页面属性
	if target.PageSize.Width == 0 && source.PageSize.Width > 0 {
		target.PageSize.Width = source.PageSize.Width
	}
	if target.PageSize.Height == 0 && source.PageSize.Height > 0 {
		target.PageSize.Height = source.PageSize.Height
	}
	
	// 合并页边距
	if target.PageMargins.Top == 0 && source.PageMargins.Top > 0 {
		target.PageMargins.Top = source.PageMargins.Top
	}
	if target.PageMargins.Bottom == 0 && source.PageMargins.Bottom > 0 {
		target.PageMargins.Bottom = source.PageMargins.Bottom
	}
	if target.PageMargins.Left == 0 && source.PageMargins.Left > 0 {
		target.PageMargins.Left = source.PageMargins.Left
	}
	if target.PageMargins.Right == 0 && source.PageMargins.Right > 0 {
		target.PageMargins.Right = source.PageMargins.Right
	}
	if target.PageMargins.Header == 0 && source.PageMargins.Header > 0 {
		target.PageMargins.Header = source.PageMargins.Header
	}
	if target.PageMargins.Footer == 0 && source.PageMargins.Footer > 0 {
		target.PageMargins.Footer = source.PageMargins.Footer
	}
	
	// 合并列属性
	if target.Columns.Count == 0 && source.Columns.Count > 0 {
		target.Columns.Count = source.Columns.Count
	}
	if target.Columns.Spacing == 0 && source.Columns.Spacing > 0 {
		target.Columns.Spacing = source.Columns.Spacing
	}
	
	// 合并节属性
	if target.SectionType == "" && source.SectionType != "" {
		target.SectionType = source.SectionType
	}
}

// mergeBorders 合并边框属性
func (ip *InheritanceProcessor) mergeBorders(target, source *types.Borders) {
	// 合并顶部边框
	if target.Top.Style == "" && source.Top.Style != "" {
		target.Top.Style = source.Top.Style
	}
	if target.Top.Width == 0 && source.Top.Width > 0 {
		target.Top.Width = source.Top.Width
	}
	if target.Top.Color.RGB == "" && source.Top.Color.RGB != "" {
		target.Top.Color = source.Top.Color
	}
	if target.Top.Space == 0 && source.Top.Space > 0 {
		target.Top.Space = source.Top.Space
	}
	
	// 合并底部边框
	if target.Bottom.Style == "" && source.Bottom.Style != "" {
		target.Bottom.Style = source.Bottom.Style
	}
	if target.Bottom.Width == 0 && source.Bottom.Width > 0 {
		target.Bottom.Width = source.Bottom.Width
	}
	if target.Bottom.Color.RGB == "" && source.Bottom.Color.RGB != "" {
		target.Bottom.Color = source.Bottom.Color
	}
	if target.Bottom.Space == 0 && source.Bottom.Space > 0 {
		target.Bottom.Space = source.Bottom.Space
	}
	
	// 合并左侧边框
	if target.Left.Style == "" && source.Left.Style != "" {
		target.Left.Style = source.Left.Style
	}
	if target.Left.Width == 0 && source.Left.Width > 0 {
		target.Left.Width = source.Left.Width
	}
	if target.Left.Color.RGB == "" && source.Left.Color.RGB != "" {
		target.Left.Color = source.Left.Color
	}
	if target.Left.Space == 0 && source.Left.Space > 0 {
		target.Left.Space = source.Left.Space
	}
	
	// 合并右侧边框
	if target.Right.Style == "" && source.Right.Style != "" {
		target.Right.Style = source.Right.Style
	}
	if target.Right.Width == 0 && source.Right.Width > 0 {
		target.Right.Width = source.Right.Width
	}
	if target.Right.Color.RGB == "" && source.Right.Color.RGB != "" {
		target.Right.Color = source.Right.Color
	}
	if target.Right.Space == 0 && source.Right.Space > 0 {
		target.Right.Space = source.Right.Space
	}
}

// mergeTableBorders 合并表格边框属性
func (ip *InheritanceProcessor) mergeTableBorders(target, source *types.TableBorders) {
	// 合并顶部边框
	if target.Top.Style == "" && source.Top.Style != "" {
		target.Top.Style = source.Top.Style
	}
	if target.Top.Width == 0 && source.Top.Width > 0 {
		target.Top.Width = source.Top.Width
	}
	if target.Top.Color.RGB == "" && source.Top.Color.RGB != "" {
		target.Top.Color = source.Top.Color
	}
	if target.Top.Space == 0 && source.Top.Space > 0 {
		target.Top.Space = source.Top.Space
	}
	
	// 合并底部边框
	if target.Bottom.Style == "" && source.Bottom.Style != "" {
		target.Bottom.Style = source.Bottom.Style
	}
	if target.Bottom.Width == 0 && source.Bottom.Width > 0 {
		target.Bottom.Width = source.Bottom.Width
	}
	if target.Bottom.Color.RGB == "" && source.Bottom.Color.RGB != "" {
		target.Bottom.Color = source.Bottom.Color
	}
	if target.Bottom.Space == 0 && source.Bottom.Space > 0 {
		target.Bottom.Space = source.Bottom.Space
	}
	
	// 合并左侧边框
	if target.Left.Style == "" && source.Left.Style != "" {
		target.Left.Style = source.Left.Style
	}
	if target.Left.Width == 0 && source.Left.Width > 0 {
		target.Left.Width = source.Left.Width
	}
	if target.Left.Color.RGB == "" && source.Left.Color.RGB != "" {
		target.Left.Color = source.Left.Color
	}
	if target.Left.Space == 0 && source.Left.Space > 0 {
		target.Left.Space = source.Left.Space
	}
	
	// 合并右侧边框
	if target.Right.Style == "" && source.Right.Style != "" {
		target.Right.Style = source.Right.Style
	}
	if target.Right.Width == 0 && source.Right.Width > 0 {
		target.Right.Width = source.Right.Width
	}
	if target.Right.Color.RGB == "" && source.Right.Color.RGB != "" {
		target.Right.Color = source.Right.Color
	}
	if target.Right.Space == 0 && source.Right.Space > 0 {
		target.Right.Space = source.Right.Space
	}
	
	// 合并内部水平边框
	if target.InsideH.Style == "" && source.InsideH.Style != "" {
		target.InsideH.Style = source.InsideH.Style
	}
	if target.InsideH.Width == 0 && source.InsideH.Width > 0 {
		target.InsideH.Width = source.InsideH.Width
	}
	if target.InsideH.Color.RGB == "" && source.InsideH.Color.RGB != "" {
		target.InsideH.Color = source.InsideH.Color
	}
	if target.InsideH.Space == 0 && source.InsideH.Space > 0 {
		target.InsideH.Space = source.InsideH.Space
	}
	
	// 合并内部垂直边框
	if target.InsideV.Style == "" && source.InsideV.Style != "" {
		target.InsideV.Style = source.InsideV.Style
	}
	if target.InsideV.Width == 0 && source.InsideV.Width > 0 {
		target.InsideV.Width = source.InsideV.Width
	}
	if target.InsideV.Color.RGB == "" && source.InsideV.Color.RGB != "" {
		target.InsideV.Color = source.InsideV.Color
	}
	if target.InsideV.Space == 0 && source.InsideV.Space > 0 {
		target.InsideV.Space = source.InsideV.Space
	}
}

// mergeTableShading 合并表格底纹属性
func (ip *InheritanceProcessor) mergeTableShading(target, source *types.TableShading) {
	if target.Fill.RGB == "" && source.Fill.RGB != "" {
		target.Fill = source.Fill
	}
	if target.Pattern == "" && source.Pattern != "" {
		target.Pattern = source.Pattern
	}
}

// mergeCellPadding 合并单元格内边距属性
func (ip *InheritanceProcessor) mergeCellPadding(target, source *types.CellPadding) {
	if target.Top == 0 && source.Top > 0 {
		target.Top = source.Top
	}
	if target.Bottom == 0 && source.Bottom > 0 {
		target.Bottom = source.Bottom
	}
	if target.Left == 0 && source.Left > 0 {
		target.Left = source.Left
	}
	if target.Right == 0 && source.Right > 0 {
		target.Right = source.Right
	}
}

// cloneStyle 克隆样式
func (ip *InheritanceProcessor) cloneStyle(style *types.AdvancedStyle) *types.AdvancedStyle {
	if style == nil {
		return nil
	}
	
	cloned := &types.AdvancedStyle{
		ID:          style.ID,
		Name:        style.Name,
		Type:        style.Type,
		Properties:  style.Properties,
		Inheritance: style.Inheritance,
		Theme:       style.Theme,
		Conditions:  make([]types.ConditionalStyle, len(style.Conditions)),
		Conflicts:   make([]types.StyleConflict, len(style.Conflicts)),
		Validation:  style.Validation,
		Created:     style.Created,
		Modified:    style.Modified,
		Version:     style.Version,
	}
	
	// 复制条件样式
	copy(cloned.Conditions, style.Conditions)
	
	// 复制冲突
	copy(cloned.Conflicts, style.Conflicts)
	
	return cloned
}

// ValidateInheritance 验证继承关系
func (ip *InheritanceProcessor) ValidateInheritance(styles map[string]*types.AdvancedStyle) []string {
	var errors []string
	
	// 检查循环引用
	for styleID, style := range styles {
		if err := ip.checkCircularReference(styleID, style, styles, make(map[string]bool)); err != nil {
			errors = append(errors, err.Error())
		}
	}
	
	// 检查无效的基于样式
	for styleID, style := range styles {
		if style.Inheritance.BasedOn != "" {
			if _, exists := styles[style.Inheritance.BasedOn]; !exists {
				errors = append(errors, fmt.Sprintf("样式 %s 基于不存在的样式 %s", styleID, style.Inheritance.BasedOn))
			}
		}
		
		if style.Inheritance.Next != "" {
			if _, exists := styles[style.Inheritance.Next]; !exists {
				errors = append(errors, fmt.Sprintf("样式 %s 的下一样式 %s 不存在", styleID, style.Inheritance.Next))
			}
		}
		
		if style.Inheritance.Linked != "" {
			if _, exists := styles[style.Inheritance.Linked]; !exists {
				errors = append(errors, fmt.Sprintf("样式 %s 的链接样式 %s 不存在", styleID, style.Inheritance.Linked))
			}
		}
	}
	
	return errors
}

// checkCircularReference 检查循环引用
func (ip *InheritanceProcessor) checkCircularReference(styleID string, style *types.AdvancedStyle, allStyles map[string]*types.AdvancedStyle, visited map[string]bool) error {
	if visited[styleID] {
		return fmt.Errorf("检测到循环引用: %s", styleID)
	}
	
	visited[styleID] = true
	defer delete(visited, styleID)
	
	if style.Inheritance.BasedOn != "" {
		parentStyle, exists := allStyles[style.Inheritance.BasedOn]
		if !exists {
			return fmt.Errorf("基于样式 %s 不存在", style.Inheritance.BasedOn)
		}
		
		return ip.checkCircularReference(style.Inheritance.BasedOn, parentStyle, allStyles, visited)
	}
	
	return nil
}

// GetInheritanceChain 获取继承链
func (ip *InheritanceProcessor) GetInheritanceChain(styleID string, styles map[string]*types.AdvancedStyle) []string {
	var chain []string
	currentID := styleID
	
	for currentID != "" {
		chain = append(chain, currentID)
		style, exists := styles[currentID]
		if !exists {
			break
		}
		currentID = style.Inheritance.BasedOn
	}
	
	return chain
}

// GetDescendants 获取后代样式
func (ip *InheritanceProcessor) GetDescendants(styleID string, inheritanceTree map[string][]string) []string {
	var descendants []string
	
	children, exists := inheritanceTree[styleID]
	if !exists {
		return descendants
	}
	
	for _, childID := range children {
		descendants = append(descendants, childID)
		// 递归获取后代
		childDescendants := ip.GetDescendants(childID, inheritanceTree)
		descendants = append(descendants, childDescendants...)
	}
	
	return descendants
}

// GetInheritanceLevel 获取继承层级
func (ip *InheritanceProcessor) GetInheritanceLevel(styleID string, styles map[string]*types.AdvancedStyle) int {
	level := 0
	currentID := styleID
	
	for currentID != "" {
		style, exists := styles[currentID]
		if !exists {
			break
		}
		currentID = style.Inheritance.BasedOn
		level++
	}
	
	return level
}

// SortByInheritanceLevel 按继承层级排序
func (ip *InheritanceProcessor) SortByInheritanceLevel(styleIDs []string, styles map[string]*types.AdvancedStyle) []string {
	// 创建样式ID和层级的映射
	levelMap := make(map[string]int)
	for _, styleID := range styleIDs {
		levelMap[styleID] = ip.GetInheritanceLevel(styleID, styles)
	}
	
	// 按层级排序
	sorted := make([]string, len(styleIDs))
	copy(sorted, styleIDs)
	
	// 简单的冒泡排序
	for i := 0; i < len(sorted)-1; i++ {
		for j := 0; j < len(sorted)-i-1; j++ {
			if levelMap[sorted[j]] > levelMap[sorted[j+1]] {
				sorted[j], sorted[j+1] = sorted[j+1], sorted[j]
			}
		}
	}
	
	return sorted
} 