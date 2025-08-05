package styles

import (
	"fmt"
	"strings"
	"time"

	"docs-parser/internal/core/types"
)

// Validator 样式验证器
type Validator struct {
	inheritanceProcessor *InheritanceProcessor
}

// NewValidator 创建新的样式验证器
func NewValidator() *Validator {
	return &Validator{
		inheritanceProcessor: NewInheritanceProcessor(),
	}
}

// ValidateStyleManager 验证样式管理器
func (v *Validator) ValidateStyleManager(manager *types.StyleManager) (*types.StyleValidation, error) {
	validation := &types.StyleValidation{
		Valid:       true,
		Errors:      []string{},
		Warnings:    []string{},
		Suggestions: []string{},
		LastChecked: time.Now(),
	}

	// 验证样式完整性
	if err := v.validateStyleCompleteness(manager, validation); err != nil {
		return validation, err
	}

	// 验证继承关系
	if err := v.validateInheritance(manager, validation); err != nil {
		return validation, err
	}

	// 验证样式冲突
	if err := v.validateConflicts(manager, validation); err != nil {
		return validation, err
	}

	// 验证样式一致性
	if err := v.validateConsistency(manager, validation); err != nil {
		return validation, err
	}

	// 验证样式命名
	if err := v.validateNaming(manager, validation); err != nil {
		return validation, err
	}

	// 验证样式属性
	if err := v.validateProperties(manager, validation); err != nil {
		return validation, err
	}

	// 更新验证状态
	validation.Valid = len(validation.Errors) == 0

	return validation, nil
}

// validateStyleCompleteness 验证样式完整性
func (v *Validator) validateStyleCompleteness(manager *types.StyleManager, validation *types.StyleValidation) error {
	if manager == nil {
		validation.Errors = append(validation.Errors, "样式管理器为空")
		return nil
	}

	if len(manager.Styles) == 0 {
		validation.Warnings = append(validation.Warnings, "没有找到任何样式")
		return nil
	}

	// 检查每个样式的完整性
	for styleID, style := range manager.Styles {
		if style == nil {
			validation.Errors = append(validation.Errors, fmt.Sprintf("样式 %s 为空", styleID))
			continue
		}

		// 检查必需字段
		if style.ID == "" {
			validation.Errors = append(validation.Errors, fmt.Sprintf("样式 %s 缺少ID", styleID))
		}

		if style.Name == "" {
			validation.Warnings = append(validation.Warnings, fmt.Sprintf("样式 %s 缺少名称", styleID))
		}

		if style.Type == "" {
			validation.Errors = append(validation.Errors, fmt.Sprintf("样式 %s 缺少类型", styleID))
		}

		// 检查样式类型是否有效
		if !v.isValidStyleType(style.Type) {
			validation.Errors = append(validation.Errors, fmt.Sprintf("样式 %s 的类型 %s 无效", styleID, style.Type))
		}
	}

	return nil
}

// validateInheritance 验证继承关系
func (v *Validator) validateInheritance(manager *types.StyleManager, validation *types.StyleValidation) error {
	// 检查循环引用
	inheritanceErrors := v.inheritanceProcessor.ValidateInheritance(manager.Styles)
	for _, err := range inheritanceErrors {
		validation.Errors = append(validation.Errors, err)
	}

	// 检查继承链长度
	for styleID := range manager.Styles {
		chain := v.inheritanceProcessor.GetInheritanceChain(styleID, manager.Styles)
		if len(chain) > 10 {
			validation.Warnings = append(validation.Warnings, 
				fmt.Sprintf("样式 %s 的继承链过长 (%d 层)", styleID, len(chain)))
		}
	}

	return nil
}

// validateConflicts 验证样式冲突
func (v *Validator) validateConflicts(manager *types.StyleManager, validation *types.StyleValidation) error {
	// 检查命名冲突
	nameMap := make(map[string][]string)
	for styleID, style := range manager.Styles {
		if style.Name != "" {
			nameMap[style.Name] = append(nameMap[style.Name], styleID)
		}
	}

	for name, styleIDs := range nameMap {
		if len(styleIDs) > 1 {
			validation.Warnings = append(validation.Warnings, 
				fmt.Sprintf("样式名称 '%s' 被多个样式使用: %v", name, styleIDs))
		}
	}

	// 检查属性冲突
	for styleID, style := range manager.Styles {
		conflicts := v.detectPropertyConflicts(style)
		for _, conflict := range conflicts {
			validation.Warnings = append(validation.Warnings, 
				fmt.Sprintf("样式 %s: %s", styleID, conflict))
		}
	}

	return nil
}

// validateConsistency 验证样式一致性
func (v *Validator) validateConsistency(manager *types.StyleManager, validation *types.StyleValidation) error {
	// 检查同类型样式的一致性
	typeGroups := make(map[types.StyleType][]*types.AdvancedStyle)
	for _, style := range manager.Styles {
		typeGroups[style.Type] = append(typeGroups[style.Type], style)
	}

	for styleType, styles := range typeGroups {
		if len(styles) > 1 {
			v.checkStyleTypeConsistency(styleType, styles, validation)
		}
	}

	return nil
}

// validateNaming 验证样式命名
func (v *Validator) validateNaming(manager *types.StyleManager, validation *types.StyleValidation) error {
	for styleID, style := range manager.Styles {
		// 检查命名规范
		if style.Name != "" {
			if !v.isValidStyleName(style.Name) {
				validation.Warnings = append(validation.Warnings, 
					fmt.Sprintf("样式 %s 的名称 '%s' 不符合命名规范", styleID, style.Name))
			}

			// 检查名称长度
			if len(style.Name) > 50 {
				validation.Warnings = append(validation.Warnings, 
					fmt.Sprintf("样式 %s 的名称过长 (%d 字符)", styleID, len(style.Name)))
			}
		}

		// 检查ID格式
		if !v.isValidStyleID(styleID) {
			validation.Warnings = append(validation.Warnings, 
				fmt.Sprintf("样式ID '%s' 格式不规范", styleID))
		}
	}

	return nil
}

// validateProperties 验证样式属性
func (v *Validator) validateProperties(manager *types.StyleManager, validation *types.StyleValidation) error {
	for styleID, style := range manager.Styles {
		// 验证字体属性
		v.validateFontProperties(styleID, &style.Properties, validation)

		// 验证段落属性
		v.validateParagraphProperties(styleID, &style.Properties, validation)

		// 验证表格属性
		v.validateTableProperties(styleID, &style.Properties, validation)

		// 验证页面属性
		v.validatePageProperties(styleID, &style.Properties, validation)
	}

	return nil
}

// validateFontProperties 验证字体属性
func (v *Validator) validateFontProperties(styleID string, props *types.StyleProperties, validation *types.StyleValidation) {
	// 检查字体大小
	if props.Size < 0 {
		validation.Errors = append(validation.Errors, 
			fmt.Sprintf("样式 %s 的字体大小不能为负数", styleID))
	}

	if props.Size > 1000 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的字体大小过大 (%f)", styleID, props.Size))
	}

	// 检查颜色值
	if props.Color.RGB != "" && !v.isValidColor(props.Color.RGB) {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的颜色值 '%s' 格式不正确", styleID, props.Color.RGB))
	}
}

// validateParagraphProperties 验证段落属性
func (v *Validator) validateParagraphProperties(styleID string, props *types.StyleProperties, validation *types.StyleValidation) {
	// 检查缩进值
	if props.Indentation.Left < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的左缩进不能为负数", styleID))
	}

	if props.Indentation.Right < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的右缩进不能为负数", styleID))
	}

	// 检查间距值
	if props.Spacing.Before < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的段前间距不能为负数", styleID))
	}

	if props.Spacing.After < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的段后间距不能为负数", styleID))
	}

	if props.Spacing.Line < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的行间距不能为负数", styleID))
	}
}

// validateTableProperties 验证表格属性
func (v *Validator) validateTableProperties(styleID string, props *types.StyleProperties, validation *types.StyleValidation) {
	// 检查单元格内边距
	if props.CellPadding.Top < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的单元格上内边距不能为负数", styleID))
	}

	if props.CellPadding.Bottom < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的单元格下内边距不能为负数", styleID))
	}

	if props.CellPadding.Left < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的单元格左内边距不能为负数", styleID))
	}

	if props.CellPadding.Right < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的单元格右内边距不能为负数", styleID))
	}
}

// validatePageProperties 验证页面属性
func (v *Validator) validatePageProperties(styleID string, props *types.StyleProperties, validation *types.StyleValidation) {
	// 检查页面尺寸
	if props.PageSize.Width <= 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的页面宽度必须大于0", styleID))
	}

	if props.PageSize.Height <= 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的页面高度必须大于0", styleID))
	}

	// 检查页边距
	if props.PageMargins.Top < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的上边距不能为负数", styleID))
	}

	if props.PageMargins.Bottom < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的下边距不能为负数", styleID))
	}

	if props.PageMargins.Left < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的左边距不能为负数", styleID))
	}

	if props.PageMargins.Right < 0 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的右边距不能为负数", styleID))
	}

	// 检查列数
	if props.Columns.Count < 1 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的列数必须大于0", styleID))
	}

	if props.Columns.Count > 10 {
		validation.Warnings = append(validation.Warnings, 
			fmt.Sprintf("样式 %s 的列数过多 (%d)", styleID, props.Columns.Count))
	}
}

// detectPropertyConflicts 检测属性冲突
func (v *Validator) detectPropertyConflicts(style *types.AdvancedStyle) []string {
	var conflicts []string

	// 检查字体冲突
	if style.Properties.Bold && style.Properties.Italic {
		conflicts = append(conflicts, "粗体和斜体同时启用")
	}

	// 检查缩进冲突
	if style.Properties.Indentation.Left > 0 && style.Properties.Indentation.Right > 0 {
		totalIndent := style.Properties.Indentation.Left + style.Properties.Indentation.Right
		if totalIndent > style.Properties.PageSize.Width {
			conflicts = append(conflicts, "左右缩进总和超过页面宽度")
		}
	}

	// 检查间距冲突
	if style.Properties.Spacing.Before > 0 && style.Properties.Spacing.After > 0 {
		totalSpacing := style.Properties.Spacing.Before + style.Properties.Spacing.After
		if totalSpacing > 100 {
			conflicts = append(conflicts, "段前段后间距过大")
		}
	}

	return conflicts
}

// checkStyleTypeConsistency 检查同类型样式的一致性
func (v *Validator) checkStyleTypeConsistency(styleType types.StyleType, styles []*types.AdvancedStyle, validation *types.StyleValidation) {
	// 检查同类型样式的基本属性一致性
	baseProps := make(map[string]interface{})

	for _, style := range styles {
		switch styleType {
		case types.StyleTypeParagraph:
			v.checkParagraphStyleConsistency(style, baseProps, validation)
		case types.StyleTypeCharacter:
			v.checkCharacterStyleConsistency(style, baseProps, validation)
		case types.StyleTypeTable:
			v.checkTableStyleConsistency(style, baseProps, validation)
		}
	}
}

// checkParagraphStyleConsistency 检查段落样式一致性
func (v *Validator) checkParagraphStyleConsistency(style *types.AdvancedStyle, baseProps map[string]interface{}, validation *types.StyleValidation) {
	// 检查段落样式的基本属性
	if style.Properties.Alignment != "" {
		if baseAlignment, exists := baseProps["alignment"]; exists {
			if baseAlignment != style.Properties.Alignment {
				validation.Warnings = append(validation.Warnings, 
					fmt.Sprintf("段落样式 %s 的对齐方式与其他样式不一致", style.ID))
			}
		} else {
			baseProps["alignment"] = style.Properties.Alignment
		}
	}
}

// checkCharacterStyleConsistency 检查字符样式一致性
func (v *Validator) checkCharacterStyleConsistency(style *types.AdvancedStyle, baseProps map[string]interface{}, validation *types.StyleValidation) {
	// 检查字符样式的基本属性
	if style.Properties.Font.Name != "" {
		if baseFont, exists := baseProps["font"]; exists {
			if baseFont != style.Properties.Font.Name {
				validation.Warnings = append(validation.Warnings, 
					fmt.Sprintf("字符样式 %s 的字体与其他样式不一致", style.ID))
			}
		} else {
			baseProps["font"] = style.Properties.Font.Name
		}
	}
}

// checkTableStyleConsistency 检查表格样式一致性
func (v *Validator) checkTableStyleConsistency(style *types.AdvancedStyle, baseProps map[string]interface{}, validation *types.StyleValidation) {
	// 检查表格样式的基本属性
	if style.Properties.TableBorders.Top.Style != "" {
		if baseBorderStyle, exists := baseProps["border_style"]; exists {
			if baseBorderStyle != style.Properties.TableBorders.Top.Style {
				validation.Warnings = append(validation.Warnings, 
					fmt.Sprintf("表格样式 %s 的边框样式与其他样式不一致", style.ID))
			}
		} else {
			baseProps["border_style"] = style.Properties.TableBorders.Top.Style
		}
	}
}

// isValidStyleType 检查样式类型是否有效
func (v *Validator) isValidStyleType(styleType types.StyleType) bool {
	validTypes := []types.StyleType{
		types.StyleTypeParagraph,
		types.StyleTypeCharacter,
		types.StyleTypeTable,
		types.StyleTypeList,
		types.StyleTypePage,
		types.StyleTypeSection,
		types.StyleTypeTheme,
		types.StyleTypeCondition,
	}

	for _, validType := range validTypes {
		if styleType == validType {
			return true
		}
	}

	return false
}

// isValidStyleName 检查样式名称是否有效
func (v *Validator) isValidStyleName(name string) bool {
	// 检查是否包含特殊字符
	invalidChars := []string{"<", ">", ":", "\"", "/", "\\", "|", "?", "*"}
	for _, char := range invalidChars {
		if strings.Contains(name, char) {
			return false
		}
	}

	// 检查是否为空或只包含空白字符
	if strings.TrimSpace(name) == "" {
		return false
	}

	return true
}

// isValidStyleID 检查样式ID是否有效
func (v *Validator) isValidStyleID(id string) bool {
	// 检查是否为空
	if id == "" {
		return false
	}

	// 检查是否包含特殊字符
	invalidChars := []string{"<", ">", ":", "\"", "/", "\\", "|", "?", "*", " "}
	for _, char := range invalidChars {
		if strings.Contains(id, char) {
			return false
		}
	}

	return true
}

// isValidColor 检查颜色值是否有效
func (v *Validator) isValidColor(color string) bool {
	// 检查十六进制颜色格式
	if strings.HasPrefix(color, "#") {
		if len(color) != 7 {
			return false
		}
		// 这里可以添加更详细的十六进制验证
		return true
	}

	// 检查RGB格式
	if strings.HasPrefix(color, "rgb(") && strings.HasSuffix(color, ")") {
		// 这里可以添加RGB格式验证
		return true
	}

	return false
} 