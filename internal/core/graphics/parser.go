package graphics

import (
	"docs-parser/internal/core/types"
	"fmt"
)

// Parser 图形解析器接口
type Parser interface {
	// ParseGraphics 解析文档中的图形元素
	ParseGraphics(documentPath string) (*types.DocumentGraphics, error)

	// ParseImage 解析图片元素
	ParseImage(imagePath string, metadata map[string]interface{}) (*types.GraphicElement, error)

	// ParseShape 解析形状元素
	ParseShape(shapeData []byte, metadata map[string]interface{}) (*types.GraphicElement, error)

	// ParseChart 解析图表元素
	ParseChart(chartData []byte, metadata map[string]interface{}) (*types.GraphicElement, error)

	// ParseSmartArt 解析SmartArt元素
	ParseSmartArt(smartArtData []byte, metadata map[string]interface{}) (*types.GraphicElement, error)

	// ParseTextbox 解析文本框元素
	ParseTextbox(textboxData []byte, metadata map[string]interface{}) (*types.GraphicElement, error)

	// ParseFormula 解析公式元素
	ParseFormula(formulaData []byte, metadata map[string]interface{}) (*types.GraphicElement, error)

	// ExtractImageData 提取图片数据
	ExtractImageData(imagePath string) (*types.ImageData, error)

	// ValidateGraphicElement 验证图形元素
	ValidateGraphicElement(element *types.GraphicElement) error
}

// DefaultParser 默认图形解析器实现
type DefaultParser struct {
	imageParser    ImageParser
	shapeParser    ShapeParser
	chartParser    ChartParser
	smartArtParser SmartArtParser
	textboxParser  TextboxParser
	formulaParser  FormulaParser
}

// NewDefaultParser 创建默认图形解析器
func NewDefaultParser() *DefaultParser {
	return &DefaultParser{
		imageParser:    NewImageParser(),
		shapeParser:    NewShapeParser(),
		chartParser:    NewChartParser(),
		smartArtParser: NewSmartArtParser(),
		textboxParser:  NewTextboxParser(),
		formulaParser:  NewFormulaParser(),
	}
}

// ParseGraphics 解析文档中的图形元素
func (dp *DefaultParser) ParseGraphics(documentPath string) (*types.DocumentGraphics, error) {
	// 基础实现，具体逻辑由子解析器处理
	graphics := &types.DocumentGraphics{
		Elements: []*types.GraphicElement{},
		Count:    0,
		Groups:   []types.GraphicGroup{},
	}

	return graphics, nil
}

// ParseImage 解析图片元素
func (dp *DefaultParser) ParseImage(imagePath string, metadata map[string]interface{}) (*types.GraphicElement, error) {
	return dp.imageParser.Parse(imagePath, metadata)
}

// ParseShape 解析形状元素
func (dp *DefaultParser) ParseShape(shapeData []byte, metadata map[string]interface{}) (*types.GraphicElement, error) {
	return dp.shapeParser.Parse(shapeData, metadata)
}

// ParseChart 解析图表元素
func (dp *DefaultParser) ParseChart(chartData []byte, metadata map[string]interface{}) (*types.GraphicElement, error) {
	return dp.chartParser.Parse(chartData, metadata)
}

// ParseSmartArt 解析SmartArt元素
func (dp *DefaultParser) ParseSmartArt(smartArtData []byte, metadata map[string]interface{}) (*types.GraphicElement, error) {
	return dp.smartArtParser.Parse(smartArtData, metadata)
}

// ParseTextbox 解析文本框元素
func (dp *DefaultParser) ParseTextbox(textboxData []byte, metadata map[string]interface{}) (*types.GraphicElement, error) {
	return dp.textboxParser.Parse(textboxData, metadata)
}

// ParseFormula 解析公式元素
func (dp *DefaultParser) ParseFormula(formulaData []byte, metadata map[string]interface{}) (*types.GraphicElement, error) {
	return dp.formulaParser.Parse(formulaData, metadata)
}

// ExtractImageData 提取图片数据
func (dp *DefaultParser) ExtractImageData(imagePath string) (*types.ImageData, error) {
	return dp.imageParser.ExtractData(imagePath)
}

// ValidateGraphicElement 验证图形元素
func (dp *DefaultParser) ValidateGraphicElement(element *types.GraphicElement) error {
	if element == nil {
		return fmt.Errorf("graphic element is nil")
	}

	if element.ID == "" {
		return fmt.Errorf("graphic element ID is required")
	}

	if element.Type == "" {
		return fmt.Errorf("graphic element type is required")
	}

	// 验证位置信息
	if element.Position.X < 0 || element.Position.Y < 0 {
		return fmt.Errorf("invalid position: x=%f, y=%f", element.Position.X, element.Position.Y)
	}

	// 验证尺寸信息
	if element.Size.Width <= 0 || element.Size.Height <= 0 {
		return fmt.Errorf("invalid size: width=%f, height=%f", element.Size.Width, element.Size.Height)
	}

	return nil
}

// ImageParser 图片解析器接口
type ImageParser interface {
	Parse(imagePath string, metadata map[string]interface{}) (*types.GraphicElement, error)
	ExtractData(imagePath string) (*types.ImageData, error)
}

// ShapeParser 形状解析器接口
type ShapeParser interface {
	Parse(shapeData []byte, metadata map[string]interface{}) (*types.GraphicElement, error)
}

// ChartParser 图表解析器接口
type ChartParser interface {
	Parse(chartData []byte, metadata map[string]interface{}) (*types.GraphicElement, error)
}

// SmartArtParser SmartArt解析器接口
type SmartArtParser interface {
	Parse(smartArtData []byte, metadata map[string]interface{}) (*types.GraphicElement, error)
}

// TextboxParser 文本框解析器接口
type TextboxParser interface {
	Parse(textboxData []byte, metadata map[string]interface{}) (*types.GraphicElement, error)
}

// FormulaParser 公式解析器接口
type FormulaParser interface {
	Parse(formulaData []byte, metadata map[string]interface{}) (*types.GraphicElement, error)
}

// 基础解析器实现
type defaultImageParser struct{}
type defaultShapeParser struct{}
type defaultChartParser struct{}
type defaultSmartArtParser struct{}
type defaultTextboxParser struct{}
type defaultFormulaParser struct{}

// NewImageParser 创建图片解析器
func NewImageParser() ImageParser {
	return &defaultImageParser{}
}

// NewShapeParser 创建形状解析器
func NewShapeParser() ShapeParser {
	return &defaultShapeParser{}
}

// NewChartParser 创建图表解析器
func NewChartParser() ChartParser {
	return &defaultChartParser{}
}

// NewSmartArtParser 创建SmartArt解析器
func NewSmartArtParser() SmartArtParser {
	return &defaultSmartArtParser{}
}

// NewTextboxParser 创建文本框解析器
func NewTextboxParser() TextboxParser {
	return &defaultTextboxParser{}
}

// NewFormulaParser 创建公式解析器
func NewFormulaParser() FormulaParser {
	return &defaultFormulaParser{}
}

// 基础实现方法
func (dip *defaultImageParser) Parse(imagePath string, metadata map[string]interface{}) (*types.GraphicElement, error) {
	// 基础实现
	return &types.GraphicElement{
		Type: types.GraphicTypeImage,
		ID:   fmt.Sprintf("image_%s", imagePath),
		Position: types.GraphicPosition{
			X:         0,
			Y:         0,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           100,
			Height:          100,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: true,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
		},
		Content:  types.GraphicContent{},
		Metadata: types.GraphicMetadata{},
		Anchor:   types.Anchor{},
		ZIndex:   0,
		Visible:  true,
		Locked:   false,
	}, nil
}

func (dip *defaultImageParser) ExtractData(imagePath string) (*types.ImageData, error) {
	// 基础实现
	return &types.ImageData{
		Source: imagePath,
		Format: "unknown",
	}, nil
}

func (dsp *defaultShapeParser) Parse(shapeData []byte, metadata map[string]interface{}) (*types.GraphicElement, error) {
	// 基础实现
	return &types.GraphicElement{
		Type: types.GraphicTypeShape,
		ID:   "shape_1",
		Position: types.GraphicPosition{
			X:         0,
			Y:         0,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           100,
			Height:          100,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: true,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
		},
		Content:  types.GraphicContent{},
		Metadata: types.GraphicMetadata{},
		Anchor:   types.Anchor{},
		ZIndex:   0,
		Visible:  true,
		Locked:   false,
	}, nil
}

func (dcp *defaultChartParser) Parse(chartData []byte, metadata map[string]interface{}) (*types.GraphicElement, error) {
	// 基础实现
	return &types.GraphicElement{
		Type: types.GraphicTypeChart,
		ID:   "chart_1",
		Position: types.GraphicPosition{
			X:         0,
			Y:         0,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           100,
			Height:          100,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: true,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
		},
		Content:  types.GraphicContent{},
		Metadata: types.GraphicMetadata{},
		Anchor:   types.Anchor{},
		ZIndex:   0,
		Visible:  true,
		Locked:   false,
	}, nil
}

func (dsap *defaultSmartArtParser) Parse(smartArtData []byte, metadata map[string]interface{}) (*types.GraphicElement, error) {
	// 基础实现
	return &types.GraphicElement{
		Type: types.GraphicTypeSmartArt,
		ID:   "smartart_1",
		Position: types.GraphicPosition{
			X:         0,
			Y:         0,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           100,
			Height:          100,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: true,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
		},
		Content:  types.GraphicContent{},
		Metadata: types.GraphicMetadata{},
		Anchor:   types.Anchor{},
		ZIndex:   0,
		Visible:  true,
		Locked:   false,
	}, nil
}

func (dtp *defaultTextboxParser) Parse(textboxData []byte, metadata map[string]interface{}) (*types.GraphicElement, error) {
	// 基础实现
	return &types.GraphicElement{
		Type: types.GraphicTypeTextbox,
		ID:   "textbox_1",
		Position: types.GraphicPosition{
			X:         0,
			Y:         0,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           100,
			Height:          100,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: true,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
		},
		Content:  types.GraphicContent{},
		Metadata: types.GraphicMetadata{},
		Anchor:   types.Anchor{},
		ZIndex:   0,
		Visible:  true,
		Locked:   false,
	}, nil
}

func (dfp *defaultFormulaParser) Parse(formulaData []byte, metadata map[string]interface{}) (*types.GraphicElement, error) {
	// 基础实现
	return &types.GraphicElement{
		Type: types.GraphicTypeFormula,
		ID:   "formula_1",
		Position: types.GraphicPosition{
			X:         0,
			Y:         0,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           100,
			Height:          100,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: true,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
		},
		Content:  types.GraphicContent{},
		Metadata: types.GraphicMetadata{},
		Anchor:   types.Anchor{},
		ZIndex:   0,
		Visible:  true,
		Locked:   false,
	}, nil
}
