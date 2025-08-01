package graphics

import (
	"archive/zip"
	"bytes"
	"docs-parser/internal/core/types"
	"encoding/xml"
	"fmt"
	"io"
	"path/filepath"
	"strconv"
	"strings"
)

// DOCXGraphicsParser DOCX图形解析器
type DOCXGraphicsParser struct {
	parser Parser
}

// NewDOCXGraphicsParser 创建DOCX图形解析器
func NewDOCXGraphicsParser() *DOCXGraphicsParser {
	return &DOCXGraphicsParser{
		parser: NewDefaultParser(),
	}
}

// ParseGraphics 解析DOCX文档中的图形元素
func (dgp *DOCXGraphicsParser) ParseGraphics(docxPath string) (*types.DocumentGraphics, error) {
	// 打开DOCX文件
	reader, err := zip.OpenReader(docxPath)
	if err != nil {
		return nil, fmt.Errorf("failed to open DOCX file: %w", err)
	}
	defer reader.Close()

	graphics := &types.DocumentGraphics{
		Elements: []*types.GraphicElement{},
		Count:    0,
		Groups:   []types.GraphicGroup{},
	}

	// 解析图片
	images, err := dgp.parseImages(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse images: %w", err)
	}
	graphics.Elements = append(graphics.Elements, images...)

	// 解析形状
	shapes, err := dgp.parseShapes(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse shapes: %w", err)
	}
	graphics.Elements = append(graphics.Elements, shapes...)

	// 解析图表
	charts, err := dgp.parseCharts(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse charts: %w", err)
	}
	graphics.Elements = append(graphics.Elements, charts...)

	// 解析SmartArt
	smartArts, err := dgp.parseSmartArts(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse SmartArt: %w", err)
	}
	graphics.Elements = append(graphics.Elements, smartArts...)

	// 解析文本框
	textboxes, err := dgp.parseTextboxes(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse textboxes: %w", err)
	}
	graphics.Elements = append(graphics.Elements, textboxes...)

	// 解析公式
	formulas, err := dgp.parseFormulas(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse formulas: %w", err)
	}
	graphics.Elements = append(graphics.Elements, formulas...)

	// 解析图形组
	groups, err := dgp.parseGroups(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse groups: %w", err)
	}
	graphics.Groups = groups

	graphics.Count = len(graphics.Elements)
	return graphics, nil
}

// parseImages 解析图片元素
func (dgp *DOCXGraphicsParser) parseImages(reader *zip.ReadCloser) ([]*types.GraphicElement, error) {
	var images []*types.GraphicElement

	// 遍历media目录中的图片文件
	for _, file := range reader.File {
		if strings.HasPrefix(file.Name, "word/media/") {
			imageElement, err := dgp.parseImageFile(file)
			if err != nil {
				continue // 跳过无法解析的图片
			}
			images = append(images, imageElement)
		}
	}

	return images, nil
}

// parseImageFile 解析单个图片文件
func (dgp *DOCXGraphicsParser) parseImageFile(file *zip.File) (*types.GraphicElement, error) {
	// 读取图片数据
	rc, err := file.Open()
	if err != nil {
		return nil, fmt.Errorf("failed to open image file: %w", err)
	}
	defer rc.Close()

	imageData, err := io.ReadAll(rc)
	if err != nil {
		return nil, fmt.Errorf("failed to read image data: %w", err)
	}

	// 获取图片格式
	format := dgp.getImageFormat(file.Name)

	// 创建图片数据
	imgData := &types.ImageData{
		Source:     file.Name,
		Format:     format,
		Data:       imageData,
		Compressed: false,
	}

	// 创建图形元素
	element := &types.GraphicElement{
		ID:   fmt.Sprintf("image_%s", filepath.Base(file.Name)),
		Type: types.GraphicTypeImage,
		Content: types.GraphicContent{
			Image: *imgData,
		},
		Visible: true,
		Locked:  false,
	}

	// 尝试从文档中获取图片的位置和尺寸信息
	dgp.extractImagePositionAndSize(element, file.Name)

	return element, nil
}

// getImageFormat 获取图片格式
func (dgp *DOCXGraphicsParser) getImageFormat(fileName string) string {
	ext := strings.ToLower(filepath.Ext(fileName))
	switch ext {
	case ".jpg", ".jpeg":
		return "jpeg"
	case ".png":
		return "png"
	case ".gif":
		return "gif"
	case ".bmp":
		return "bmp"
	case ".tiff", ".tif":
		return "tiff"
	case ".webp":
		return "webp"
	default:
		return "unknown"
	}
}

// extractImagePositionAndSize 提取图片位置和尺寸信息
func (dgp *DOCXGraphicsParser) extractImagePositionAndSize(element *types.GraphicElement, imagePath string) {
	// 这里需要解析文档中的drawing.xml来获取图片的位置和尺寸
	// 基础实现，设置默认值
	element.Position = types.GraphicPosition{
		X:         0,
		Y:         0,
		RelativeX: 0,
		RelativeY: 0,
		Unit:      "emu",
	}
	element.Size = types.Size{
		Width:           100,
		Height:          100,
		ScaleX:          1.0,
		ScaleY:          1.0,
		Unit:            "emu",
		LockAspectRatio: true,
	}
}

// parseShapes 解析形状元素
func (dgp *DOCXGraphicsParser) parseShapes(reader *zip.ReadCloser) ([]*types.GraphicElement, error) {
	var shapes []*types.GraphicElement

	// 查找drawing.xml文件
	for _, file := range reader.File {
		if file.Name == "word/drawing.xml" {
			shapeElements, err := dgp.parseDrawingXML(file)
			if err != nil {
				continue
			}
			shapes = append(shapes, shapeElements...)
		}
	}

	return shapes, nil
}

// parseDrawingXML 解析drawing.xml文件
func (dgp *DOCXGraphicsParser) parseDrawingXML(file *zip.File) ([]*types.GraphicElement, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, fmt.Errorf("failed to open drawing.xml: %w", err)
	}
	defer rc.Close()

	data, err := io.ReadAll(rc)
	if err != nil {
		return nil, fmt.Errorf("failed to read drawing.xml: %w", err)
	}

	var shapes []*types.GraphicElement

	// 解析XML中的形状信息
	decoder := xml.NewDecoder(bytes.NewReader(data))
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
			if t.Name.Local == "w:drawing" {
				shapeElement, err := dgp.parseDrawingElement(decoder, &t)
				if err == nil && shapeElement != nil {
					shapes = append(shapes, shapeElement)
				}
			}
		}
	}

	return shapes, nil
}

// parseDrawingElement 解析绘图元素
func (dgp *DOCXGraphicsParser) parseDrawingElement(decoder *xml.Decoder, startElement *xml.StartElement) (*types.GraphicElement, error) {
	element := &types.GraphicElement{
		Type: types.GraphicTypeShape,
		ID:   fmt.Sprintf("shape_%d", len(startElement.Attr)),
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
	}

	// 解析属性
	for _, attr := range startElement.Attr {
		switch attr.Name.Local {
		case "id":
			element.ID = attr.Value
		}
	}

	// 解析子元素
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
			case "wp:extent":
				dgp.parseExtent(decoder, &t, element)
			case "wp:docPr":
				dgp.parseDocPr(decoder, &t, element)
			case "a:graphic":
				dgp.parseGraphic(decoder, &t, element)
			}
		case xml.EndElement:
			if t.Name.Local == "w:drawing" {
				return element, nil
			}
		}
	}

	return element, nil
}

// parseExtent 解析尺寸信息
func (dgp *DOCXGraphicsParser) parseExtent(decoder *xml.Decoder, startElement *xml.StartElement, element *types.GraphicElement) {
	for _, attr := range startElement.Attr {
		switch attr.Name.Local {
		case "cx":
			if width, err := strconv.ParseFloat(attr.Value, 64); err == nil {
				element.Size.Width = width
			}
		case "cy":
			if height, err := strconv.ParseFloat(attr.Value, 64); err == nil {
				element.Size.Height = height
			}
		}
	}
	element.Size.Unit = "emu"
}

// parseDocPr 解析文档属性
func (dgp *DOCXGraphicsParser) parseDocPr(decoder *xml.Decoder, startElement *xml.StartElement, element *types.GraphicElement) {
	for _, attr := range startElement.Attr {
		switch attr.Name.Local {
		case "id":
			element.ID = fmt.Sprintf("shape_%s", attr.Value)
		case "name":
			// 可以设置名称
		}
	}
}

// parseGraphic 解析图形信息
func (dgp *DOCXGraphicsParser) parseGraphic(decoder *xml.Decoder, startElement *xml.StartElement, element *types.GraphicElement) {
	// 解析图形类型和样式
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
			case "a:graphicData":
				dgp.parseGraphicData(decoder, &t, element)
			}
		case xml.EndElement:
			if t.Name.Local == "a:graphic" {
				return
			}
		}
	}
}

// parseGraphicData 解析图形数据
func (dgp *DOCXGraphicsParser) parseGraphicData(decoder *xml.Decoder, startElement *xml.StartElement, element *types.GraphicElement) {
	for _, attr := range startElement.Attr {
		switch attr.Name.Local {
		case "uri":
			// 根据URI确定图形类型
			switch attr.Value {
			case "http://schemas.openxmlformats.org/drawingml/2006/picture":
				element.Type = types.GraphicTypeImage
			case "http://schemas.openxmlformats.org/drawingml/2006/chart":
				element.Type = types.GraphicTypeChart
			case "http://schemas.openxmlformats.org/drawingml/2006/shape":
				element.Type = types.GraphicTypeShape
			}
		}
	}
}

// parseCharts 解析图表元素
func (dgp *DOCXGraphicsParser) parseCharts(reader *zip.ReadCloser) ([]*types.GraphicElement, error) {
	var charts []*types.GraphicElement

	// 查找图表文件
	for _, file := range reader.File {
		if strings.HasPrefix(file.Name, "word/charts/") && strings.HasSuffix(file.Name, ".xml") {
			chartElement, err := dgp.parseChartFile(file)
			if err != nil {
				continue
			}
			charts = append(charts, chartElement)
		}
	}

	return charts, nil
}

// parseChartFile 解析图表文件
func (dgp *DOCXGraphicsParser) parseChartFile(file *zip.File) (*types.GraphicElement, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, fmt.Errorf("failed to open chart file: %w", err)
	}
	defer rc.Close()

	data, err := io.ReadAll(rc)
	if err != nil {
		return nil, fmt.Errorf("failed to read chart data: %w", err)
	}

	element := &types.GraphicElement{
		ID:   fmt.Sprintf("chart_%s", filepath.Base(file.Name)),
		Type: types.GraphicTypeChart,
		Content: types.GraphicContent{
			Chart: types.ChartData{
				Type: "unknown",
			},
		},
		Visible: true,
		Locked:  false,
	}

	// 解析图表数据
	dgp.parseChartData(data, element)

	return element, nil
}

// parseChartData 解析图表数据
func (dgp *DOCXGraphicsParser) parseChartData(data []byte, element *types.GraphicElement) {
	// 基础实现，解析图表类型
	decoder := xml.NewDecoder(bytes.NewReader(data))
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
			case "c:barChart":
				element.Content.Chart.Type = "bar"
			case "c:lineChart":
				element.Content.Chart.Type = "line"
			case "c:pieChart":
				element.Content.Chart.Type = "pie"
			case "c:scatterChart":
				element.Content.Chart.Type = "scatter"
			}
		}
	}
}

// parseSmartArts 解析SmartArt元素
func (dgp *DOCXGraphicsParser) parseSmartArts(reader *zip.ReadCloser) ([]*types.GraphicElement, error) {
	var smartArts []*types.GraphicElement

	// 查找SmartArt文件
	for _, file := range reader.File {
		if strings.HasPrefix(file.Name, "word/diagrams/") && strings.HasSuffix(file.Name, ".xml") {
			smartArtElement, err := dgp.parseSmartArtFile(file)
			if err != nil {
				continue
			}
			smartArts = append(smartArts, smartArtElement)
		}
	}

	return smartArts, nil
}

// parseSmartArtFile 解析SmartArt文件
func (dgp *DOCXGraphicsParser) parseSmartArtFile(file *zip.File) (*types.GraphicElement, error) {
	element := &types.GraphicElement{
		ID:   fmt.Sprintf("smartart_%s", filepath.Base(file.Name)),
		Type: types.GraphicTypeSmartArt,
		Content: types.GraphicContent{
			SmartArt: types.SmartArtData{
				Type: "unknown",
			},
		},
		Visible: true,
		Locked:  false,
	}

	return element, nil
}

// parseTextboxes 解析文本框元素
func (dgp *DOCXGraphicsParser) parseTextboxes(reader *zip.ReadCloser) ([]*types.GraphicElement, error) {
	var textboxes []*types.GraphicElement

	// 文本框通常在drawing.xml中定义
	// 这里简化处理，实际需要更复杂的解析逻辑
	element := &types.GraphicElement{
		ID:   "textbox_1",
		Type: types.GraphicTypeTextbox,
		Content: types.GraphicContent{
			Text: "",
		},
		Visible: true,
		Locked:  false,
	}

	textboxes = append(textboxes, element)
	return textboxes, nil
}

// parseFormulas 解析公式元素
func (dgp *DOCXGraphicsParser) parseFormulas(reader *zip.ReadCloser) ([]*types.GraphicElement, error) {
	var formulas []*types.GraphicElement

	// 查找公式文件
	for _, file := range reader.File {
		if strings.Contains(file.Name, "equation") && strings.HasSuffix(file.Name, ".xml") {
			formulaElement, err := dgp.parseFormulaFile(file)
			if err != nil {
				continue
			}
			formulas = append(formulas, formulaElement)
		}
	}

	return formulas, nil
}

// parseFormulaFile 解析公式文件
func (dgp *DOCXGraphicsParser) parseFormulaFile(file *zip.File) (*types.GraphicElement, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, fmt.Errorf("failed to open formula file: %w", err)
	}
	defer rc.Close()

	data, err := io.ReadAll(rc)
	if err != nil {
		return nil, fmt.Errorf("failed to read formula data: %w", err)
	}

	element := &types.GraphicElement{
		ID:   fmt.Sprintf("formula_%s", filepath.Base(file.Name)),
		Type: types.GraphicTypeFormula,
		Content: types.GraphicContent{
			Formula: types.FormulaData{
				Content: string(data),
				Format:  "MathML",
			},
		},
		Visible: true,
		Locked:  false,
	}

	return element, nil
}

// parseGroups 解析图形组
func (dgp *DOCXGraphicsParser) parseGroups(reader *zip.ReadCloser) ([]types.GraphicGroup, error) {
	var groups []types.GraphicGroup

	// 基础实现，实际需要从文档中解析组信息
	group := types.GraphicGroup{
		ID:       "group_1",
		Name:     "Default Group",
		Elements: []string{},
		Visible:  true,
		Locked:   false,
	}

	groups = append(groups, group)
	return groups, nil
}
