package examples

import (
	"docs-parser/internal/core/graphics"
	"docs-parser/internal/core/types"
	"fmt"
	"log"
)

// GraphicsParsingExample 演示图形解析功能
func GraphicsParsingExample() {
	fmt.Println("=== 图形解析功能示例 ===")

	// 模拟解析DOCX文档中的图形
	fmt.Println("1. 解析DOCX文档中的图形元素...")

	// 注意：这里需要提供一个实际的DOCX文件路径
	// docxParser := graphics.NewDOCXGraphicsParser()
	// docxPath := "path/to/your/document.docx"
	// graphics, err := docxParser.ParseGraphics(docxPath)

	// 由于没有实际的DOCX文件，我们创建一个模拟的图形数据
	fmt.Println("2. 创建模拟图形数据...")

	// 创建图片元素
	imageElement := &types.GraphicElement{
		ID:   "image_1",
		Type: types.GraphicTypeImage,
		Position: types.GraphicPosition{
			X:         100,
			Y:         200,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           300,
			Height:          200,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: true,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
		},
		Content: types.GraphicContent{
			Image: types.ImageData{
				Source:     "word/media/image1.png",
				Format:     "png",
				Width:      300,
				Height:     200,
				DPI:        96,
				Compressed: false,
				AltText:    "示例图片",
			},
		},
		Metadata: types.GraphicMetadata{
			FileName: "image1.png",
			Author:   "用户",
		},
		Anchor: types.Anchor{
			Type:     "paragraph",
			ID:       "para_1",
			Position: "top-left",
		},
		ZIndex:  1,
		Visible: true,
		Locked:  false,
	}

	// 创建形状元素
	shapeElement := &types.GraphicElement{
		ID:   "shape_1",
		Type: types.GraphicTypeShape,
		Position: types.GraphicPosition{
			X:         400,
			Y:         300,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           150,
			Height:          100,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: false,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
			Border: types.GraphicBorder{
				Width:   2.0,
				Style:   "solid",
				Visible: true,
				Color: types.GraphicColor{
					Type:  "rgb",
					Value: "#000000",
					Alpha: 1.0,
				},
			},
			Fill: types.Fill{
				Type: "solid",
				Color: types.GraphicColor{
					Type:  "rgb",
					Value: "#FF0000",
					Alpha: 0.8,
				},
			},
		},
		Content: types.GraphicContent{
			Text: "矩形形状",
		},
		Metadata: types.GraphicMetadata{
			FileName: "shape1",
			Author:   "用户",
		},
		Anchor: types.Anchor{
			Type:     "paragraph",
			ID:       "para_2",
			Position: "center",
		},
		ZIndex:  2,
		Visible: true,
		Locked:  false,
	}

	// 创建图表元素
	chartElement := &types.GraphicElement{
		ID:   "chart_1",
		Type: types.GraphicTypeChart,
		Position: types.GraphicPosition{
			X:         600,
			Y:         400,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           400,
			Height:          300,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: true,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
		},
		Content: types.GraphicContent{
			Chart: types.ChartData{
				Type:  "bar",
				Title: "销售数据",
				Data: []types.ChartSeries{
					{
						Name: "第一季度",
						Type: "bar",
						Data: []types.ChartPoint{
							{X: "1月", Y: 100, Label: "100"},
							{X: "2月", Y: 150, Label: "150"},
							{X: "3月", Y: 200, Label: "200"},
						},
						Style: types.ChartSeriesStyle{
							Color: types.GraphicColor{
								Type:  "rgb",
								Value: "#4CAF50",
								Alpha: 1.0,
							},
							Pattern: "solid",
							Width:   2.0,
						},
					},
				},
				Axes: types.ChartAxes{
					XAxis: types.ChartAxis{
						Title:  "月份",
						Format: "text",
					},
					YAxis: types.ChartAxis{
						Title:  "销售额",
						Min:    0,
						Max:    250,
						Step:   50,
						Format: "number",
					},
				},
				Legend: types.ChartLegend{
					Visible:  true,
					Position: "bottom",
					Title:    "图例",
				},
			},
		},
		Metadata: types.GraphicMetadata{
			FileName: "chart1",
			Author:   "用户",
		},
		Anchor: types.Anchor{
			Type:     "paragraph",
			ID:       "para_3",
			Position: "center",
		},
		ZIndex:  3,
		Visible: true,
		Locked:  false,
	}

	// 创建SmartArt元素
	smartArtElement := &types.GraphicElement{
		ID:   "smartart_1",
		Type: types.GraphicTypeSmartArt,
		Position: types.GraphicPosition{
			X:         800,
			Y:         500,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           350,
			Height:          250,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: true,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
		},
		Content: types.GraphicContent{
			SmartArt: types.SmartArtData{
				Type:   "hierarchy",
				Layout: "vertical",
				Nodes: []types.SmartArtNode{
					{
						ID:       "node_1",
						Text:     "总经理",
						Level:    0,
						ParentID: "",
						Children: []string{"node_2", "node_3"},
					},
					{
						ID:       "node_2",
						Text:     "技术部",
						Level:    1,
						ParentID: "node_1",
						Children: []string{},
					},
					{
						ID:       "node_3",
						Text:     "市场部",
						Level:    1,
						ParentID: "node_1",
						Children: []string{},
					},
				},
				Style: types.SmartArtStyle{
					Theme:  "default",
					Color:  "blue",
					Layout: "vertical",
				},
			},
		},
		Metadata: types.GraphicMetadata{
			FileName: "smartart1",
			Author:   "用户",
		},
		Anchor: types.Anchor{
			Type:     "paragraph",
			ID:       "para_4",
			Position: "center",
		},
		ZIndex:  4,
		Visible: true,
		Locked:  false,
	}

	// 创建文本框元素
	textboxElement := &types.GraphicElement{
		ID:   "textbox_1",
		Type: types.GraphicTypeTextbox,
		Position: types.GraphicPosition{
			X:         1000,
			Y:         600,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           200,
			Height:          100,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: false,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
			Border: types.GraphicBorder{
				Width:   1.0,
				Style:   "solid",
				Visible: true,
				Color: types.GraphicColor{
					Type:  "rgb",
					Value: "#666666",
					Alpha: 1.0,
				},
			},
			Fill: types.Fill{
				Type: "solid",
				Color: types.GraphicColor{
					Type:  "rgb",
					Value: "#FFFFFF",
					Alpha: 1.0,
				},
			},
		},
		Content: types.GraphicContent{
			Text: "这是一个文本框，可以包含格式化的文本内容。",
		},
		Metadata: types.GraphicMetadata{
			FileName: "textbox1",
			Author:   "用户",
		},
		Anchor: types.Anchor{
			Type:     "paragraph",
			ID:       "para_5",
			Position: "top-left",
		},
		ZIndex:  5,
		Visible: true,
		Locked:  false,
	}

	// 创建公式元素
	formulaElement := &types.GraphicElement{
		ID:   "formula_1",
		Type: types.GraphicTypeFormula,
		Position: types.GraphicPosition{
			X:         1200,
			Y:         700,
			RelativeX: 0,
			RelativeY: 0,
			Unit:      "emu",
		},
		Size: types.Size{
			Width:           150,
			Height:          50,
			ScaleX:          1.0,
			ScaleY:          1.0,
			Unit:            "emu",
			LockAspectRatio: true,
		},
		Style: types.GraphicStyle{
			Opacity: 1.0,
		},
		Content: types.GraphicContent{
			Formula: types.FormulaData{
				Content: "x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}",
				Format:  "LaTeX",
				Size:    12.0,
			},
		},
		Metadata: types.GraphicMetadata{
			FileName: "formula1",
			Author:   "用户",
		},
		Anchor: types.Anchor{
			Type:     "paragraph",
			ID:       "para_6",
			Position: "center",
		},
		ZIndex:  6,
		Visible: true,
		Locked:  false,
	}

	// 创建图形组
	graphicGroup := types.GraphicGroup{
		ID:       "group_1",
		Name:     "示例图形组",
		Elements: []string{"image_1", "shape_1", "chart_1"},
		Visible:  true,
		Locked:   false,
	}

	// 创建文档图形集合
	documentGraphics := &types.DocumentGraphics{
		Elements: []*types.GraphicElement{
			imageElement,
			shapeElement,
			chartElement,
			smartArtElement,
			textboxElement,
			formulaElement,
		},
		Count:  6,
		Groups: []types.GraphicGroup{graphicGroup},
	}

	// 验证图形元素
	fmt.Println("3. 验证图形元素...")
	defaultParser := graphics.NewDefaultParser()

	for _, element := range documentGraphics.Elements {
		if err := defaultParser.ValidateGraphicElement(element); err != nil {
			log.Printf("验证失败 - %s: %v", element.ID, err)
		} else {
			fmt.Printf("✓ %s 验证通过\n", element.ID)
		}
	}

	// 显示图形信息
	fmt.Println("\n4. 图形元素信息:")
	fmt.Printf("总图形数量: %d\n", documentGraphics.Count)
	fmt.Printf("图形组数量: %d\n", len(documentGraphics.Groups))

	fmt.Println("\n图形元素详情:")
	for _, element := range documentGraphics.Elements {
		fmt.Printf("- %s (%s): 位置(%f, %f), 尺寸(%f x %f)\n",
			element.ID, element.Type, element.Position.X, element.Position.Y,
			element.Size.Width, element.Size.Height)
	}

	fmt.Println("\n5. 图形解析功能演示完成!")
}
