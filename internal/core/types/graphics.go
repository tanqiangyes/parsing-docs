package types

import (
	"time"
)

// GraphicElement 基础图形元素
type GraphicElement struct {
	ID       string          `json:"id" xml:"id,attr"`
	Type     GraphicType     `json:"type" xml:"type,attr"`
	Position GraphicPosition `json:"position" xml:"position"`
	Size     Size            `json:"size" xml:"size"`
	Style    GraphicStyle    `json:"style" xml:"style"`
	Content  GraphicContent  `json:"content" xml:"content"`
	Metadata GraphicMetadata `json:"metadata" xml:"metadata"`
	Anchor   Anchor          `json:"anchor" xml:"anchor"`
	ZIndex   int             `json:"z_index" xml:"z-index,attr"`
	Visible  bool            `json:"visible" xml:"visible,attr"`
	Locked   bool            `json:"locked" xml:"locked,attr"`
}

// GraphicType 图形类型
type GraphicType string

const (
	GraphicTypeImage    GraphicType = "image"
	GraphicTypeShape    GraphicType = "shape"
	GraphicTypeChart    GraphicType = "chart"
	GraphicTypeSmartArt GraphicType = "smartart"
	GraphicTypeTextbox  GraphicType = "textbox"
	GraphicTypeWordArt  GraphicType = "wordart"
	GraphicTypeFormula  GraphicType = "formula"
	GraphicTypeGroup    GraphicType = "group"
)

// GraphicPosition 图形位置信息
type GraphicPosition struct {
	X         float64 `json:"x" xml:"x,attr"`
	Y         float64 `json:"y" xml:"y,attr"`
	RelativeX float64 `json:"relative_x" xml:"relative-x,attr"`
	RelativeY float64 `json:"relative_y" xml:"relative-y,attr"`
	Unit      string  `json:"unit" xml:"unit,attr"` // emu, pt, px
}

// Size 尺寸信息
type Size struct {
	Width           float64 `json:"width" xml:"width,attr"`
	Height          float64 `json:"height" xml:"height,attr"`
	ScaleX          float64 `json:"scale_x" xml:"scale-x,attr"`
	ScaleY          float64 `json:"scale_y" xml:"scale-y,attr"`
	Unit            string  `json:"unit" xml:"unit,attr"`
	LockAspectRatio bool    `json:"lock_aspect_ratio" xml:"lock-aspect-ratio,attr"`
}

// GraphicStyle 图形样式
type GraphicStyle struct {
	Border   GraphicBorder `json:"border" xml:"border"`
	Fill     Fill          `json:"fill" xml:"fill"`
	Shadow   Shadow        `json:"shadow" xml:"shadow"`
	Effects  Effects       `json:"effects" xml:"effects"`
	Opacity  float64       `json:"opacity" xml:"opacity,attr"`
	Rotation float64       `json:"rotation" xml:"rotation,attr"`
	FlipH    bool          `json:"flip_h" xml:"flip-h,attr"`
	FlipV    bool          `json:"flip_v" xml:"flip-v,attr"`
}

// GraphicBorder 图形边框样式
type GraphicBorder struct {
	Width   float64      `json:"width" xml:"width,attr"`
	Style   string       `json:"style" xml:"style,attr"` // solid, dashed, dotted
	Color   GraphicColor `json:"color" xml:"color"`
	Visible bool         `json:"visible" xml:"visible,attr"`
}

// Fill 填充样式
type Fill struct {
	Type     string         `json:"type" xml:"type,attr"` // solid, gradient, pattern, picture
	Color    GraphicColor   `json:"color" xml:"color"`
	Gradient Gradient       `json:"gradient" xml:"gradient"`
	Pattern  GraphicPattern `json:"pattern" xml:"pattern"`
	Picture  Picture        `json:"picture" xml:"picture"`
}

// Shadow 阴影效果
type Shadow struct {
	Enabled      bool         `json:"enabled" xml:"enabled,attr"`
	Color        GraphicColor `json:"color" xml:"color"`
	Blur         float64      `json:"blur" xml:"blur,attr"`
	Distance     float64      `json:"distance" xml:"distance,attr"`
	Angle        float64      `json:"angle" xml:"angle,attr"`
	Size         float64      `json:"size" xml:"size,attr"`
	Transparency float64      `json:"transparency" xml:"transparency,attr"`
}

// Effects 特效
type Effects struct {
	Glow       Glow       `json:"glow" xml:"glow"`
	Reflection Reflection `json:"reflection" xml:"reflection"`
	SoftEdge   SoftEdge   `json:"soft_edge" xml:"soft-edge"`
	Preset     string     `json:"preset" xml:"preset,attr"`
}

// Glow 发光效果
type Glow struct {
	Enabled      bool         `json:"enabled" xml:"enabled,attr"`
	Color        GraphicColor `json:"color" xml:"color"`
	Size         float64      `json:"size" xml:"size,attr"`
	Transparency float64      `json:"transparency" xml:"transparency,attr"`
}

// Reflection 反射效果
type Reflection struct {
	Enabled      bool    `json:"enabled" xml:"enabled,attr"`
	Transparency float64 `json:"transparency" xml:"transparency,attr"`
	Size         float64 `json:"size" xml:"size,attr"`
	Distance     float64 `json:"distance" xml:"distance,attr"`
	Blur         float64 `json:"blur" xml:"blur,attr"`
}

// SoftEdge 柔化边缘
type SoftEdge struct {
	Enabled bool    `json:"enabled" xml:"enabled,attr"`
	Radius  float64 `json:"radius" xml:"radius,attr"`
}

// Gradient 渐变
type Gradient struct {
	Type  string         `json:"type" xml:"type,attr"` // linear, radial, rectangular
	Stops []GradientStop `json:"stops" xml:"stops>stop"`
	Angle float64        `json:"angle" xml:"angle,attr"`
}

// GradientStop 渐变停止点
type GradientStop struct {
	Position float64      `json:"position" xml:"position,attr"`
	Color    GraphicColor `json:"color" xml:"color"`
}

// GraphicPattern 图形图案
type GraphicPattern struct {
	Type      string       `json:"type" xml:"type,attr"`
	ForeColor GraphicColor `json:"fore_color" xml:"fore-color"`
	BackColor GraphicColor `json:"back_color" xml:"back-color"`
}

// Picture 图片填充
type Picture struct {
	Source string `json:"source" xml:"source,attr"`
	Format string `json:"format" xml:"format,attr"`
	Data   []byte `json:"data" xml:"data"`
}

// GraphicContent 图形内容
type GraphicContent struct {
	Text     string       `json:"text" xml:"text"`
	Image    ImageData    `json:"image" xml:"image"`
	Chart    ChartData    `json:"chart" xml:"chart"`
	SmartArt SmartArtData `json:"smartart" xml:"smartart"`
	Formula  FormulaData  `json:"formula" xml:"formula"`
}

// ImageData 图片数据
type ImageData struct {
	Source     string `json:"source" xml:"source,attr"`
	Format     string `json:"format" xml:"format,attr"`
	Width      int    `json:"width" xml:"width,attr"`
	Height     int    `json:"height" xml:"height,attr"`
	DPI        int    `json:"dpi" xml:"dpi,attr"`
	Data       []byte `json:"data" xml:"data"`
	Compressed bool   `json:"compressed" xml:"compressed,attr"`
	AltText    string `json:"alt_text" xml:"alt-text"`
}

// ChartData 图表数据
type ChartData struct {
	Type   string        `json:"type" xml:"type,attr"`
	Title  string        `json:"title" xml:"title"`
	Data   []ChartSeries `json:"data" xml:"data>series"`
	Axes   ChartAxes     `json:"axes" xml:"axes"`
	Legend ChartLegend   `json:"legend" xml:"legend"`
}

// ChartSeries 图表系列
type ChartSeries struct {
	Name  string           `json:"name" xml:"name,attr"`
	Type  string           `json:"type" xml:"type,attr"`
	Data  []ChartPoint     `json:"data" xml:"data>point"`
	Style ChartSeriesStyle `json:"style" xml:"style"`
}

// ChartPoint 图表数据点
type ChartPoint struct {
	X     interface{} `json:"x" xml:"x"`
	Y     interface{} `json:"y" xml:"y"`
	Label string      `json:"label" xml:"label"`
}

// ChartSeriesStyle 图表系列样式
type ChartSeriesStyle struct {
	Color   GraphicColor `json:"color" xml:"color"`
	Pattern string       `json:"pattern" xml:"pattern,attr"`
	Width   float64      `json:"width" xml:"width,attr"`
}

// ChartAxes 图表坐标轴
type ChartAxes struct {
	XAxis ChartAxis `json:"x_axis" xml:"x-axis"`
	YAxis ChartAxis `json:"y_axis" xml:"y-axis"`
}

// ChartAxis 图表坐标轴
type ChartAxis struct {
	Title  string  `json:"title" xml:"title"`
	Min    float64 `json:"min" xml:"min,attr"`
	Max    float64 `json:"max" xml:"max,attr"`
	Step   float64 `json:"step" xml:"step,attr"`
	Format string  `json:"format" xml:"format,attr"`
}

// ChartLegend 图表图例
type ChartLegend struct {
	Visible  bool   `json:"visible" xml:"visible,attr"`
	Position string `json:"position" xml:"position,attr"`
	Title    string `json:"title" xml:"title"`
}

// SmartArtData SmartArt数据
type SmartArtData struct {
	Type   string         `json:"type" xml:"type,attr"`
	Layout string         `json:"layout" xml:"layout,attr"`
	Nodes  []SmartArtNode `json:"nodes" xml:"nodes>node"`
	Style  SmartArtStyle  `json:"style" xml:"style"`
}

// SmartArtNode SmartArt节点
type SmartArtNode struct {
	ID       string   `json:"id" xml:"id,attr"`
	Text     string   `json:"text" xml:"text"`
	Level    int      `json:"level" xml:"level,attr"`
	ParentID string   `json:"parent_id" xml:"parent-id,attr"`
	Children []string `json:"children" xml:"children>child"`
}

// SmartArtStyle SmartArt样式
type SmartArtStyle struct {
	Theme  string `json:"theme" xml:"theme,attr"`
	Color  string `json:"color" xml:"color,attr"`
	Layout string `json:"layout" xml:"layout,attr"`
}

// FormulaData 公式数据
type FormulaData struct {
	Content string  `json:"content" xml:"content"`
	Format  string  `json:"format" xml:"format,attr"` // MathML, LaTeX
	Size    float64 `json:"size" xml:"size,attr"`
}

// GraphicMetadata 图形元数据
type GraphicMetadata struct {
	FileName string    `json:"file_name" xml:"file-name"`
	Created  time.Time `json:"created" xml:"created"`
	Modified time.Time `json:"modified" xml:"modified"`
	Author   string    `json:"author" xml:"author"`
	Comments string    `json:"comments" xml:"comments"`
	Tags     []string  `json:"tags" xml:"tags>tag"`
}

// Anchor 锚点信息
type Anchor struct {
	Type     string  `json:"type" xml:"type,attr"` // character, paragraph, page
	ID       string  `json:"id" xml:"id,attr"`
	Position string  `json:"position" xml:"position,attr"` // top-left, center, bottom-right
	OffsetX  float64 `json:"offset_x" xml:"offset-x,attr"`
	OffsetY  float64 `json:"offset_y" xml:"offset-y,attr"`
}

// GraphicColor 图形颜色
type GraphicColor struct {
	Type   string  `json:"type" xml:"type,attr"` // rgb, hsl, theme, scheme
	Value  string  `json:"value" xml:"value,attr"`
	Alpha  float64 `json:"alpha" xml:"alpha,attr"`
	Theme  string  `json:"theme" xml:"theme,attr"`
	Scheme string  `json:"scheme" xml:"scheme,attr"`
}

// DocumentGraphics 文档图形集合
type DocumentGraphics struct {
	Elements []*GraphicElement `json:"elements" xml:"elements>element"`
	Count    int               `json:"count" xml:"count,attr"`
	Groups   []GraphicGroup    `json:"groups" xml:"groups>group"`
}

// GraphicGroup 图形组
type GraphicGroup struct {
	ID       string   `json:"id" xml:"id,attr"`
	Name     string   `json:"name" xml:"name,attr"`
	Elements []string `json:"elements" xml:"elements>element"`
	Visible  bool     `json:"visible" xml:"visible,attr"`
	Locked   bool     `json:"locked" xml:"locked,attr"`
}
