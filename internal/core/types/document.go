package types

import (
	"time"
)

// Document 表示解析后的Word文档
type Document struct {
	Metadata    DocumentMetadata `json:"metadata"`
	Content     DocumentContent  `json:"content"`
	Styles      DocumentStyles   `json:"styles"`
	FormatRules FormatRules      `json:"format_rules"`
}

// DocumentMetadata 文档元数据
type DocumentMetadata struct {
	Title       string    `json:"title"`
	Author      string    `json:"author"`
	Subject     string    `json:"subject"`
	Keywords    []string  `json:"keywords"`
	Created     time.Time `json:"created"`
	Modified    time.Time `json:"modified"`
	LastSavedBy string    `json:"last_saved_by"`
	Revision    int       `json:"revision"`
	Version     string    `json:"version"`
	FileSize    int64     `json:"file_size"`
	WordCount   int       `json:"word_count"`
	PageCount   int       `json:"page_count"`
}

// DocumentContent 文档内容
type DocumentContent struct {
	Paragraphs []Paragraph `json:"paragraphs"`
	Sections   []Section   `json:"sections"`
	Headers    []Header    `json:"headers"`
	Footers    []Footer    `json:"footers"`
	Tables     []Table     `json:"tables"`
	Images     []Image     `json:"images"`
	Comments   []Comment   `json:"comments"`
	Bookmarks  []Bookmark  `json:"bookmarks"`
}

// Paragraph 段落
type Paragraph struct {
	ID          string           `json:"id"`
	Text        string           `json:"text"`
	Style       ParagraphStyle   `json:"style"`
	Alignment   Alignment        `json:"alignment"`
	Indentation Indentation      `json:"indentation"`
	Spacing     Spacing          `json:"spacing"`
	Borders     Borders          `json:"borders"`
	Shading     Shading          `json:"shading"`
	Runs        []TextRun        `json:"runs"`
	PageBreak   bool             `json:"page_break"`
	KeepLines   bool             `json:"keep_lines"`
	KeepNext    bool             `json:"keep_next"`
	OutlineLevel int             `json:"outline_level"`
}

// TextRun 文本运行
type TextRun struct {
	ID       string     `json:"id"`
	Text     string     `json:"text"`
	Font     Font       `json:"font"`
	Bold     bool       `json:"bold"`
	Italic   bool       `json:"italic"`
	Underline Underline `json:"underline"`
	Color    Color      `json:"color"`
	Highlight Highlight  `json:"highlight"`
	Size     float64    `json:"size"`
	Position Position   `json:"position"`
}

// Section 节
type Section struct {
	ID              string        `json:"id"`
	PageSize        PageSize      `json:"page_size"`
	PageMargins     PageMargins   `json:"page_margins"`
	HeaderDistance  float64       `json:"header_distance"`
	FooterDistance  float64       `json:"footer_distance"`
	Columns         Columns       `json:"columns"`
	PageNumbering   PageNumbering `json:"page_numbering"`
	LineNumbering   LineNumbering `json:"line_numbering"`
}

// Table 表格
type Table struct {
	ID       string        `json:"id"`
	Rows     []TableRow    `json:"rows"`
	Style    TableStyle    `json:"style"`
	Borders  TableBorders  `json:"borders"`
	Shading  TableShading  `json:"shading"`
	Width    float64       `json:"width"`
	Alignment Alignment     `json:"alignment"`
}

// TableRow 表格行
type TableRow struct {
	ID       string       `json:"id"`
	Cells    []TableCell  `json:"cells"`
	Height   float64      `json:"height"`
	Header   bool         `json:"header"`
	Repeat   bool         `json:"repeat"`
}

// TableCell 表格单元格
type TableCell struct {
	ID       string       `json:"id"`
	Content  []Paragraph  `json:"content"`
	Width    float64      `json:"width"`
	Height   float64      `json:"height"`
	Borders  CellBorders  `json:"borders"`
	Shading  CellShading  `json:"shading"`
	VerticalAlignment VerticalAlignment `json:"vertical_alignment"`
	Merge    CellMerge    `json:"merge"`
}

// FormatRules 格式规则
type FormatRules struct {
	FontRules      []FontRule      `json:"font_rules"`
	ParagraphRules []ParagraphRule `json:"paragraph_rules"`
	TableRules     []TableRule     `json:"table_rules"`
	PageRules      []PageRule      `json:"page_rules"`
	StyleRules     []StyleRule     `json:"style_rules"`
}

// FontRule 字体规则
type FontRule struct {
	ID          string   `json:"id"`
	Name        string   `json:"name"`
	Size        float64  `json:"size"`
	Color       Color    `json:"color"`
	Bold        bool     `json:"bold"`
	Italic      bool     `json:"italic"`
	Underline   Underline `json:"underline"`
	Highlight   Highlight `json:"highlight"`
	Position    Position `json:"position"`
	Spacing     float64  `json:"spacing"`
	Scale       float64  `json:"scale"`
	Kerning     float64  `json:"kerning"`
}

// ParagraphRule 段落规则
type ParagraphRule struct {
	ID          string     `json:"id"`
	Name        string     `json:"name"`
	Alignment   Alignment  `json:"alignment"`
	Indentation Indentation `json:"indentation"`
	Spacing     Spacing    `json:"spacing"`
	Borders     Borders    `json:"borders"`
	Shading     Shading    `json:"shading"`
	OutlineLevel int       `json:"outline_level"`
	KeepLines   bool       `json:"keep_lines"`
	KeepNext    bool       `json:"keep_next"`
	PageBreak   bool       `json:"page_break"`
}

// 基础类型定义
type Font struct {
	Name     string  `json:"name"`
	Size     float64 `json:"size"`
	Color    Color   `json:"color"`
	Bold     bool    `json:"bold"`
	Italic   bool    `json:"italic"`
	Underline Underline `json:"underline"`
	Highlight Highlight `json:"highlight"`
}

type Color struct {
	RGB string `json:"rgb"`
	Theme int  `json:"theme"`
}

type Alignment string
const (
	AlignLeft   Alignment = "left"
	AlignCenter Alignment = "center"
	AlignRight  Alignment = "right"
	AlignJustify Alignment = "justify"
)

type Indentation struct {
	Left   float64 `json:"left"`
	Right  float64 `json:"right"`
	First  float64 `json:"first"`
	Hanging float64 `json:"hanging"`
}

type Spacing struct {
	Before float64 `json:"before"`
	After  float64 `json:"after"`
	Line   float64 `json:"line"`
}

type Borders struct {
	Top    Border `json:"top"`
	Bottom Border `json:"bottom"`
	Left   Border `json:"left"`
	Right  Border `json:"right"`
}

type Border struct {
	Style  BorderStyle `json:"style"`
	Width  float64     `json:"width"`
	Color  Color       `json:"color"`
	Space  float64     `json:"space"`
}

type BorderStyle string
const (
	BorderNone BorderStyle = "none"
	BorderSingle BorderStyle = "single"
	BorderDouble BorderStyle = "double"
	BorderDotted BorderStyle = "dotted"
	BorderDashed BorderStyle = "dashed"
)

type Shading struct {
	Fill   Color `json:"fill"`
	Pattern Pattern `json:"pattern"`
}

type Pattern string
const (
	PatternNone Pattern = "none"
	PatternSolid Pattern = "solid"
	PatternClear Pattern = "clear"
)

type Underline string
const (
	UnderlineNone Underline = "none"
	UnderlineSingle Underline = "single"
	UnderlineDouble Underline = "double"
	UnderlineDotted Underline = "dotted"
	UnderlineDashed Underline = "dashed"
)

type Highlight string
const (
	HighlightNone Highlight = "none"
	HighlightYellow Highlight = "yellow"
	HighlightGreen Highlight = "green"
	HighlightPink Highlight = "pink"
	HighlightBlue Highlight = "blue"
	HighlightRed Highlight = "red"
)

type Position string
const (
	PositionNormal Position = "normal"
	PositionSuperscript Position = "superscript"
	PositionSubscript Position = "subscript"
)

type VerticalAlignment string
const (
	VAlignTop VerticalAlignment = "top"
	VAlignCenter VerticalAlignment = "center"
	VAlignBottom VerticalAlignment = "bottom"
)

type CellMerge struct {
	Horizontal int `json:"horizontal"`
	Vertical   int `json:"vertical"`
}

// 其他类型定义
type Header struct {
	ID      string      `json:"id"`
	Content []Paragraph `json:"content"`
}

type Footer struct {
	ID      string      `json:"id"`
	Content []Paragraph `json:"content"`
}

type Image struct {
	ID       string  `json:"id"`
	Path     string  `json:"path"`
	Width    float64 `json:"width"`
	Height   float64 `json:"height"`
	AltText  string  `json:"alt_text"`
}

type Comment struct {
	ID      string `json:"id"`
	Author  string `json:"author"`
	Date    time.Time `json:"date"`
	Text    string `json:"text"`
}

type Bookmark struct {
	ID   string `json:"id"`
	Name string `json:"name"`
}

type PageSize struct {
	Width  float64 `json:"width"`
	Height float64 `json:"height"`
}

type PageMargins struct {
	Top    float64 `json:"top"`
	Bottom float64 `json:"bottom"`
	Left   float64 `json:"left"`
	Right  float64 `json:"right"`
	Header float64 `json:"header"`
	Footer float64 `json:"footer"`
}

type Columns struct {
	Count    int     `json:"count"`
	Spacing  float64 `json:"spacing"`
	Equal    bool    `json:"equal"`
}

type PageNumbering struct {
	Start     int    `json:"start"`
	Format    string `json:"format"`
	Restart   bool   `json:"restart"`
}

type LineNumbering struct {
	Start     int    `json:"start"`
	Increment int    `json:"increment"`
	Restart   bool   `json:"restart"`
}

// ParagraphStyle 段落样式
type ParagraphStyle struct {
	Name string `json:"name"`
	ID   string `json:"id"`
	// 添加字体信息
	Font     Font       `json:"font"`
	Alignment Alignment  `json:"alignment"`
	Indentation Indentation `json:"indentation"`
	Spacing Spacing    `json:"spacing"`
}

type TableStyle struct {
	Name string `json:"name"`
	ID   string `json:"id"`
}

type TableBorders struct {
	Top    Border `json:"top"`
	Bottom Border `json:"bottom"`
	Left   Border `json:"left"`
	Right  Border `json:"right"`
	InsideH Border `json:"inside_h"`
	InsideV Border `json:"inside_v"`
}

type TableShading struct {
	Fill   Color `json:"fill"`
	Pattern Pattern `json:"pattern"`
}

type CellBorders struct {
	Top    Border `json:"top"`
	Bottom Border `json:"bottom"`
	Left   Border `json:"left"`
	Right  Border `json:"right"`
}

type CellShading struct {
	Fill   Color `json:"fill"`
	Pattern Pattern `json:"pattern"`
}

type DocumentStyles struct {
	ParagraphStyles []ParagraphStyle `json:"paragraph_styles"`
	CharacterStyles []CharacterStyle `json:"character_styles"`
	TableStyles     []TableStyle     `json:"table_styles"`
}

type CharacterStyle struct {
	Name string `json:"name"`
	ID   string `json:"id"`
}

type TableRule struct {
	ID       string `json:"id"`
	Name     string `json:"name"`
	Borders  TableBorders `json:"borders"`
	Shading  TableShading `json:"shading"`
	Width    float64 `json:"width"`
	Alignment Alignment `json:"alignment"`
}

type PageRule struct {
	ID              string        `json:"id"`
	Name            string        `json:"name"`
	PageSize        PageSize      `json:"page_size"`
	PageMargins     PageMargins   `json:"page_margins"`
	HeaderDistance  float64       `json:"header_distance"`
	FooterDistance  float64       `json:"footer_distance"`
	Columns         Columns       `json:"columns"`
	PageNumbering   PageNumbering `json:"page_numbering"`
	LineNumbering   LineNumbering `json:"line_numbering"`
}

type StyleRule struct {
	ID          string `json:"id"`
	Name        string `json:"name"`
	Type        string `json:"type"`
	BasedOn     string `json:"based_on"`
	Next        string `json:"next"`
	Linked      string `json:"linked"`
	QuickFormat bool   `json:"quick_format"`
	Hidden      bool   `json:"hidden"`
}

// FormatIssue 格式问题
type FormatIssue struct {
	ID          string      `json:"id"`
	Type        string      `json:"type"`
	Severity    string      `json:"severity"`
	Location    string      `json:"location"`
	Description string      `json:"description"`
	Current     interface{} `json:"current"`
	Expected    interface{} `json:"expected"`
	Rule        string      `json:"rule"`
	Suggestions []string    `json:"suggestions"`
} 