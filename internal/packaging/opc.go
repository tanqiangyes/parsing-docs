package packaging

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

// OPCContainer 表示Open Packaging Convention容器
type OPCContainer struct {
	Path     string
	Reader   *zip.ReadCloser
	Files    map[string]*zip.File
	Metadata map[string]interface{}
}

// NewOPCContainer 创建新的OPC容器
func NewOPCContainer(path string) *OPCContainer {
	return &OPCContainer{
		Path:     path,
		Files:    make(map[string]*zip.File),
		Metadata: make(map[string]interface{}),
	}
}

// Open 打开OPC容器
func (oc *OPCContainer) Open() error {
	reader, err := zip.OpenReader(oc.Path)
	if err != nil {
		return fmt.Errorf("failed to open OPC container: %w", err)
	}
	
	oc.Reader = reader

	// 索引所有文件
	for _, file := range reader.File {
		oc.Files[file.Name] = file
	}

	return nil
}

// GetFile 获取指定文件
func (oc *OPCContainer) GetFile(name string) (*zip.File, error) {
	if file, exists := oc.Files[name]; exists {
		return file, nil
	}
	return nil, fmt.Errorf("file not found: %s", name)
}

// GetFiles 获取匹配的文件
func (oc *OPCContainer) GetFiles(prefix string) []*zip.File {
	var files []*zip.File
	for name, file := range oc.Files {
		if strings.HasPrefix(name, prefix) {
			files = append(files, file)
		}
	}
	return files
}

// ReadFile 读取文件内容
func (oc *OPCContainer) ReadFile(name string) ([]byte, error) {
	file, err := oc.GetFile(name)
	if err != nil {
		return nil, err
	}

	rc, err := file.Open()
	if err != nil {
		return nil, fmt.Errorf("failed to open file %s: %w", name, err)
	}
	defer rc.Close()

	return io.ReadAll(rc)
}

// HasFile 检查文件是否存在
func (oc *OPCContainer) HasFile(name string) bool {
	_, exists := oc.Files[name]
	return exists
}

// GetContentTypes 获取内容类型映射
func (oc *OPCContainer) GetContentTypes() (map[string]string, error) {
	content, err := oc.ReadFile("[Content_Types].xml")
	if err != nil {
		return nil, fmt.Errorf("failed to read content types: %w", err)
	}

	// 简单的XML解析，提取内容类型映射
	contentStr := string(content)
	types := make(map[string]string)

	// 解析Override元素
	lines := strings.Split(contentStr, "\n")
	for _, line := range lines {
		if strings.Contains(line, "Override") {
			// 提取PartName和ContentType
			if strings.Contains(line, "PartName=") && strings.Contains(line, "ContentType=") {
				// 简化解析，实际应该使用XML解析器
				parts := strings.Split(line, " ")
				for _, part := range parts {
					if strings.HasPrefix(part, "PartName=\"") {
						partName := strings.TrimPrefix(part, "PartName=\"")
						partName = strings.TrimSuffix(partName, "\"")
						types[partName] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
					}
				}
			}
		}
	}

	return types, nil
}

// GetRelationships 获取关系文件
func (oc *OPCContainer) GetRelationships() (map[string]string, error) {
	content, err := oc.ReadFile("_rels/.rels")
	if err != nil {
		return nil, fmt.Errorf("failed to read relationships: %w", err)
	}

	// 简单的XML解析，提取关系映射
	contentStr := string(content)
	relationships := make(map[string]string)

	lines := strings.Split(contentStr, "\n")
	for _, line := range lines {
		if strings.Contains(line, "Relationship") {
			// 提取Id和Target
			if strings.Contains(line, "Id=") && strings.Contains(line, "Target=") {
				parts := strings.Split(line, " ")
				var id, target string
				for _, part := range parts {
					if strings.HasPrefix(part, "Id=\"") {
						id = strings.TrimPrefix(part, "Id=\"")
						id = strings.TrimSuffix(id, "\"")
					}
					if strings.HasPrefix(part, "Target=\"") {
						target = strings.TrimPrefix(part, "Target=\"")
						target = strings.TrimSuffix(target, "\"")
					}
				}
				if id != "" && target != "" {
					relationships[id] = target
				}
			}
		}
	}

	return relationships, nil
}

// Validate 验证OPC容器
func (oc *OPCContainer) Validate() error {
	// 检查必需的文件
	requiredFiles := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"word/document.xml",
	}

	for _, file := range requiredFiles {
		if !oc.HasFile(file) {
			return fmt.Errorf("required file missing: %s", file)
		}
	}

	return nil
}

// Close 关闭OPC容器
func (oc *OPCContainer) Close() error {
	// 清理资源
	if oc.Reader != nil {
		oc.Reader.Close()
		oc.Reader = nil
	}
	oc.Files = nil
	oc.Metadata = nil
	return nil
} 