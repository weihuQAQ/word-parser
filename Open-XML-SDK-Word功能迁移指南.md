# Open XML SDK 项目分析与 Word 功能迁移到 JavaScript/TypeScript 指南

> 文档创建时间：2025 年 10 月 20 日  
> 基于项目：Open-XML-SDK-main  
> 目标：将 Word 处理核心功能迁移到 JavaScript/TypeScript

---

## 📊 项目概览

**Open XML SDK** 是一个由 Microsoft 开发的 .NET 开源框架，用于处理 Microsoft Office Word、Excel 和 PowerPoint 文档。这是一个 .NET Foundation 项目，采用 MIT 许可证。

### 核心功能

该 SDK 主要提供以下能力：

- ✅ 高性能生成 Word、Excel 和 PowerPoint 文档
- ✅ 文档修改（添加、更新、删除内容和元数据）
- ✅ 使用正则表达式搜索和替换内容
- ✅ 文件拆分与合并
- ✅ 更新 Word/PowerPoint 中图表的缓存数据和嵌入式电子表格

### 技术栈

- **.NET SDK**: 9.0.100
- **C# 语言版本**: 13
- **目标框架**: .NET Standard 2.0, .NET Framework 3.5/4.0/4.6, .NET 6.0/8.0
- **核心依赖**:
  - System.IO.Packaging 8.0.1
  - System.Collections.Immutable 8.0.0
  - System.Text.Json 9.0.0
  - Microsoft.CodeAnalysis 4.11.0（用于源代码生成）

### 项目结构

```
Open-XML-SDK-main/
├── src/                          # 源代码
│   ├── DocumentFormat.OpenXml.Framework      # 核心框架层（320个文件）
│   ├── DocumentFormat.OpenXml                # 主要库（59个文件）
│   ├── DocumentFormat.OpenXml.Linq           # LINQ 支持
│   └── DocumentFormat.OpenXml.Features       # 扩展特性
├── gen/                          # 代码生成器
│   ├── DocumentFormat.OpenXml.Generator        # Roslyn 源代码生成器
│   └── DocumentFormat.OpenXml.Generator.Models # 生成器模型
├── generated/                    # 生成的代码（395个文件）
│   ├── DocumentFormat.OpenXml/               # 279个生成的文件
│   └── DocumentFormat.OpenXml.Linq/          # 114个生成的文件
├── test/                         # 测试项目
│   └── DocumentFormat.OpenXml.Tests.Assets/  # 901个测试文件
├── data/                         # 数据定义（440个JSON文件）
│   ├── namespaces.json           # 命名空间映射
│   ├── schemas/                  # 155个架构定义
│   ├── parts/                    # 128个部件定义
│   └── typed/                    # 157个类型定义
└── samples/                      # 示例项目（10+个）
```

### 关键技术特点

#### 1. 源代码生成器 (Source Generator)

项目使用 Roslyn 源代码生成器在编译时生成大量代码：

- 279 个生成的 Word/Excel/PowerPoint 类
- 114 个生成的 LINQ 扩展类
- 基于 JSON schemas 驱动生成

#### 2. 特性系统 (Features System)

采用类似 ASP.NET Core 的特性模式，实现策略模式：

- **IDisposableFeature** - 资源释放管理
- **IPackageEventsFeature** - 包事件通知
- **IPartEventsFeature** - 部件事件通知
- **IPartRootEventsFeature** - 部件根元素事件
- **IParagraphIdGeneratorFeature** - 段落 ID 生成
- **IPartRootXElementFeature** - XLinq 集成

#### 3. 文档类型支持

- **WordprocessingDocument** - Word 文档处理
- **SpreadsheetDocument** - Excel 文档处理
- **PresentationDocument** - PowerPoint 文档处理

每种文档类型都支持常规 OOXML 格式和 Flat OPC 格式。

---

## 🏗️ Word 文档底层结构

### .docx 文件格式

```
.docx 文件 = ZIP 压缩包 + OPC (Open Packaging Convention)
├── [Content_Types].xml          # 定义所有部件的内容类型
├── _rels/
│   └── .rels                     # 包级别关系
└── word/
    ├── document.xml              # 主文档内容
    ├── styles.xml                # 样式定义
    ├── numbering.xml             # 编号定义
    ├── settings.xml              # 文档设置
    ├── fontTable.xml             # 字体表
    ├── _rels/
    │   └── document.xml.rels     # 文档关系（图片、页眉、页脚等）
    ├── header1.xml               # 页眉
    ├── footer1.xml               # 页脚
    └── media/                    # 嵌入的媒体文件
        ├── image1.png
        └── image2.jpg
```

### 核心概念

#### 1. OPC (Open Packaging Convention)

- 基于 ZIP 的容器格式
- 通过关系（Relationships）连接各个部件
- 每个部件有唯一的 URI 和 Content Type

#### 2. Parts（部件）

```json
// MainDocumentPart.json 定义
{
  "Name": "MainDocumentPart",
  "RelationshipType": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  "Target": "document",
  "RootElement": "document",
  "Children": [
    "CustomXmlParts",
    "GlossaryDocumentPart",
    "ThemePart",
    "WordprocessingCommentsPart",
    "DocumentSettingsPart",
    "StyleDefinitionsPart",
    "NumberingDefinitionsPart",
    "HeaderParts",
    "FooterParts",
    "ImageParts",
    "ChartParts"
  ]
}
```

#### 3. Relationships（关系）

```xml
<!-- _rels/.rels 示例 -->
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
</Relationships>
```

---

## 🎯 迁移到 JavaScript/TypeScript 的准备工作

### 一、核心技术栈选型

#### .NET 依赖 → JavaScript 替代方案

| .NET 核心依赖                  | JavaScript/TypeScript 替代        | 说明                  |
| ------------------------------ | --------------------------------- | --------------------- |
| `System.IO.Packaging`          | **JSZip** / ADM-ZIP               | OPC 包管理和 ZIP 处理 |
| `System.IO.Compression`        | JSZip 内置                        | ZIP 压缩/解压         |
| `System.Xml`                   | **fast-xml-parser** / xmlbuilder2 | XML 解析和生成        |
| `System.Collections.Immutable` | Immutable.js（可选）              | 不可变集合            |
| `System.IO.Stream`             | Node.js Streams / Web Streams     | 流式处理              |

#### 推荐的 NPM 包

```json
{
  "dependencies": {
    "jszip": "^3.10.1", // ZIP 处理（必需）
    "fast-xml-parser": "^4.3.2", // XML 解析（性能最佳）
    "xmlbuilder2": "^3.1.1", // XML 生成
    "uuid": "^9.0.1" // 生成唯一 ID
  },
  "devDependencies": {
    "typescript": "^5.3.0",
    "@types/node": "^20.0.0",
    "vitest": "^1.0.0", // 测试框架
    "prettier": "^3.0.0",
    "eslint": "^8.0.0"
  }
}
```

### 二、核心架构设计

#### 1. 类层次结构

```typescript
// ============================================
// 包管理层
// ============================================
interface IPackage {
  fileOpenAccess: "Read" | "Write" | "ReadWrite";
  packageProperties: IPackageProperties;
  parts: Map<string, IPackagePart>;
  relationships: IRelationshipCollection;
}

interface IPackagePart {
  uri: string;
  contentType: string;
  getStream(): Promise<ReadableStream>;
}

interface IRelationship {
  id: string;
  type: string;
  target: string;
  targetMode: "Internal" | "External";
}

// ============================================
// 文档类
// ============================================
class WordprocessingDocument {
  documentType: "Document" | "Template" | "MacroEnabledDocument" | "MacroEnabledTemplate";
  mainDocumentPart?: MainDocumentPart;
  coreFilePropertiesPart?: CoreFilePropertiesPart;
  extendedFilePropertiesPart?: ExtendedFilePropertiesPart;

  static async create(path: string, type: DocumentType): Promise<WordprocessingDocument>;
  static async open(path: string, isEditable: boolean): Promise<WordprocessingDocument>;
  async save(): Promise<void>;
  close(): void;
}

// ============================================
// 部件（Parts）系统
// ============================================
abstract class OpenXmlPart {
  uri: string;
  contentType: string;
  relationshipType: string;
  relationships: IRelationshipCollection;

  abstract getRootElement(): OpenXmlElement;
  abstract async loadAsync(): Promise<void>;
  abstract async saveAsync(): Promise<void>;
}

class MainDocumentPart extends OpenXmlPart {
  document?: Document;
  stylesPart?: StyleDefinitionsPart;
  numberingPart?: NumberingDefinitionsPart;
  fontTablePart?: FontTablePart;
  headerParts: HeaderPart[] = [];
  footerParts: FooterPart[] = [];
  imageParts: ImagePart[] = [];

  addHeaderPart(): HeaderPart;
  addFooterPart(): FooterPart;
  addImagePart(contentType: string): ImagePart;
}

// ============================================
// 元素系统
// ============================================
abstract class OpenXmlElement {
  parent?: OpenXmlElement;
  localName: string;
  namespaceUri: string;
  prefix: string;

  private attributes: Map<string, OpenXmlAttribute>;
  private extendedAttributes: OpenXmlAttribute[] = [];

  // 核心方法
  appendChild(child: OpenXmlElement): void;
  removeChild(child: OpenXmlElement): void;
  insertBefore(newChild: OpenXmlElement, refChild: OpenXmlElement): void;
  clone(): OpenXmlElement;

  // XML 序列化
  toXml(): string;
  static fromXml(xml: string): OpenXmlElement;

  // 属性操作
  getAttribute(name: string): string | undefined;
  setAttribute(name: string, value: string): void;

  // 遍历
  descendants(): Iterable<OpenXmlElement>;
  ancestors(): Iterable<OpenXmlElement>;
}

// 叶子元素（无子元素）
abstract class OpenXmlLeafElement extends OpenXmlElement {
  // 不能有子元素
}

// 叶子文本元素
abstract class OpenXmlLeafTextElement extends OpenXmlLeafElement {
  text: string;
}

// 组合元素（有子元素）
abstract class OpenXmlCompositeElement extends OpenXmlElement {
  protected children: OpenXmlElement[] = [];

  get firstChild(): OpenXmlElement | undefined;
  get lastChild(): OpenXmlElement | undefined;

  appendChildren(...elements: OpenXmlElement[]): void;
  removeAllChildren(): void;
}

// ============================================
// 具体 Word 元素示例
// ============================================
class Document extends OpenXmlCompositeElement {
  body?: Body;

  constructor() {
    super();
    this.localName = "document";
    this.namespaceUri = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    this.prefix = "w";
  }
}

class Body extends OpenXmlCompositeElement {
  paragraphs: Paragraph[] = [];
  tables: Table[] = [];

  addParagraph(): Paragraph {
    const p = new Paragraph();
    this.appendChild(p);
    this.paragraphs.push(p);
    return p;
  }
}

class Paragraph extends OpenXmlCompositeElement {
  paragraphProperties?: ParagraphProperties;
  runs: Run[] = [];

  addRun(): Run {
    const r = new Run();
    this.appendChild(r);
    this.runs.push(r);
    return r;
  }
}

class Run extends OpenXmlCompositeElement {
  runProperties?: RunProperties;
  texts: Text[] = [];

  addText(content: string): Text {
    const t = new Text(content);
    this.appendChild(t);
    this.texts.push(t);
    return t;
  }
}

class Text extends OpenXmlLeafTextElement {
  constructor(text: string = "") {
    super();
    this.localName = "t";
    this.namespaceUri = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    this.prefix = "w";
    this.text = text;
  }
}

// ============================================
// 属性类
// ============================================
class RunProperties extends OpenXmlCompositeElement {
  bold?: Bold;
  italic?: Italic;
  fontSize?: FontSize;
  color?: Color;

  setBold(value: boolean): void {
    if (value) {
      this.bold = new Bold();
      this.appendChild(this.bold);
    } else {
      if (this.bold) this.removeChild(this.bold);
      this.bold = undefined;
    }
  }
}
```

#### 2. 关系系统实现

```typescript
interface IRelationship {
  id: string;
  type: string;
  target: string;
  targetMode: "Internal" | "External";
}

class RelationshipCollection implements IRelationshipCollection {
  private relationships: Map<string, IRelationship> = new Map();
  private idCounter: number = 1;

  add(type: string, target: string, targetMode: "Internal" | "External" = "Internal"): IRelationship {
    const id = `rId${this.idCounter++}`;
    const rel: IRelationship = { id, type, target, targetMode };
    this.relationships.set(id, rel);
    return rel;
  }

  getById(id: string): IRelationship | undefined {
    return this.relationships.get(id);
  }

  getByType(type: string): IRelationship[] {
    return Array.from(this.relationships.values()).filter((rel) => rel.type === type);
  }

  toXml(): string {
    const builder = new XmlBuilder();
    builder.startElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");

    for (const rel of this.relationships.values()) {
      builder.startElement("Relationship");
      builder.addAttribute("Id", rel.id);
      builder.addAttribute("Type", rel.type);
      builder.addAttribute("Target", rel.target);
      if (rel.targetMode === "External") {
        builder.addAttribute("TargetMode", "External");
      }
      builder.endElement();
    }

    builder.endElement();
    return builder.toString();
  }

  static fromXml(xml: string): RelationshipCollection {
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
    const parsed = parser.parse(xml);
    const collection = new RelationshipCollection();

    const rels = parsed.Relationships?.Relationship;
    if (!rels) return collection;

    const relArray = Array.isArray(rels) ? rels : [rels];
    for (const rel of relArray) {
      collection.relationships.set(rel["@_Id"], {
        id: rel["@_Id"],
        type: rel["@_Type"],
        target: rel["@_Target"],
        targetMode: rel["@_TargetMode"] || "Internal",
      });
    }

    return collection;
  }
}

// 常见关系类型常量
const RelationshipTypes = {
  OFFICE_DOCUMENT: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  STYLES: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  NUMBERING: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
  FONT_TABLE: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
  SETTINGS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
  IMAGE: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  HEADER: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  FOOTER: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
  HYPERLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  COMMENTS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  ENDNOTES: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
  FOOTNOTES: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
};
```

#### 3. Content Types 管理

```typescript
class ContentTypeManager {
  private defaults: Map<string, string> = new Map([
    ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
    ["xml", "application/xml"],
    ["png", "image/png"],
    ["jpg", "image/jpeg"],
    ["jpeg", "image/jpeg"],
    ["gif", "image/gif"],
  ]);

  private overrides: Map<string, string> = new Map();

  addDefault(extension: string, contentType: string): void {
    this.defaults.set(extension, contentType);
  }

  addOverride(partName: string, contentType: string): void {
    // 确保以 / 开头
    if (!partName.startsWith("/")) {
      partName = "/" + partName;
    }
    this.overrides.set(partName, contentType);
  }

  getContentType(partName: string): string | undefined {
    // 先检查 override
    if (this.overrides.has(partName)) {
      return this.overrides.get(partName);
    }

    // 然后检查扩展名 default
    const extension = partName.split(".").pop()?.toLowerCase();
    if (extension && this.defaults.has(extension)) {
      return this.defaults.get(extension);
    }

    return undefined;
  }

  toXml(): string {
    const builder = new XmlBuilder();
    builder.startElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");

    // 添加 defaults
    for (const [ext, contentType] of this.defaults) {
      builder.startElement("Default");
      builder.addAttribute("Extension", ext);
      builder.addAttribute("ContentType", contentType);
      builder.endElement();
    }

    // 添加 overrides
    for (const [partName, contentType] of this.overrides) {
      builder.startElement("Override");
      builder.addAttribute("PartName", partName);
      builder.addAttribute("ContentType", contentType);
      builder.endElement();
    }

    builder.endElement();
    return builder.toString();
  }

  static fromXml(xml: string): ContentTypeManager {
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
    const parsed = parser.parse(xml);
    const manager = new ContentTypeManager();

    // 解析 defaults
    const defaults = parsed.Types?.Default;
    if (defaults) {
      const defaultArray = Array.isArray(defaults) ? defaults : [defaults];
      for (const def of defaultArray) {
        manager.defaults.set(def["@_Extension"], def["@_ContentType"]);
      }
    }

    // 解析 overrides
    const overrides = parsed.Types?.Override;
    if (overrides) {
      const overrideArray = Array.isArray(overrides) ? overrides : [overrides];
      for (const override of overrideArray) {
        manager.overrides.set(override["@_PartName"], override["@_ContentType"]);
      }
    }

    return manager;
  }
}

// 预定义的 Content Types
const ContentTypes = {
  // Package
  RELATIONSHIPS: "application/vnd.openxmlformats-package.relationships+xml",

  // WordprocessingML
  DOCUMENT: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  TEMPLATE: "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
  STYLES: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  NUMBERING: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
  SETTINGS: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
  FONT_TABLE: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
  HEADER: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
  FOOTER: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",

  // Images
  PNG: "image/png",
  JPEG: "image/jpeg",
  GIF: "image/gif",
};
```

#### 4. ZIP 包管理器

```typescript
import JSZip from "jszip";
import { promises as fs } from "fs";

class PackageManager {
  private zip: JSZip;
  private contentTypes: ContentTypeManager;

  constructor() {
    this.zip = new JSZip();
    this.contentTypes = new ContentTypeManager();
  }

  // 打开现有包
  async open(path: string): Promise<void> {
    const data = await fs.readFile(path);
    this.zip = await JSZip.loadAsync(data);

    // 加载 [Content_Types].xml
    const contentTypesXml = await this.zip.file("[Content_Types].xml")?.async("string");
    if (contentTypesXml) {
      this.contentTypes = ContentTypeManager.fromXml(contentTypesXml);
    }
  }

  // 创建新包
  createNew(): void {
    this.zip = new JSZip();
    this.contentTypes = new ContentTypeManager();

    // 添加基本的 _rels/.rels
    const rels = new RelationshipCollection();
    this.addPart("_rels/.rels", ContentTypes.RELATIONSHIPS, rels.toXml());
  }

  // 添加部件
  async addPart(uri: string, contentType: string, content: Buffer | string): Promise<void> {
    // 去除开头的 /
    if (uri.startsWith("/")) {
      uri = uri.substring(1);
    }

    this.zip.file(uri, content);

    // 更新 content types
    const extension = uri.split(".").pop()?.toLowerCase();
    if (extension && !this.contentTypes.getContentType(uri)) {
      this.contentTypes.addOverride("/" + uri, contentType);
    }
  }

  // 获取部件
  async getPart(uri: string): Promise<Buffer | null> {
    if (uri.startsWith("/")) {
      uri = uri.substring(1);
    }

    const file = this.zip.file(uri);
    if (!file) return null;

    return await file.async("nodebuffer");
  }

  // 获取部件内容（字符串）
  async getPartString(uri: string): Promise<string | null> {
    if (uri.startsWith("/")) {
      uri = uri.substring(1);
    }

    const file = this.zip.file(uri);
    if (!file) return null;

    return await file.async("string");
  }

  // 检查部件是否存在
  hasPart(uri: string): boolean {
    if (uri.startsWith("/")) {
      uri = uri.substring(1);
    }
    return this.zip.file(uri) !== null;
  }

  // 删除部件
  removePart(uri: string): void {
    if (uri.startsWith("/")) {
      uri = uri.substring(1);
    }
    this.zip.remove(uri);
  }

  // 保存包
  async save(path: string): Promise<void> {
    // 更新 [Content_Types].xml
    this.zip.file("[Content_Types].xml", this.contentTypes.toXml());

    // 生成 ZIP
    const content = await this.zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 9 },
    });

    await fs.writeFile(path, content);
  }

  // 保存为 Buffer
  async saveToBuffer(): Promise<Buffer> {
    this.zip.file("[Content_Types].xml", this.contentTypes.toXml());

    return await this.zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 9 },
    });
  }

  // 列出所有部件
  listParts(): string[] {
    const parts: string[] = [];
    this.zip.forEach((relativePath, file) => {
      if (!file.dir) {
        parts.push(relativePath);
      }
    });
    return parts;
  }
}
```

### 三、命名空间管理

```typescript
// 命名空间定义（基于 data/namespaces.json）
const Namespaces = {
  // 主要 WordprocessingML 命名空间
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",

  // 关系
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",

  // DrawingML
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",

  // 数学
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",

  // 包
  pkg: "http://schemas.microsoft.com/office/2006/xmlPackage",

  // Office 特定
  o: "urn:schemas-microsoft-com:office:office",
  v: "urn:schemas-microsoft-com:vml",

  // Word 版本特定
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  w16: "http://schemas.microsoft.com/office/word/2018/wordml",

  // XML
  xml: "http://www.w3.org/XML/1998/namespace",
};

class NamespaceManager {
  private prefixToUri: Map<string, string> = new Map();
  private uriToPrefix: Map<string, string> = new Map();

  constructor() {
    // 注册默认命名空间
    for (const [prefix, uri] of Object.entries(Namespaces)) {
      this.registerNamespace(prefix, uri);
    }
  }

  registerNamespace(prefix: string, uri: string): void {
    this.prefixToUri.set(prefix, uri);
    this.uriToPrefix.set(uri, prefix);
  }

  getPrefix(uri: string): string | undefined {
    return this.uriToPrefix.get(uri);
  }

  getUri(prefix: string): string | undefined {
    return this.prefixToUri.get(prefix);
  }

  getQName(namespaceUri: string, localName: string): string {
    const prefix = this.getPrefix(namespaceUri);
    return prefix ? `${prefix}:${localName}` : localName;
  }
}
```

### 四、XML 序列化实现

```typescript
import { XMLParser, XMLBuilder } from "fast-xml-parser";

class XmlSerializer {
  private parser: XMLParser;
  private builder: XMLBuilder;
  private namespaceManager: NamespaceManager;

  constructor() {
    this.namespaceManager = new NamespaceManager();

    this.parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      parseAttributeValue: false,
      trimValues: false,
      preserveOrder: true,
      allowBooleanAttributes: true,
    });

    this.builder = new XMLBuilder({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      format: true,
      indentBy: "  ",
      suppressEmptyNode: true,
    });
  }

  // 将 OpenXmlElement 序列化为 XML 字符串
  serialize(element: OpenXmlElement): string {
    const xmlObj = this.elementToObject(element);
    return this.builder.build(xmlObj);
  }

  // 将 OpenXmlElement 转换为对象
  private elementToObject(element: OpenXmlElement): any {
    const qName = this.namespaceManager.getQName(element.namespaceUri, element.localName);
    const obj: any = {
      [qName]: {},
    };

    // 添加属性
    for (const [name, value] of element.getAttributes()) {
      obj[qName][`@_${name}`] = value;
    }

    // 处理子元素或文本内容
    if (element instanceof OpenXmlLeafTextElement) {
      obj[qName]["#text"] = element.text;
    } else if (element instanceof OpenXmlCompositeElement) {
      const children: any[] = [];
      for (const child of element.children) {
        children.push(this.elementToObject(child));
      }
      if (children.length > 0) {
        Object.assign(obj[qName], ...children);
      }
    }

    return obj;
  }

  // 从 XML 字符串反序列化为 OpenXmlElement
  deserialize(xml: string): OpenXmlElement {
    const parsed = this.parser.parse(xml);
    return this.objectToElement(parsed);
  }

  // 从对象转换为 OpenXmlElement
  private objectToElement(obj: any): OpenXmlElement {
    // 这里需要根据元素的 QName 创建对应的类实例
    // 实际实现需要一个元素工厂
    throw new Error("Not implemented - needs element factory");
  }
}
```

### 五、实施路线图

#### 阶段 1：基础设施（2-4 周）

- [x] 研究 Open XML SDK 架构
- [ ] 实现 ZIP 包管理器（PackageManager）
- [ ] 实现 XML 解析器和生成器
- [ ] 实现 Content Types 管理（ContentTypeManager）
- [ ] 实现关系管理器（RelationshipCollection）
- [ ] 搭建基本的包结构读取
- [ ] 编写单元测试

**验收标准：**

- 能够打开 .docx 文件
- 能够读取 [Content_Types].xml
- 能够读取 \_rels/.rels
- 能够列出所有部件

#### 阶段 2：核心对象模型（4-6 周）

- [ ] 实现 `OpenXmlElement` 基类层次
  - [ ] OpenXmlElement
  - [ ] OpenXmlLeafElement
  - [ ] OpenXmlLeafTextElement
  - [ ] OpenXmlCompositeElement
- [ ] 实现 `OpenXmlPart` 基类层次
  - [ ] OpenXmlPart
  - [ ] OpenXmlPartRootElement
- [ ] 实现 `WordprocessingDocument` 类
- [ ] 实现 `MainDocumentPart` 类
- [ ] 实现基本的 Word 元素
  - [ ] Document
  - [ ] Body
  - [ ] Paragraph
  - [ ] Run
  - [ ] Text
- [ ] 实现属性类
  - [ ] ParagraphProperties
  - [ ] RunProperties
- [ ] 编写集成测试

**验收标准：**

- 能够打开 .docx 文件并读取文档结构
- 能够遍历段落和运行
- 能够提取纯文本
- 能够创建新文档
- 能够添加段落和文本
- 能够保存文档

#### 阶段 3：代码生成器（3-4 周）

- [ ] 解析 JSON schema 定义
  - [ ] 读取 data/namespaces.json
  - [ ] 读取 data/parts/\*.json
  - [ ] 读取 data/schemas/\*.json
- [ ] 生成 TypeScript 类定义
  - [ ] 生成元素类
  - [ ] 生成属性类
  - [ ] 生成枚举类型
- [ ] 生成验证器
- [ ] 生成类型声明文件（.d.ts）
- [ ] 自动化构建流程

**验收标准：**

- 代码生成器能够运行
- 生成的代码通过 TypeScript 编译
- 生成的类可以实例化和使用

#### 阶段 4：高级功能（4-6 周）

- [ ] 样式管理
  - [ ] StyleDefinitionsPart
  - [ ] Style 元素
  - [ ] 段落样式
  - [ ] 字符样式
- [ ] 编号管理
  - [ ] NumberingDefinitionsPart
  - [ ] AbstractNum
  - [ ] NumberingInstance
- [ ] 页眉页脚
  - [ ] HeaderPart
  - [ ] FooterPart
  - [ ] HeaderReference
  - [ ] FooterReference
- [ ] 表格支持
  - [ ] Table
  - [ ] TableRow
  - [ ] TableCell
  - [ ] TableProperties
- [ ] 图片支持
  - [ ] ImagePart
  - [ ] Drawing
  - [ ] Inline / Anchor
- [ ] 超链接
  - [ ] Hyperlink 元素
  - [ ] 外部关系
- [ ] 书签和引用
  - [ ] BookmarkStart / BookmarkEnd
  - [ ] 交叉引用

**验收标准：**

- 能够应用和创建样式
- 能够创建编号列表
- 能够添加页眉页脚
- 能够创建表格
- 能够插入图片
- 能够添加超链接

#### 阶段 5：验证和优化（2-3 周）

- [ ] Schema 验证器
  - [ ] 元素结构验证
  - [ ] 属性验证
  - [ ] 数据类型验证
- [ ] 单元测试（覆盖率 > 80%）
- [ ] 集成测试
- [ ] 兼容性测试
  - [ ] 测试生成的文档能否被 MS Word 打开
  - [ ] 测试能否正确读取各种版本的 Word 文档
- [ ] 性能测试和优化
  - [ ] 大文档处理
  - [ ] 内存使用优化
  - [ ] 流式处理
- [ ] 文档编写

**验收标准：**

- 测试覆盖率达到 80%+
- 生成的文档能被 MS Word 正常打开
- 能够处理 100+ 页的大文档
- 性能达到可接受水平

### 六、最小可行产品（MVP）范围

**第一版（MVP）建议只支持：**

#### ✅ 基本读取

- 打开 .docx 文件
- 读取文档结构
- 提取纯文本
- 访问段落和运行
- 读取基本格式（加粗、斜体、字体）

#### ✅ 基本修改

- 创建新文档
- 添加段落
- 添加文本运行
- 设置基本格式
  - 加粗（Bold）
  - 斜体（Italic）
  - 字体大小（FontSize）
  - 字体颜色（Color）
- 保存文档

#### ⏸️ 延后功能（V2.0+）

- 复杂样式系统
- 表格完整支持
- 图片和媒体
- 页眉页脚
- 页面设置
- 编号和项目符号
- 完整的 Schema 验证
- 宏支持
- 修订跟踪

### 七、代码示例

#### 快速原型：最小实现

```typescript
import JSZip from "jszip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";
import { promises as fs } from "fs";

// ============================================
// 简化版 Word 文档类
// ============================================
class SimpleWordDocument {
  private zip!: JSZip;
  private parser: XMLParser;
  private builder: XMLBuilder;
  private documentXml: any;

  constructor() {
    this.parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      parseAttributeValue: false,
    });

    this.builder = new XMLBuilder({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      format: true,
    });
  }

  // 打开现有文档
  async open(path: string): Promise<void> {
    const data = await fs.readFile(path);
    this.zip = await JSZip.loadAsync(data);

    // 读取主文档
    const docXmlString = await this.zip.file("word/document.xml")?.async("string");
    if (!docXmlString) {
      throw new Error("document.xml not found");
    }

    this.documentXml = this.parser.parse(docXmlString);
  }

  // 创建新文档
  async create(): Promise<void> {
    this.zip = new JSZip();

    // 创建基本结构
    this.createContentTypes();
    this.createRels();
    this.createDocument();
  }

  private createContentTypes(): void {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
    this.zip.file("[Content_Types].xml", xml);
  }

  private createRels(): void {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
    this.zip.file("_rels/.rels", xml);
  }

  private createDocument(): void {
    this.documentXml = {
      "?xml": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" },
      "w:document": {
        "@_xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "w:body": {
          "w:p": [],
        },
      },
    };
  }

  // 获取所有文本
  getText(): string {
    const body = this.documentXml["w:document"]["w:body"];
    const paragraphs = body["w:p"];

    if (!paragraphs) return "";

    const pArray = Array.isArray(paragraphs) ? paragraphs : [paragraphs];

    return pArray
      .map((p: any) => {
        const runs = p["w:r"];
        if (!runs) return "";

        const rArray = Array.isArray(runs) ? runs : [runs];
        return rArray
          .map((r: any) => {
            const text = r?.["w:t"];
            if (!text) return "";

            return typeof text === "object" ? text["#text"] || "" : text;
          })
          .join("");
      })
      .join("\n");
  }

  // 添加段落
  addParagraph(text: string): void {
    const body = this.documentXml["w:document"]["w:body"];

    if (!body["w:p"]) {
      body["w:p"] = [];
    }

    const paragraphs = body["w:p"];
    const pArray = Array.isArray(paragraphs) ? paragraphs : [paragraphs];

    const newParagraph = {
      "w:r": {
        "w:t": {
          "@_xml:space": "preserve",
          "#text": text,
        },
      },
    };

    pArray.push(newParagraph);
    body["w:p"] = pArray;
  }

  // 保存文档
  async save(path: string): Promise<void> {
    // 更新 document.xml
    const docXmlString = this.builder.build(this.documentXml);
    this.zip.file("word/document.xml", docXmlString);

    // 生成 ZIP
    const content = await this.zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
    });

    await fs.writeFile(path, content);
  }
}

// ============================================
// 使用示例
// ============================================
async function example() {
  // 示例 1：读取现有文档
  console.log("=== 读取文档 ===");
  const doc1 = new SimpleWordDocument();
  await doc1.open("test.docx");
  const text = doc1.getText();
  console.log("文档内容：", text);

  // 示例 2：创建新文档
  console.log("\n=== 创建新文档 ===");
  const doc2 = new SimpleWordDocument();
  await doc2.create();
  doc2.addParagraph("Hello, World!");
  doc2.addParagraph("这是第二段。");
  await doc2.save("output.docx");
  console.log("新文档已创建：output.docx");

  // 示例 3：修改现有文档
  console.log("\n=== 修改文档 ===");
  const doc3 = new SimpleWordDocument();
  await doc3.open("test.docx");
  doc3.addParagraph("这是新增的段落。");
  await doc3.save("modified.docx");
  console.log("修改后的文档已保存：modified.docx");
}

// 运行示例
example().catch(console.error);
```

### 八、潜在挑战和解决方案

| 挑战              | C# 实现                  | JavaScript/TypeScript 解决方案          | 优先级 |
| ----------------- | ------------------------ | --------------------------------------- | ------ |
| **内存管理**      | 自动 GC，Disposable 模式 | 手动管理，WeakMap 缓存，及时释放引用    | 🔴 高  |
| **流式处理**      | System.IO.Stream         | Node.js Streams / Web Streams API       | 🟡 中  |
| **二进制数据**    | byte[]                   | Buffer (Node.js) / ArrayBuffer (浏览器) | 🔴 高  |
| **XML 命名空间**  | System.Xml 原生支持      | 需要自己实现命名空间映射和管理          | 🔴 高  |
| **Strong Typing** | 编译时类型检查           | TypeScript + 运行时验证（Zod/Yup）      | 🟡 中  |
| **性能**          | JIT 编译，原生代码       | V8 优化 + WebAssembly（必要时）         | 🟢 低  |
| **大文件处理**    | Stream 支持              | 分块处理，流式 API                      | 🟡 中  |
| **跨平台**        | .NET Runtime             | Node.js / 浏览器                        | 🟢 低  |
| **代码量**        | 395 个生成文件           | 相同或更多，需要代码生成器              | 🔴 高  |

#### 具体解决方案

##### 1. 内存管理

```typescript
class ResourceManager {
  private resources: Set<IDisposable> = new Set();

  register(resource: IDisposable): void {
    this.resources.add(resource);
  }

  dispose(): void {
    for (const resource of this.resources) {
      resource.dispose();
    }
    this.resources.clear();
  }
}

interface IDisposable {
  dispose(): void;
}

class WordprocessingDocument implements IDisposable {
  private resourceManager = new ResourceManager();

  dispose(): void {
    this.resourceManager.dispose();
    // 清理其他资源
  }
}

// 使用
async function processDocument() {
  const doc = await WordprocessingDocument.open("test.docx", true);
  try {
    // 处理文档
  } finally {
    doc.dispose(); // 确保资源被释放
  }
}
```

##### 2. 大文件流式处理

```typescript
class StreamingReader {
  async *readParagraphs(doc: WordprocessingDocument): AsyncGenerator<Paragraph> {
    // 使用 SAX 风格的 XML 解析器
    const parser = new SaxParser();

    for await (const event of parser.parse(doc.getPartStream("word/document.xml"))) {
      if (event.type === "startElement" && event.name === "w:p") {
        yield this.parseParagraph(event);
      }
    }
  }
}

// 使用
async function processLargeDocument() {
  const doc = await WordprocessingDocument.open("large.docx", false);
  const reader = new StreamingReader();

  for await (const paragraph of reader.readParagraphs(doc)) {
    console.log(paragraph.getText());
  }
}
```

##### 3. 性能优化

```typescript
// 使用 WeakMap 缓存
class ElementCache {
  private cache = new WeakMap<any, OpenXmlElement>();

  get(key: any): OpenXmlElement | undefined {
    return this.cache.get(key);
  }

  set(key: any, element: OpenXmlElement): void {
    this.cache.set(key, element);
  }
}

// 延迟加载
class MainDocumentPart extends OpenXmlPart {
  private _document?: Document;

  get document(): Document {
    if (!this._document) {
      this._document = this.loadDocument();
    }
    return this._document;
  }

  private loadDocument(): Document {
    // 只在需要时才解析
  }
}
```

### 九、测试策略

#### 测试文件准备

```bash
# 使用 Open XML SDK 的测试资产
test-assets/
├── basic/
│   ├── empty.docx                 # 空白文档
│   ├── hello-world.docx           # 简单文档
│   └── multi-paragraph.docx       # 多段落
├── formatting/
│   ├── bold-italic.docx           # 文本格式
│   ├── fonts.docx                 # 字体
│   └── colors.docx                # 颜色
├── structures/
│   ├── tables.docx                # 表格
│   ├── images.docx                # 图片
│   └── headers-footers.docx       # 页眉页脚
├── versions/
│   ├── office2007.docx
│   ├── office2010.docx
│   ├── office2013.docx
│   └── office2016.docx
└── edge-cases/
    ├── corrupted.docx             # 损坏的文件
    ├── large.docx                 # 大文件（100+ 页）
    └── complex.docx               # 复杂结构
```

#### 单元测试示例

```typescript
import { describe, it, expect } from "vitest";

describe("WordprocessingDocument", () => {
  describe("create", () => {
    it("should create a new document", async () => {
      const doc = await WordprocessingDocument.create("test.docx", "Document");
      expect(doc).toBeDefined();
      expect(doc.documentType).toBe("Document");
      expect(doc.mainDocumentPart).toBeDefined();
    });
  });

  describe("open", () => {
    it("should open an existing document", async () => {
      const doc = await WordprocessingDocument.open("test-assets/basic/hello-world.docx", false);
      expect(doc).toBeDefined();
      expect(doc.mainDocumentPart).toBeDefined();
    });

    it("should throw error for non-existent file", async () => {
      await expect(WordprocessingDocument.open("non-existent.docx", false)).rejects.toThrow();
    });
  });
});

describe("MainDocumentPart", () => {
  it("should read paragraphs", async () => {
    const doc = await WordprocessingDocument.open("test-assets/basic/multi-paragraph.docx", false);
    const paragraphs = doc.mainDocumentPart?.document?.body?.paragraphs;

    expect(paragraphs).toBeDefined();
    expect(paragraphs!.length).toBeGreaterThan(0);
  });

  it("should add paragraph", async () => {
    const doc = await WordprocessingDocument.create("temp.docx", "Document");
    const body = doc.mainDocumentPart!.document!.body!;

    const p = body.addParagraph();
    p.addRun().addText("Hello, World!");

    expect(body.paragraphs.length).toBe(1);
    expect(body.paragraphs[0].getText()).toBe("Hello, World!");
  });
});

describe("Paragraph", () => {
  it("should get text", () => {
    const p = new Paragraph();
    p.addRun().addText("Hello ");
    p.addRun().addText("World!");

    expect(p.getText()).toBe("Hello World!");
  });

  it("should clone", () => {
    const p = new Paragraph();
    p.addRun().addText("Original");

    const clone = p.clone() as Paragraph;
    expect(clone.getText()).toBe("Original");
    expect(clone).not.toBe(p);
  });
});

describe("Run", () => {
  it("should apply bold", () => {
    const r = new Run();
    r.runProperties = new RunProperties();
    r.runProperties.setBold(true);

    expect(r.runProperties.bold).toBeDefined();
  });

  it("should set font size", () => {
    const r = new Run();
    r.runProperties = new RunProperties();
    r.runProperties.fontSize = new FontSize("24");

    expect(r.runProperties.fontSize?.val).toBe("24");
  });
});
```

#### 集成测试

```typescript
describe("Integration Tests", () => {
  it("should create, modify and save document", async () => {
    // 创建
    const doc = await WordprocessingDocument.create("temp.docx", "Document");
    const body = doc.mainDocumentPart!.document!.body!;

    // 添加内容
    const p1 = body.addParagraph();
    const r1 = p1.addRun();
    r1.runProperties = new RunProperties();
    r1.runProperties.setBold(true);
    r1.addText("Hello, ");

    const r2 = p1.addRun();
    r2.addText("World!");

    // 保存
    await doc.save();
    doc.dispose();

    // 重新打开验证
    const doc2 = await WordprocessingDocument.open("temp.docx", false);
    const text = doc2.mainDocumentPart!.document!.body!.getText();
    expect(text).toBe("Hello, World!");
    doc2.dispose();
  });

  it("should handle large document", async () => {
    const doc = await WordprocessingDocument.create("large.docx", "Document");
    const body = doc.mainDocumentPart!.document!.body!;

    // 添加 1000 个段落
    for (let i = 0; i < 1000; i++) {
      const p = body.addParagraph();
      p.addRun().addText(`Paragraph ${i + 1}`);
    }

    await doc.save();
    expect(body.paragraphs.length).toBe(1000);
  }, 30000); // 30秒超时
});
```

### 十、参考资源

#### 必读规范文档

1. **ISO/IEC 29500** - Office Open XML 文件格式标准

   - Part 1: Fundamentals and Markup Language Reference
   - Part 4: Transitional Migration Features (WordprocessingML)
   - 下载：https://standards.iso.org/ittf/PubliclyAvailableStandards/

2. **ECMA-376** - Office Open XML 标准（免费版本）

   - 与 ISO 29500 内容相同
   - 下载：https://www.ecma-international.org/publications-and-standards/standards/ecma-376/

3. **Open Packaging Conventions (OPC)**
   - ECMA-376 Part 2
   - 定义了 ZIP 包的结构和关系系统

#### 在线工具

- **Open XML SDK 2.5 Productivity Tool**

  - 可视化查看 .docx 文件结构
  - 生成 C# 代码
  - https://github.com/OfficeDev/Open-XML-SDK/releases/tag/v2.5

- **OOXML Viewer (VS Code 扩展)**
  - 在 VS Code 中查看和编辑 OOXML 文件
  - 支持 diff 功能

#### 现有 JavaScript 库参考

```bash
# 可以参考学习的库（但功能有限）
1. docxtemplater
   - 用途：模板填充
   - 优点：简单易用
   - 缺点：不支持完整的对象模型

2. officegen
   - 用途：文档生成
   - 优点：API 友好
   - 缺点：功能有限，不支持读取

3. docx (by dolanmiu)
   - 用途：文档创建
   - 优点：TypeScript，类型安全
   - 缺点：不支持完整读取和修改

4. mammoth.js
   - 用途：将 .docx 转为 HTML
   - 优点：转换质量高
   - 缺点：单向转换，无法生成 .docx

# 建议：研究这些库的源码，学习它们的实现思路，但要自己实现完整的对象模型
```

#### Open XML SDK 测试资产

```bash
# 使用原项目的测试文件
Open-XML-SDK-main/test/DocumentFormat.OpenXml.Tests.Assets/
├── 419 个 .docx 文件
├── 180 个 .pptx 文件
├── 107 个 .xlsx 文件
└── 各种边界情况和版本测试
```

#### 学习资源

- **官方文档**: https://learn.microsoft.com/office/open-xml/
- **GitHub**: https://github.com/OfficeDev/Open-XML-SDK
- **Stack Overflow**: 标签 `openxml` 和 `openxml-sdk`
- **博客系列**: Eric White 的 Open XML 博客（Archive）

---

## 📝 总结

### 工作量估算

| 阶段                 | 工作量       | 复杂度 | 优先级 |
| -------------------- | ------------ | ------ | ------ |
| 阶段 1：基础设施     | 2-4 周       | 🟡 中  | 🔴 P0  |
| 阶段 2：核心对象模型 | 4-6 周       | 🔴 高  | 🔴 P0  |
| 阶段 3：代码生成器   | 3-4 周       | 🔴 高  | 🟡 P1  |
| 阶段 4：高级功能     | 4-6 周       | 🔴 高  | 🟢 P2  |
| 阶段 5：验证和优化   | 2-3 周       | 🟡 中  | 🟡 P1  |
| **总计**             | **15-23 周** | -      | -      |

**全职开发：4-6 个月**  
**兼职开发：8-12 个月**

### 关键成功因素

1. ✅ **理解底层格式**

   - ZIP + XML + OPC 规范
   - 关系系统
   - Content Types 管理

2. ✅ **选对工具**

   - JSZip（ZIP 处理）
   - fast-xml-parser（XML 解析，性能最佳）
   - xmlbuilder2（XML 生成）

3. ✅ **数据驱动架构**

   - 使用 Open XML SDK 的 JSON schemas
   - 构建代码生成器
   - 减少手写代码

4. ✅ **分阶段实施**

   - 先做 MVP（基本读写）
   - 逐步扩展功能
   - 持续测试和验证

5. ✅ **充分测试**
   - 使用 Open XML SDK 的测试资产（901 个文件）
   - 与 MS Word 兼容性测试
   - 性能和内存测试

### 风险和缓解措施

| 风险         | 影响  | 概率  | 缓解措施                         |
| ------------ | ----- | ----- | -------------------------------- |
| 规范理解不足 | 🔴 高 | 🟡 中 | 深入学习 ISO 29500，参考 C# 实现 |
| 性能问题     | 🟡 中 | 🔴 高 | 流式处理，延迟加载，缓存优化     |
| 内存泄漏     | 🔴 高 | 🟡 中 | 资源管理模式，及时释放引用       |
| 兼容性问题   | 🔴 高 | 🟡 中 | 充分测试，使用标准规范           |
| 代码量过大   | 🟡 中 | 🔴 高 | 代码生成器，模块化设计           |
| 维护困难     | 🟡 中 | 🟡 中 | 良好的文档，清晰的架构           |

### 建议的开发策略

#### 第一步：2-3 周原型验证（必做）

```typescript
// 目标：验证核心概念的可行性
✅ ZIP 文件读写
✅ XML 解析和生成
✅ 基本文档结构读取
✅ 简单的文本提取
✅ 创建最简单的文档
```

#### 第二步：MVP 开发（6-8 周）

```typescript
// 目标：可用的基础版本
✅ 打开和保存 .docx
✅ 段落和文本操作
✅ 基本格式设置
✅ 单元测试覆盖
```

#### 第三步：功能扩展（8-10 周）

```typescript
// 目标：接近完整功能
✅ 样式系统
✅ 表格支持
✅ 图片处理
✅ 高级格式
```

#### 第四步：优化和发布（3-4 周）

```typescript
// 目标：生产就绪
✅ 性能优化
✅ 完整测试
✅ 文档编写
✅ 发布 NPM 包
```

### 最终建议

迁移 Open XML SDK 的 Word 功能到 JavaScript/TypeScript 是一个**大型但完全可行**的工程。关键点：

1. 🎯 **不要一次性实现所有功能** - 从 MVP 开始
2. 📚 **深入理解 OpenXML 规范** - 这是成功的基础
3. 🔧 **利用现有资源** - 使用 Open XML SDK 的数据定义和测试文件
4. ✅ **持续测试** - 每个功能都要验证兼容性
5. 📦 **考虑发布开源** - 社区可以帮助完善和维护

**预计全职开发时间：4-6 个月**

---

_文档版本：1.0_  
_最后更新：2025 年 10 月 20 日_
