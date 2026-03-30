# Markdown 語法指南

專案支援以下 Markdown 語法：

## 1. 標題 (Headings)
支援 H1 到 H3：
```markdown
# 標題一 (H1)
## 標題二 (H2)
### 標題三 (H3)
```

## 2. 文字格式 (Text Formatting)
```markdown
**粗體文字**
*斜體文字*
`行內程式碼`
```

## 3. 清單 (Lists)
**無序清單 (Unordered List)**
支援使用 `-` 或 `*`：
```markdown
* 項目一
- 項目二
```

**有序清單 (Ordered List)**
```markdown
1. 第一項
2. 第二項
```

## 4. 區塊與引用 (Blocks & Quotes)
**引用區塊 (Blockquote)**
```markdown
> 這是一段引言
```

## 5. 連結與圖片 (Links & Images)
**加入連結 (Links)**
```markdown
[連結文字](https://example.com)
```
*(連結會自動在新分頁開啟)*

**加入圖片 (Images)**
```markdown
![圖片圖片](圖片網址或路徑)
```

## 6. 特殊自定義語法 (Custom Syntax)
除了標準 Markdown，專案還支援以下擴充語法：

**插入影片 (Video Embed)**
會自動將 YouTube 等連結轉換為響應式的 iframe 影片區塊。
```markdown
::video[https://www.youtube.com/embed/影片ID]
```

**可展開的 ESP 說明區塊 (Collapsible ESP Details Block)**
可用來製作點擊後可展開看詳細內容的區塊：
```markdown
:::esp[這裡寫摘要或標題文字]
這裡是展開後的詳細內容...
（可以放段落或清單等內容）
:::
```
