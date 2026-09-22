# PPT 模板数据渲染器

## 1. 项目简介

这是一个独立的 Python MVP（最小可行原型），用于验证一个核心流程：**使用外部数据和图片自动填充固定版式的 PowerPoint 模板**。

```text
PPTX 模板 + JSON 数据 + 图片
          ↓
     Python Renderer
          ↓
       输出 PPTX
```

职责划分非常明确：

- **模板**（`.pptx`）负责版面设计与占位符类型定义；
- **JSON 数据**负责提供实际渲染值；
- **Renderer** 只负责把数据填入模板，不做任何视觉设计决策。

生成的输出是合法的 PowerPoint 文件，可正常打开。本项目是未来 “AI + PPTist” 演示文稿系统的前期验证，目前**未与** Vue、FastAPI、PPTist、数据库或任何 AI 服务集成。

## 2. 当前能力

目前已支持：

- **文本占位符**：`{{text:key}}` 替换为 JSON 中对应值
- **图片占位符**：`{{image:key}}` 对应的 Shape 作为图片区域
- **文本样式保持**：字体、字号、加粗、颜色、对齐、文本框位置与尺寸均不被修改
- **图片区域替换**：删除占位 Shape，在原位置、按原尺寸插入真实图片
- **图片 Cover Crop**：保持图片纵横比、居中裁剪填满区域，绝不溢出占位区域
- **JSON 数据驱动**：唯一的约定就是 `占位符 key → data[key]`
- **多页 PPT 模板**：模板包含多少页即可渲染多少页
- **内联文本占位符**：如 `Total reported sales: {{text:sales}} year to date.`
- **跨 Run 文本占位符**：PowerPoint 将占位符拆分到多个 Run 时仍能正确替换
- **模板校验**：非法占位符（如 `{{title}}`、`{{file:x}}`、`{{image}}`）会抛出 `MalformedPlaceholderError`
- **自动化验证**：`verify.py` 对工作流和占位符语义进行自动检查

MVP 目前**只支持两种占位符类型**：

```text
{{text:key}}
{{image:key}}
```

不支持 number、chart、table 等其他类型。

## 3. 占位符规范

```text
{{text:title}}            → data["title"]           渲染为文本
{{text:sales}}            → data["sales"]           渲染为文本
{{image:product_image}}   → data["product_image"]   作为图片源插入
```

**类型由模板定义，而不是由数据推断。** JSON 中不包含任何类型元数据：

```json
{
  "title": "2026年销售报告",
  "sales": "1258万元",
  "product_image": "assets/product.png"
}
```

- 模板中写 `{{text:sales}}`：值一律作为普通文本渲染，即使值看起来像文件路径。
- 模板中写 `{{image:product_image}}`：值才作为图片源使用。

> Renderer 不根据 `.png`、`.jpg` 等文件扩展名猜测数据类型。

其他规则：

- 只有当 Shape 的**全部文本**恰好是一个 `{{image:key}}` token 时，该 Shape 才成为图片区域。
- 图片相对路径先相对 JSON 文件所在目录解析，再相对当前工作目录解析。
- JSON 中缺失的 key 会保持占位符原样（便于调试和校验）。

## 4. 项目结构

```text
ppt-template-renderer/
├── templates/
│   └── sales_report_template.pptx
├── data/
│   └── example.json
├── assets/
│   └── product.png
├── output/
│   └── sales_report.pptx
├── src/
│   ├── __init__.py
│   ├── template_parser.py
│   ├── text_renderer.py
│   ├── image_renderer.py
│   └── renderer.py
├── build_test_template.py
├── main.py
├── verify.py
├── requirements.txt
└── README.md
```

各文件职责：

| 文件 | 职责 |
|---|---|
| `templates/sales_report_template.pptx` | 自动生成的 3 页测试模板（封面 / 指标卡片 / 产品展示） |
| `data/example.json` | 测试数据，仅纯 key-value，无类型信息 |
| `assets/product.png` | Pillow 自动生成的测试产品图 |
| `src/template_parser.py` | 扫描 Slide，识别 `{{text:}}` / `{{image:}}` 占位符，报告非法占位符 |
| `src/text_renderer.py` | Run 级文本替换，保留原有样式 |
| `src/image_renderer.py` | 图片区域替换 + Cover Crop 裁剪 |
| `src/renderer.py` | 通用渲染入口 `render(template_path, data, output_path)` |
| `build_test_template.py` | 生成测试模板与测试图片（无需手工做 PPT） |
| `main.py` | 命令行入口 |
| `verify.py` | 自动化验证（端到端工作流 + 占位符语义测试） |

## 5. 使用方法

### 安装依赖

```bash
pip install -r requirements.txt
```

安装 `python-pptx` 和 `Pillow`。

### 生成测试模板

```bash
python build_test_template.py
```

用 Python 自动生成 `templates/sales_report_template.pptx`（3 页、含新语法占位符）和 `assets/product.png`，无需手工制作 PPT。

### 渲染 PPT

```bash
python main.py --template templates/sales_report_template.pptx --data data/example.json --output output/sales_report.pptx
```

读取模板 + JSON 数据，执行图片替换和文本替换，输出到 `--output` 指定的路径，并打印实际替换的 text/image key 列表。

### 运行验证

```bash
python verify.py
```

自动执行两类检查：端到端工作流检查（文件存在、3 页齐全、值已渲染、无残留占位符、图片已插入、尺寸未变、zip 合法）和占位符语义检查（详见第 8 节）。全部通过时输出 `ALL CHECKS PASSED`，否则以非零退出码失败。

## 6. Renderer 接口

```python
render(template_path, data, output_path)
```

| 参数 | 说明 |
|---|---|
| `template_path` | 源 `.pptx` 模板路径（只读，永不被修改） |
| `data` | `dict`：key → 值。文本 key 对应字符串，图片 key 对应图片路径 |
| `output_path` | 生成的 `.pptx` 输出路径（父目录自动创建） |

返回值包含实际替换的 `text_keys` 和 `image_keys`。可选参数 `base_dirs` 用于控制相对图片路径的解析目录。

Renderer 是**完全通用**的：内部没有任何“销售报告”相关逻辑，不存在 `replace_sales()` 之类的业务函数。销售报告只是测试数据；核心机制始终是：

```text
{{text:variable_name}} / {{image:variable_name}}
        ↓
data["variable_name"]
        ↓
render
```

## 7. 当前实现原理

### 文本渲染

Renderer 遍历每页 Slide 的 Shape 与段落，把段落的多个 Run 拼接成完整文本，用正则定位 `{{text:key}}`，然后**只改写 Run 的文本内容**——不重建任何 XML 样式节点。因此字体、字号、加粗、颜色、对齐和文本框几何信息全部保留。占位符被 PowerPoint 拆分到多个 Run 时，从占位符起始的 Run 开始写入替换值（该值继承起始 Run 的样式），其余被覆盖的 Run 清空对应片段。

### 图片渲染

Renderer 将**整个文本内容恰好为** `{{image:key}}` 的 Shape 识别为图片区域：

1. 记录该 Shape 的原始位置与尺寸（EMU）；
2. 删除占位 Shape；
3. 在同一区域插入真实图片（尺寸与区域完全一致）；
4. 通过 OOXML 的 `srcRect` 裁剪属性实现 **Cover Crop**：当图片纵横比与区域不一致时，居中裁掉多余部分，保证图片铺满区域、保持纵横比且不溢出边界。

## 8. 当前验证结果

```text
python build_test_template.py          # OK
python main.py ...                     # OK
python verify.py                       # ALL CHECKS PASSED
```

`verify.py` 的两部分检查**全部通过**：

- **端到端工作流**：模板与输出存在、PPTX 可解析、3 页齐全、所有期望值已渲染、无残留 `{{...}}`、图片已插入且在原区域内、页面尺寸未变、zip 结构合法、图片已嵌入包内。
- **占位符语义**：`{{text:key}}` 渲染为文本；`{{image:key}}` 渲染为图片；`.png` 值经 `{{text:key}}` 仍按纯文本渲染且不生成图片；`{{title}}`、`{{file:x}}`、`{{image}}` 等非法占位符被拒绝；Cover Crop 数值验证通过（居中对称裁剪）；单 Run、内联、跨 Run 场景的文本样式保持均通过。

## 9. 当前限制

- 缺少数据 Key 时，目前占位符会保持原样，不会直接报错。
- 图片占位符必须独占整个 Shape（Shape 内只能有这一个 `{{image:key}}` token）。
- 跨 Run 的文本占位符采用起始 Run 的样式（不会按值重新排版样式）。
- 当前主要处理 Slide Shape。
- 不处理 Notes、Masters、Layouts 中的占位符。
- Group 中的文本可以扫描，但 Group 暂不作为图片区域处理。
- 相对图片路径仅按“JSON 目录 → 当前工作目录”顺序解析，不支持 URL。

## 10. 后续迁移方向

本项目目前是**独立的 Python 原型**，尚未与任何系统集成。未来可能的集成链路：

```text
Vue3
  ↓
FastAPI
  ↓
Template Renderer
  ↓
Presentation Schema
  ↓
PPTist
```

当前刻意保持通用的接口：

```python
render(template_path, data, output_path)
```

以及“占位符 key ↔ 数据 key”的纯映射机制，使得渲染核心后续可以整体嵌入更大的 AI + PPT 系统：上层（AI 或用户）只负责产出 JSON 数据，渲染核心无需改动。届时需要适配的部分包括：将模板 Shape 映射为 PPTist 的 JSON slide schema、图片源从本地路径扩展为 URL/对象存储、以及把缺失 key 的宽松处理改为严格校验。这些均属于未来工作，本阶段不做实现。

## 11. 设计原则

1. **模板负责结构**——版式、颜色、位置全部由 `.pptx` 模板决定。
2. **模板负责占位符类型**——`text` 还是 `image` 写在模板里。
3. **JSON 负责提供数据**——只有值，没有类型元数据。
4. **Renderer 负责执行渲染**——只做数据填充，不做设计决策。
5. **Renderer 不根据数据内容猜测类型**——扩展名不作为判断依据。
6. **第一版只支持 `text` 和 `image` 两种类型**。
7. **不在 Renderer 中加入业务领域逻辑**——销售报告只是测试数据。
