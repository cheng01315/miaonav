# 喵喵工具集 · 网站导航（miaonav）

> 🌐 语言切换：[English](./README_EN.md)

> 一个由浏览器书签一键生成的「精选工具网站导航页」，支持分类平铺、标签联动筛选、暗色模式与本地可视化编辑。

- 🌐 在线演示：**https://www.meowtool.com/miaonav**
- 🍴 项目派生自：**https://github.com/Pintree-io/pintree/tree/pintree-old-pages**（Pintree 旧版页面分支）
- ✍️ 作者：Cheng

---

## 一、项目简介

本项目基于开源项目 **Pintree** 的 `pintree-old-pages` 分支二次开发而来。原版 Pintree 的定位是「把浏览器书签变成导航网站」，而本 fork 在其基础上做了一套**面向中文用户的产品化改造**：

- 重新设计了首页交互（分类区块平铺 + 二级/三级标签联动筛选 + 滚动高亮）；
- 新增了一个**桌面端可视化编辑工具**，无需安装浏览器扩展、无需手写 JSON，即可批量管理导航数据；
- 完善品牌视觉、SEO 与统计埋点，可直接对外发布上线。

---

## 二、与原版 Pintree 的区别（新增功能）

### V1.0.1

发布日期：2026.9.3

**1.移动端搜索框（新增）**

- 在顶栏下方新增**吸顶搜索栏**（`lg:hidden`，仅小屏显示），带搜索图标 + 「清空」按钮，回车即搜。
- 逻辑复用桌面那套 `searchBookmarks` / `clearSearchResults`，**两端输入框内容同步**，清空时一起清。
- 细节调整：放大镜先是加了 `my-auto` 修复垂直居中，随后按你的要求**直接移除图标**，左内边距由 `pl-9` 收到 `pl-4`，更干净。
- 顺带覆盖了一个历史空白：640–1024px 原本没有任何搜索框，现在由这条移动搜索栏接管。

**2.移动端卡片布局（新增适配）**

- `@media (max-width:640px)` 里把网格从 `auto-fill minmax(210px,1fr)` 改成固定 **`repeat(2, 1fr)`**，即手机每行 2 张卡片。
- 配套收紧：卡片内边距、图标（40→32px）、描述缩成 1 行；分类跳转的 `scroll-margin-top` 调到 8.5rem，免得被吸顶搜索框挡住标题。

**3.SEO（基础补全）**

- 首页原本**完全没有 H1**，已把桌面侧边栏品牌名「喵喵工具集」升级为**全页唯一 `<h1>`**。
- 移动顶栏 / 抽屉菜单里的同名文字保持 `<a>`，避免重复 H1 稀释权重。

**涉及文件**

- `index.html`：移动端搜索框 DOM、H1 标题

- `css/styles.css`：移动端 2 列网格 + 卡片适配

  

### V1.0.0 

发布日期：2026.8.27

| 维度 | 原版 Pintree（pintree-old-pages） | 本 fork（miaonav） |
| --- | --- | --- |
| 数据编辑方式 | 必须安装 Chrome 扩展「Pintree Bookmarks Exporter」导出书签 JSON，再手动替换 `json/pintree.json` | 提供**桌面 GUI 编辑工具**（`Website navigation tool/Website navigation tool.py`），可导入 Excel / JSON、可视化增删改、排序、一键导出 |
| 图标获取 | 依赖远程 favicon 链接 | 内置**多源 favicon 下载服务**（Google / faviconkit / Yandex / favicon.im 兜底），自动下载到本地 `assets/logo/` 并统一转 PNG |
| 首页布局 | 文件夹卡片网格 / 书签列表 | **分类区块平铺**：每个分类带 emoji 图标与数量统计 |
| 筛选交互 | 基本无 | 每个分类支持**二级标签 + 可展开的三级联动标签**，前 18 条预览 + 「查看更多」详情页 |
| 导航体验 | 普通侧边栏 | 侧边栏 **scroll-spy 滚动联动高亮**当前分类；详情页带面包屑与递归分组 |
| 数据字段 | 标题 / 链接 / 图标 | 额外支持 **description 描述字段**（卡片显示两行简介） |
| 品牌与视觉 | 通用 Pintree 皮肤 | 全新品牌「喵喵工具集 / meowtool」，自定义 logo、favicon、OG 分享图，全套 `.mn-*` 自定义样式 |
| 统计与分析 | 无 | 集成 Umami、Google Analytics、Microsoft Clarity |
| SEO / 合规 | 基础 meta | 完善 canonical / Open Graph / Twitter Card |

#### 主要新增能力详解

1. **桌面可视化编辑工具（`Website navigation tool/Website navigation tool.py`）**
   
   - 导入 / 导出 Excel（`.xlsx`）与 JSON；
   - 表格化增、删、改网站条目；
   - 搜索与筛选；
   - 分类排序、网址排序（上移 / 下移 / 置顶 / 置底）；
   - 一键导出符合 Pintree 结构的 `pintree.json`；
   - 自动探测项目根目录下的 `assets/logo` 并管理图标。
   - 依赖：`openpyxl`（Excel）、`requests`（图标下载）、`Pillow`（图标转 PNG）。运行：
     ```bash
     pip install openpyxl requests pillow
     python "Website navigation tool/Website navigation tool.py"
     ```
   
2. **首页分类平铺 + 标签联动筛选**
   - 首页按顶层分类分区块展示，每个区块头部显示 emoji 图标与网站总数；
   - 区块内提供二级标签（子分类）行，点击可展开第三级标签，形成二级 / 三级联动；
   - 每个区块默认预览前 18 个网站，超过则显示「查看更多」进入完整详情页（面包屑 + 递归分组展示）。

3. **滚动联动高亮（scroll-spy）**
   - 使用 `IntersectionObserver` 在首页滚动时自动高亮侧边栏对应的分类项，提升长列表的浏览定位体验。

4. **图标本地化与统一**
   - 编辑工具自动下载网站 favicon 并落地到 `assets/logo/`，统一转为 PNG，避免依赖第三方实时服务、提升加载速度与稳定性。

---

## 三、项目结构

```
miaonav/
├── index.html                 # 导航站主页面（首页平铺 + 标签筛选 + 详情页）
├── css/
│   ├── styles.css            # 自定义样式（含 .mn-* 系列新增样式）
│   └── tailwind.css          # Tailwind 构建产物
├── json/
│   └── pintree.json          # 导航数据（分类 / 链接 / 图标 / 描述）
├── assets/
│   ├── logo.svg              # 站点 logo
│   ├── og.webp               # 社交分享图
│   ├── favicon/              # 站点 favicon 资源
│   ├── logo/                 # 各网站图标（编辑工具自动下载）
│   └── default-icon.svg     # 图标缺失时的默认占位
└── Website navigation tool/
    └── Website navigation tool.py   # 桌面端可视化编辑工具（新增）
```

---

## 四、使用说明

### 方式一：使用桌面编辑工具（推荐）
1. 安装依赖：`pip install openpyxl requests pillow`
2. 运行 `Website navigation tool/Website navigation tool.py`
3. 导入现有 `json/pintree.json` 或 Excel，进行编辑、排序
4. 点击「导出 json」生成 `json/pintree.json`，刷新网页即可生效

### 方式二：手动编辑 JSON（兼容原版）
直接按 Pintree 结构编辑 `json/pintree.json`：
```json
{
  "type": "folder",
  "title": "搜索工具",
  "children": [
    {
      "type": "link",
      "title": "Felo",
      "icon": "assets/logo/felo.ai.png",
      "url": "https://felo.ai/search",
      "description": "基于人工智能技术的智能搜索平台"
    }
  ]
}
```

### 本地预览
由于浏览器安全限制，请通过本地 HTTP 服务打开（不要直接双击 `index.html`）：
```bash
python -m http.server 8000
# 浏览器访问 http://localhost:8000
```

### 部署
将整个目录（含 `index.html`、`css/`、`json/`、`assets/`）托管到任意静态空间即可。当前线上版本部署于 **https://www.meowtool.com/miaonav**。

---

## 五、致谢

- 本项目的页面骨架与数据格式派生自 **Pintree**（[pintree-old-pages](https://github.com/Pintree-io/pintree/tree/pintree-old-pages)），在此表示感谢。
- 原项目采用 **MIT License**，本 fork 沿用相同开源协议。

---

## 六、许可证

MIT License — 基于 Pintree 修改分发，请保留原作者与项目出处。
