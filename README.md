# 喵喵工具集 · 网站导航（miaonav）

> 🌐 语言切换：[English](./README_EN.md)

一个由浏览器书签生成的静态「精选工具网站导航页」，支持分类平铺、标签联动筛选、暗色模式、中英双语，并附带一个桌面可视化编辑工具。

- 🌐 无英语分页版本：**https://github.com/cheng01315/miaonav/tree/v1.01_miaonav**
- 🌐 在线演示：**https://www.meowtool.com/miaonav**
- 🍴 派生自：[Pintree `pintree-old-pages` 分支](https://github.com/Pintree-io/pintree/tree/pintree-old-pages)
- ✍️ 作者：Cheng

---

## 一、功能特性

**导航页面（`index.html` / `en.html`）**

- 中英双语：两个同源页面，右上角一键切换；含 canonical / hreflang 声明
- 首页 9 大分类平铺：emoji 图标 + 站点数量统计，侧边栏滚动联动高亮（scroll-spy）
- 二级 / 三级标签联动筛选，每区块预览 18 个站点，超出可进「查看更多」详情页（面包屑 + 递归分组）
- 搜索同时匹配网站标题与描述；移动端提供吸顶搜索栏与双列卡片布局
- 暗色模式，偏好持久化到 `localStorage`，跨语言跳转不闪烁
- 站点图标全部本地化（`assets/logo/`），加载失败回退默认占位图

**桌面编辑工具（`Website navigation tool/`）**

- 导入 / 导出 Excel（`.xlsx`）与 JSON，表格化增删改查与排序
- 一键导出符合 Pintree 结构的 `pintree.json`，多源 favicon 自动下载并统一转 PNG
- 分类 emoji 可视化设置（右键「设置 emoji…」，导出时写入 JSON）
- 内置百度 / 腾讯翻译，增量翻译生成 `pintree.en.json`（边翻边落盘、中断可续翻）
<<<<<<< HEAD
- 提供中文界面版（`Website navigation tool.py`）与英文界面版（`Website navigation tool EN.py`），两者功能完全一致
- 完善的 SEO 标签（canonical / Open Graph / Twitter Card）

---

## 二、项目结构

```
miaonav/
├── index.html                     # 中文导航页
├── en.html                        # 英文导航页（与 index.html 同源，手工保持同步）
├── css/
│   ├── styles.css                 # 自定义样式（.mn-* 系列 + 移动端适配）
│   └── tailwind.css               # Tailwind 构建产物
├── json/
│   ├── pintree.json               # 中文导航数据（9 个顶级分类，带 emoji）
│   ├── pintree.en.json            # 英文导航数据（翻译工具生成）
│   └── _translation_config.json   # 翻译 API 凭证（见下文说明，请勿公开）
├── assets/
│   ├── logo.svg                   # 站点 Logo
│   ├── og.webp                    # 社交分享图
│   ├── favicon/                   # 站点 favicon
│   ├── default-icon.svg           # 站点图标加载失败时的占位图
│   └── logo/                      # 各网站图标（本地 PNG）
└── Website navigation tool/
    ├── Website navigation tool.py     # 桌面可视化编辑工具（Tkinter，中文界面）
    ├── Website navigation tool EN.py  # 桌面可视化编辑工具（Tkinter，英文界面）
    ├── translation.py                 # 百度 / 腾讯翻译接口封装
    └── _translation_cache.json        # 增量翻译缓存（运行时自动生成）
```

---

## 三、快速开始

### 1. 本地预览

页面通过 `fetch` 读取 JSON，请用本地 HTTP 服务打开（不要直接双击 `file://`）：

```bash
python -m http.server 8000
# 中文页 http://localhost:8000/index.html
# 英文页 http://localhost:8000/en.html
```

### 2. 使用桌面编辑工具（推荐）

```bash
pip install openpyxl requests pillow
python "Website navigation tool/Website navigation tool.py"
# 需要英文界面时改用：
python "Website navigation tool/Website navigation tool EN.py"
```

1. 导入现有 `json/pintree.json` 或 Excel，进行编辑、排序、设置 emoji
2. 点击「导出 json」生成 `json/pintree.json`，刷新网页即生效
3. 翻译区共 4 个按钮：
   - **翻译设置…**：选择服务商（百度 / 腾讯）、填写 APPID 与密钥、QPS 限速、测试连接
   - **立刻翻译**：以中文 JSON 为基准，仅翻译新增 / 缺失条目（增量）
   - **重译旧内容**：清空缓存，全量重新翻译并覆盖英文 JSON
   - **导出英文 json**：基于当前内存数据翻译，仅输出 `pintree.en.json`

> 🔑 翻译凭证保存在 `json/_translation_config.json`，请替换为你自己申请的百度 / 腾讯密钥。该文件含敏感信息，公开仓库建议将其移出 Git 跟踪。

### 3. 手动编辑 JSON

数据为 Pintree 书签结构，示例：

```json
{
  "type": "folder",
  "title": "搜索工具",
  "emoji": "🔍",
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

### 4. 部署

纯静态站点，将整个目录上传到任意静态托管即可。所有本地资源（CSS / JS / 图片 / JSON / 页面互链）均使用相对路径，部署在根目录或任意子路径下都能正常工作；第三方统计脚本与 SEO 规范链接为绝对 URL，属正常需要。

---

## 四、更新日志

### V2.1.0（2026-09-24）

- 新增英文界面版桌面编辑工具 `Website navigation tool EN.py`：与中文版（`Website navigation tool.py`）功能完全一致，UI 全英文，方便非中文用户使用

### V2.0.0（2026-09-23）

- 翻译能力升级：接入百度 / 腾讯翻译，4 个翻译按钮；增量缓存分批原子落盘，中断 / 停止后可续翻且不产生脏数据；翻译进度窗带进度条
- 新增英文页面 `en.html` 与语言切换胶囊按钮，运行时文案全面 I18N
- 分类图标改为读取 JSON 中的 `emoji` 字段（9 个顶级分类已配置）
- 暗色模式持久化，跨语言整页跳转不再闪回浅色
- 搜索支持匹配网站描述（description）
- 站点图标与页脚图片全部改为本地相对引用，移除硬编码的 `<base href="/miaonav/">`

### V1.0.1（2026-09-03）

- 新增移动端吸顶搜索栏，两端搜索内容同步
- 手机端卡片改为每行 2 列，并收紧图标与间距
- 补全 SEO：侧边栏品牌名设为全页唯一 H1

### V1.0.0（2026-08-27）

首个 fork 版本：桌面可视化编辑工具、首页分类平铺与标签筛选、scroll-spy 侧边栏、图标本地化、description 字段、品牌视觉与统计埋点。

---

## 五、致谢

页面骨架与数据格式派生自 **Pintree**（[pintree-old-pages](https://github.com/Pintree-io/pintree/tree/pintree-old-pages)），感谢原作者。

---

## 六、许可证

MIT License — 基于 Pintree 修改分发，请保留原作者与项目出处。
