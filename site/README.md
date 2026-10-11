# DocKiln 文档站

使用 Docusaurus Classic 默认主题，与 Docusaurus 官网保持相同的导航、侧边栏、
代码块和明暗主题风格。站点地址：<https://shenyankm.github.io/DocKiln/>。

## 本地开发

需要 Node.js 20 或更新版本；GitHub Actions 使用 Node.js 24。

```bash
cd site
npm ci
npm start
```

生产构建与预览：

```bash
npm run build
npm run serve
```

首页位于 `src/pages/index.js`；文档直接读取仓库的 `docs/*.md`，导航顺序配置于
`sidebars.js`。不要复制另一份指南。Markdown 中的本地 `benchmarks/` 和 `probes/`
链接在网页中显示为未公开复验材料，不上传其文件。

修改 `site/`、`docs/*.md` 或 Pages 工作流并推送到 `dev` 后自动构建和部署。
PR 只验证构建。断链会导致构建失败；仅 `site/build/` 会上传到 GitHub Pages。
