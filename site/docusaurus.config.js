const repository = 'https://github.com/shenyankm/DocKiln';

export default {
  title: 'DocKiln',
  tagline: '让文档进入你的 Python 工作流',
  url: 'https://shenyankm.github.io',
  baseUrl: '/DocKiln/',
  organizationName: 'shenyankm',
  projectName: 'DocKiln',
  trailingSlash: true,
  onBrokenLinks: 'throw',
  i18n: {defaultLocale: 'zh-CN', locales: ['zh-CN']},
  markdown: {
    format: 'md',
    hooks: {onBrokenMarkdownLinks: 'throw'},
    preprocessor: ({fileContent}) => fileContent
      // Local evidence is deliberately excluded from the public website.
      .replace(/\[([^\]]+)\]\((?:benchmarks|probes)\/[^)]+\)/g, '$1（本地复验材料，未公开）')
      .replace(/\]\(\.\.\/(README[^)]*|THIRD_PARTY_NOTICES\.md|scripts\/[^)]+)\)/g, `](${repository}/blob/dev/$1)`),
  },
  presets: [['classic', {
    docs: {
      path: '../docs',
      include: ['*.md'],
      sidebarPath: './sidebars.js',
      editUrl: ({docPath}) => `${repository}/edit/dev/docs/${docPath}`,
    },
    blog: false,
    theme: {customCss: './src/css/custom.css'},
  }]],
  themeConfig: {
    colorMode: {respectPrefersColorScheme: true},
    docs: {sidebar: {hideable: true}},
    navbar: {
      title: 'DocKiln',
      items: [
        {type: 'docSidebar', sidebarId: 'guides', position: 'left', label: '文档'},
        {href: `${repository}/blob/dev/README.zh-CN.md`, label: '快速开始', position: 'left'},
        {href: repository, label: 'GitHub', position: 'right'},
      ],
    },
    footer: {
      style: 'dark',
      links: [
        {title: '文档', items: [
          {label: 'DOCX 编辑', to: '/docs/docx-dom'},
          {label: '转换与安全', to: '/docs/enhanced-conversion'},
        ]},
        {title: '项目', items: [
          {label: 'GitHub', href: repository},
          {label: 'Issues', href: `${repository}/issues`},
          {label: 'CI 状态', href: `${repository}/actions`},
        ]},
        {title: '来源与许可', items: [
          {label: '上游 Aspose.Words FOSS', href: 'https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python'},
          {label: 'MIT 代码许可证', href: `${repository}/blob/dev/LICENSE`},
          {label: '第三方来源与许可', href: `${repository}/blob/dev/THIRD_PARTY_NOTICES.md`},
          {label: 'SIL OFL-1.1 字体许可证', href: `${repository}/blob/dev/src/dockiln/writers/pdf/fonts/OFL.txt`},
        ]},
      ],
      copyright: '© 2026 Aspose.Words FOSS Family & shenyankm。独立维护，非 Aspose 官方发布版，不保证 API 或渲染兼容。使用 Docusaurus 构建。',
    },
    prism: {additionalLanguages: ['python', 'bash']},
  },
};
