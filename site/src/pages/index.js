import React from 'react';
import Layout from '@theme/Layout';
import Link from '@docusaurus/Link';
import CodeBlock from '@theme/CodeBlock';

const features = [
  ['转换与生成', '读取 DOCX、旧版 DOC 与 Flat OPC，写出 DOCX、Flat OPC 和 PDF。'],
  ['中文 PDF', '内置中文字体，可选备用字体与文字塑形，也可使用独立的 LibreOffice PDF 后端。'],
  ['提取与编辑', '提取有序内容与格式，或使用 DocxDocument 编辑受支持的 DOCX 内容并保留未修改包部件。'],
];

export default function Home() {
  return (
    <Layout title="Python 文档处理" description="DocKiln：独立维护的 Python 文档工具库，用于转换、中文 PDF、结构化提取与 DOCX 编辑。">
      <header className="hero hero--primary">
        <div className="container text--center">
          <h1 className="hero__title">DocKiln</h1>
          <p className="hero__subtitle">让文档进入你的 Python 工作流</p>
          <p>转换、提取、编辑 · Python 3.10–3.14 · 无需 Microsoft Word</p>
          <Link className="button button--secondary button--lg" to="/docs/enhanced-conversion">阅读文档 →</Link>
        </div>
      </header>
      <main className="container margin-top--xl margin-bottom--xl">
        <section className="row margin-bottom--xl" aria-label="核心能力">
          {features.map(([title, text]) => (
            <article className="col col--4" key={title}><h2>{title}</h2><p>{text}</p></article>
          ))}
        </section>
        <section className="row margin-bottom--xl">
          <div className="col col--5">
            <h2>从一次转换开始</h2>
            <p>在干净的虚拟环境中安装。发行包名为 <code>dockiln</code>，导入名为 <code>dockiln</code>。</p>
            <p>需要可复现部署时，将 <code>@dev</code> 替换为已验证的完整提交哈希。旧代码需将导入入口改为 <code>dockiln</code>。</p>
            <Link to="https://github.com/shenyankm/DocKiln/blob/dev/README.zh-CN.md#安装">完整安装与使用说明 ↗</Link>
          </div>
          <div className="col col--7">
            <CodeBlock language="bash">{'python -m pip install "git+https://github.com/shenyankm/DocKiln.git@dev"'}</CodeBlock>
            <CodeBlock language="python">{`import dockiln as dk

doc = dk.Document("report.docx")
doc.save("report.pdf", dk.SaveFormat.PDF)
doc.save("copy.docx", dk.SaveFormat.DOCX)
data = doc.to_dict()`}</CodeBlock>
          </div>
        </section>
        <section className="row">
          <div className="col col--6">
            <h2>格式支持</h2>
            <table><thead><tr><th>格式</th><th>内置读取</th><th>内置写出</th></tr></thead><tbody>
              <tr><td>DOCX</td><td>支持</td><td>支持</td></tr>
              <tr><td>DOC / Word 97–2003</td><td>支持</td><td>—</td></tr>
              <tr><td>Flat OPC</td><td>支持</td><td>支持</td></tr>
              <tr><td>PDF</td><td>—</td><td>支持</td></tr>
            </tbody></table>
            <p>Flat OPC XML 文件需显式指定加载/保存格式。内置 RTF 仅限 OLE2/DOC-backed 文件；标准文本 RTF 可用 LibreOffice 后端转 PDF。</p>
          </div>
          <div className="col col--6">
            <h2>先验证，再交付</h2>
            <p>转换会重建受支持内容，不是 DOCX 无损往返；包级编辑使用 <code>DocxDocument</code>。</p>
            <p>内置引擎不提供完整 Word 排版、PDF 读取、Markdown/TXT 读写或 OCR。复杂字段、修订、浮动内容、脚注及分节页眉页脚仍有保真限制。</p>
            <p>用代表性文档验证输出。不可信文档需隔离处理，资源限制不是安全沙箱。</p>
            <Link to="/docs/docx-dom">了解 DOM 支持范围 →</Link>
          </div>
        </section>
      </main>
    </Layout>
  );
}
