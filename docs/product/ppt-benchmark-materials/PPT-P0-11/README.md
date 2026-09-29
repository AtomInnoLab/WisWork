# PPT-P0-11 扫描与部分解析材料

状态：**工程回归输入及一份真实扫描件已冻结；等价纯文本的人工校对和 PowerPoint 执行仍未完成**。这不是完整的正式专业验收材料包，不计入已就绪候选材料数量或 20 项通过数。

## 新增真实扫描件（2026-09-30）

NASA NTRS 的[NACA 研究备忘录 L50B01](https://ntrs.nasa.gov/citations/19930086231)是 1950 年纸质报告的实际扫描，记录页标明 *Work of the US Gov. Public Use Permitted*。本包 `naca-rm-l50b01-1950-real-scan.pdf` 保存其 30 页扫描图像和原有 OCR 层；`naca-rm-l50b01-scan-audit.json` 记录原始[下载地址](https://ntrs.nasa.gov/api/citations/19930086231/downloads/19930086231.pdf)、原始字节摘要、436 字节下载封装头、去除封装后 PDF 的摘要，以及逐页可提取字符数。只移除 PDF 标头前的封装字节，没有重新绘制页面。第 2 页目视确认为扫描空白页且无提取文字；部分图表页虽有 OCR 字符，但数量少且明显存在误识别。

`naca-rm-l50b01-ocr-draft.txt` 是从该扫描件提取的**未经人工校对 OCR 草稿**，不能当作等价辅助文本，也不能据此生成未经核对的结论。原来的 Deardorff 文本 PDF 与辅助文本属于另一篇论文，不能冒充这份 NACA 报告的对应文本。使用真实扫描件做正式 P0-11 前，仍须人工标注扫描/文本页、校对与任务范围相同的辅助文本及来源事实，并完成真实 PowerPoint 宿主验收。`sha256sum -c SHA256SUMS` 可验证六份输入/审计文件；`pdfinfo naca-rm-l50b01-1950-real-scan.pdf` 应报告 30 页。

共享 PDF 解析器已用该真实扫描件回归：识别 30 个源页定位，第 2 页文本段保持零长度，第 3 页标题可定位。该结果只证明页码和提取覆盖如实传递，不证明 OCR 内容准确。

跨组件回归又以同一 12.9 MB 扫描 PDF 经 Office 附件客户端分块上传至真实 PC 服务，重建 PC 服务后由 Agent 附件列表读取，保留 `sectionCount: 30` 和 `pagesWithoutExtractedText: [2]`；读取正文仍返回第 2 页零长度及第 3 页标题的来源定位。Taskpane 根据该字段提示该页缺文字并请求补料；这是本地模拟客户端与真实服务的工程证据，不是 PowerPoint 宿主验收。

## 三份原工程输入与来源

| 文件 | 用途 | 核验结果 |
| --- | --- | --- |
| `deardorff-2020-image-only.pdf` | 模拟扫描失败输入；将同一篇公开论文逐页渲染后写为仅含图片的 PDF | 11 页，2,390,574 字节；没有可提取的正文字符；**并非纸质文件的真实扫描件** |
| `deardorff-2020-article.pdf` | 可解析的原始文本 PDF | 11 页，原文件未改，382,769 字节 |
| `deardorff-2020-assistive-text.txt` | 用户补交的等价纯文本辅助材料 | 从原 PDF 用 `pdftotext -layout` 提取，63,819 字节；须以原 PDF 校对页码和事实 |

论文：Ariel Deardorff, “Assessing the impact of introductory programming workshops on the computational reproducibility of biomedical workflows,” *PLOS ONE* 15(7), e0230697 (2020)，DOI `10.1371/journal.pone.0230697`。[出版页](https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0230697)、[原始 PDF](https://journals.plos.org/plosone/article/file?id=10.1371/journal.pone.0230697&type=printable)、[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)。图像化 PDF 和辅助文本是 WisWork 基于获许可论文制作的衍生文件，保留作者、来源及修改说明；固定于 2026-09-28。每份文件的 SHA256 在 `SHA256SUMS`。

`generate-materials.py` 可重建衍生文件，需要 Poppler `pdftoppm`/`pdftotext` 和 Pillow；重新构建若摘要改变，须提升本包版本。图像化 PDF 的第一页已渲染并检查可读性。共享解析器回归测试要求它返回 `pdf_no_extractable_text`，PC 附件服务将其标为解析失败；原 PDF 与辅助文本则正常提取。另用两页工程 fixture 检查“第一页有文字、第二页无文字”时，PC 持久记录 `pagesWithoutExtractedText: [2]`，插件清单和界面提示补料。该 fixture 不是本包三份任务输入。此项仅证明文本提取边界，不等于 OCR 或真实扫描件处理能力。

## 工程回归固定任务提示

> 先上传图像化 PDF 与原始文本 PDF，分别报告提取覆盖情况，不能从图像化 PDF 猜测文字。若图像化 PDF 无法提取，请明确请求等价文本；我再上传 `deardorff-2020-assistive-text.txt`。继续完成 8 页中文科研摘要，逐条保留来源页码，对无法从原 PDF 核对的文字注明限制。保存、关闭、用真实 PowerPoint 重开并检查文本及图表可编辑，保留解析状态、补料步骤、截图、QA 和 PPTX 摘要。

正式 PPT-P0-11 可使用上述 NASA 真实扫描 PDF；仍须完成对应辅助文本和人工页码/事实标注，在真实 PowerPoint 完成宿主证据。原来的图像化论文仅是工程模拟，不能代替真实扫描。
