# PPT-P0-11 扫描场景模拟输入

状态：**工程回归输入已冻结；真实扫描件、人工审阅和 PowerPoint 执行仍未完成**。这不是正式专业验收材料包，不计入已就绪候选材料数量或 20 项通过数。

## 三份输入与来源

| 文件 | 用途 | 核验结果 |
| --- | --- | --- |
| `deardorff-2020-image-only.pdf` | 模拟扫描失败输入；将同一篇公开论文逐页渲染后写为仅含图片的 PDF | 11 页，2,390,574 字节；没有可提取的正文字符；**并非纸质文件的真实扫描件** |
| `deardorff-2020-article.pdf` | 可解析的原始文本 PDF | 11 页，原文件未改，382,769 字节 |
| `deardorff-2020-assistive-text.txt` | 用户补交的等价纯文本辅助材料 | 从原 PDF 用 `pdftotext -layout` 提取，63,819 字节；须以原 PDF 校对页码和事实 |

论文：Ariel Deardorff, “Assessing the impact of introductory programming workshops on the computational reproducibility of biomedical workflows,” *PLOS ONE* 15(7), e0230697 (2020)，DOI `10.1371/journal.pone.0230697`。[出版页](https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0230697)、[原始 PDF](https://journals.plos.org/plosone/article/file?id=10.1371/journal.pone.0230697&type=printable)、[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)。图像化 PDF 和辅助文本是 WisWork 基于获许可论文制作的衍生文件，保留作者、来源及修改说明；固定于 2026-09-28。每份文件的 SHA256 在 `SHA256SUMS`。

`generate-materials.py` 可重建衍生文件，需要 Poppler `pdftoppm`/`pdftotext` 和 Pillow；重新构建若摘要改变，须提升本包版本。图像化 PDF 的第一页已渲染并检查可读性。共享解析器回归测试要求它返回 `pdf_no_extractable_text`，原 PDF 与辅助文本则正常提取。此项仅证明文本提取边界，不等于 OCR 或真实扫描件处理能力。

## 工程回归固定任务提示

> 先上传图像化 PDF 与原始文本 PDF，分别报告提取覆盖情况，不能从图像化 PDF 猜测文字。若图像化 PDF 无法提取，请明确请求等价文本；我再上传 `deardorff-2020-assistive-text.txt`。继续完成 8 页中文科研摘要，逐条保留来源页码，对无法从原 PDF 核对的文字注明限制。保存、关闭、用真实 PowerPoint 重开并检查文本及图表可编辑，保留解析状态、补料步骤、截图、QA 和 PPTX 摘要。

正式 PPT-P0-11 执行仍须换用**获授权的真实扫描 PDF**，由人工标记扫描页和文本页，核验脱敏/使用权限，并在真实 PowerPoint 完成宿主证据。此工程材料不冒充已满足该输入要求。
