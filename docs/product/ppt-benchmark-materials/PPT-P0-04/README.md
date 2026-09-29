# PPT-P0-04：Rule 702 历史规则差异候选材料

状态：**候选，待审**。固定历史 as-of **2024-12-01**；本材料不证明今天仍适用，不构成个案法律意见。8 页是拟制作演示文稿的范围，并非已制作或验收的 PPT。

## 官方原件

| ID         | 冻结原件                                                       | 角色及准确位置                                                                                                                                    |
| ---------- | -------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| old        | [2022 规则汇编](originals/2022-federal-rules-evidence.pdf)     | 官方汇编版本日 2022-12-01（封面印刷年 2023）；PDF 29 / 印刷 15：修订前 Rule 702，仅作比较                                                         |
| historical | [2024 规则汇编](originals/2024-federal-rules-evidence.pdf)     | 版本日 2024-12-01；PDF 30–31 / 印刷 15–16：历史时点 Rule 702；PDF 43–44 / 印刷 28–29：Rule 1101 范围和例外                                        |
| amendment  | [2023 官方修订包](originals/2023-courts-amendment-package.pdf) | 全包 224 PDF 页；PDF 194：2023-04-24 最高法院命令；PDF 198 / 证据规则部分印刷 4：清洁 Rule 702；PDF 209–213 / 另起印刷 1–5：黑线及 Committee Note |

来源 URL、真实字节大小、SHA-256、取得时间和页级提取散列见 [materials-manifest.json](materials-manifest.json)。三个 PDF 原件保存在本目录，缺任一原件即不能作为完整候选。规则、命令和解释保持各自身份；Committee Note 是官方解释，不是法规新增条款或个案裁判。

## 法域、效力和实施边界

声明域为 `law`，法域是美国联邦法院。联邦证据规则为依法律授权制定的法院证据规则；不要把它描述为所有美国州法院或所有行政程序当然适用的实体法。Rule 1101（historical PDF 43–44）列明法院、案件类别及不适用情形。2023-04-24 命令（amendment PDF 194）规定 2023-12-01 生效：其后的新程序，以及在公正且可行范围内当时待决的程序。该后半限定不可省略。

2024 历史汇编 Rule 702 的修改注记仍列 2023 修订；不能因此宣称整个 2024 汇编没有其他规则修订。旧版无显式新增短语也不意味着旧法不要求可靠性门槛或由陪审团决定全部问题；官方 Note 将本次变化表述为澄清和强调。

## 核对与输入边界

```sh
node docs/product/ppt-benchmark-materials/PPT-P0-04/verify-materials.mjs
sha256sum docs/product/ppt-benchmark-materials/PPT-P0-04/originals/*.pdf
pdftotext -f 29 -l 29 -layout docs/product/ppt-benchmark-materials/PPT-P0-04/originals/2022-federal-rules-evidence.pdf -
pdftotext -f 30 -l 31 -layout docs/product/ppt-benchmark-materials/PPT-P0-04/originals/2024-federal-rules-evidence.pdf -
pdftotext -f 209 -l 213 -layout docs/product/ppt-benchmark-materials/PPT-P0-04/originals/2023-courts-amendment-package.pdf -
```

`extracts/` 是 `pdftotext -layout` 的指定原页输出，不是新法规、替代原件或签章认证。`basis.json` 的绝对 UTF-16 offset 只对应它引用的独立 extract 文本文件，window 从 0 开始；不能直接复用为产品 PDF 解析器的 offset。生产使用必须上传实际原件、读取实际解析窗口后重新取得 literal basis。跨行连字符和黑线删除/新增格式不能凭提取串推断；amendment PDF 209 的删除线和下划线已渲染核看，清洁规则用 PDF 198 和 historical PDF 30–31 校对。文件摘要不证明法律适用或官方数字签章。

## 许可边界

保留的是美国政府发布的规则、最高法院命令及其官方解释原件；按 [17 USC 105(a) 的官方文本](https://usc-cdn.house.gov/view.xhtml?edition=prelim&num=0&req=granuleid%3AUSC-prelim-title17-section105)记录美国政府作品的美国版权边界。政府持有的转让版权、第三方引文/材料、徽章标识及其他司法辖区权利不能据此一概豁免。只使用本案必要的规则和 Note 段落，不导出与任务无关的修订包内容；官方来源不等于所有附件任意再许可。专业及版权范围复审仍待完成。

## 未完成的验收

没有法律专业审阅、实际 PowerPoint 制作、布局验收、宿主 QA、保存重开核对或用户案件适用判断；不标任务 passed 或材料就绪。原件位置/摘要/字面提取的本机核对只是材料完整性检查。八页范围与比较归因分别见 [eight-page-scope.md](eight-page-scope.md) 和 [comparison.md](comparison.md)。

取得备注：默认代理导致 TLS 错误，直连官方 HTTPS 成功。USCourts 标题为“Federal Rules of Evidence 2023”的下载链接实际返回 2025 汇编，已排除；未将页面标题当作历史版本证据。下载使用 `HTTPS_PROXY= HTTP_PROXY= ALL_PROXY= curl -fL <manifest URL> -o <manifest path>`，再次下载须先比较固定摘要，变更不能自动覆盖冻结原件。
