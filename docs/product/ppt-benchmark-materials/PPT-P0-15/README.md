# PPT-P0-15 候选材料包：现稿文字与布局修改

状态：**独立八页现稿、原始论文、图片授权清单及双路径任务口径已冻结；科研审阅、生成绑定稿的真实导入回执与 PowerPoint 执行未完成**。本包只计候选材料，不计专业任务通过。

## 输入、来源与权利

- `wiswork-image-dense-research-draft.pptx`：WisWork 自制的八页**任意现稿**，含原生文字、形状、14 次图片放置和两张原生图表；与 [P0-13](../PPT-P0-13/README.md) 原稿字节相同，不预设 WisWork 项目或导入映射。
- `deardorff-2020-article.pdf`：Deardorff, _PLOS ONE_ 15(7), e0230697 (2020)，[DOI](https://doi.org/10.1371/journal.pone.0230697)，CC BY 4.0。与 [P0-01](../PPT-P0-01/README.md) 原始论文 PDF 字节相同。研究内容尚待科研审阅。
- `images/schematic-01.png` 至 `schematic-12.png` 与 `asset-rights.json`：自制授权示意图，**不代表论文测量数据**。现稿已嵌入图片；独立文件用于权利核对和生成绑定稿的素材输入。
- `scenario.json`：任意现稿的两对象修改、禁止修改页、手工位移冲突和生成绑定稿目标。生成绑定稿必须在验收时由 WisWork 从同一论文生产并导入，实际 Project/Slide/Host ID 与回执不能由材料包预造。

`reference-plan.json` 与 `reference-deck.json` 将论文 PDF 原文、图片权利声明、八页施工图、目标页自制示意图及两张图表底层值接入 WisWork 生产输入。`wiswork-generated-candidate.pptx` 由同一原生页模型编译，含八页原生对象、授权示意图与两张图表；它仍只是生成路径的可检查候选，尚无 PowerPoint 导入映射。可运行 `node --import tsx generate-bound-candidate.cjs` 重建。

`SHA256SUMS` 冻结上述 17 份原输入及新增的计划、页模型，共 19 份文件；`verify-materials.mjs` 核对全部摘要、论文摘录、图片权利、两份八页原生稿、图表缓存及嵌入工作簿和第 4 页目标对象。复用原文与素材不复用其他任务的执行结果。

本机 `node tools/ppt-agent-electron-real-relay-smoke.mjs --benchmark=P0-15 --built-taskpane` 已用构建后的 Taskpane、真实 Electron PC 与 Rust Relay 完成论文 PDF、权利 JSON 上传、计划保存、八页生产、PPTX/PDF 回读与任务恢复；生成候选稿还通过 LibreOffice 八页渲染。它只验证生成路径的工程链路，不能制造另一份真实 PowerPoint 文档的导入映射，也不能替代任意现稿路径的宿主修改和保存重开证据。

## 固定任务提示

> 先打开本包八页任意现稿，读取第 4 页实际宿主对象 ID、标题原文、左图说明的原生坐标与截图；只把标题“参与者与证据路径”改为“参与者与研究证据”，并把左图说明下移 0.08 英寸，保持说明文字、字体及宽高不变。第 1、2、3、5、6、7、8 页不得修改，尤其核对第 5、6 页的原生图表。先显示写前保存点与差异，等待确认后执行，只复核受影响页。另做独立冲突尝试：在确认前由用户手动再移动说明 0.04 英寸，要求旧提案拒绝覆盖；重新读取基线后再修改。随后在另一份文稿中使用同一论文与授权示意图，经 WisWork 生成并导入八页同等内容，保留父成果及业务页/宿主页映射；只在该生成绑定稿上重复一页两对象修改与局部 QA。分别保存任意现稿与绑定稿的前后文件、对象清单、截图和回执；关闭并在 PowerPoint 重开，核对原生对象仍可编辑。两条路径的证据不能互相替代。

## 执行前与通过门槛

1. 科研审阅人核对现稿与论文的事实、图表数值和页码；如需修改材料，提升版本并重新计算全部摘要。
2. 在真实 PowerPoint 中建立两份不同文档身份。生成绑定稿须保存 WisWork 计划、SlideIR、编译产物、导入回执和页映射；本包没有伪造已存在的父成果。
3. 按 [P0 验收清单](../../wiswork-ppt-agent-p0-acceptance-2026-09-23.md)记录首试/补跑、陈旧提案拒绝、局部 QA 状态、八页对象和截图、文件 SHA256、保存关闭重开及人工审阅。当前均未执行。
