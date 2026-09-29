# 本机项目治理公开接线

依据原方案 §13.1，接续已审查的资源归属清理、内部删除事务及真实入口写保护。检查点399ac2ee的Shell108文件924通过、Store26文件267通过，10个核心文件冻结摘要一致。实机验收暂缓，工程验证只用合成项目。

## 单元一：本机严格适配器

新增 presentation-project-governance.ts 及测试。七操作使用32KiB严格平面请求；实际归属只在明确初始化时授权，读取独立控制不创建内容；策略数字CAS，删除使用原intent和真实drain，删除请求不加入自己等待的Work。生命周期内层2MiB，最大包装外层加14字节；其它业务既有预算不改。

作者20项与独立40项已通过。边界修复须新增真实最大合法响应，审查后提交。

## 单元二：插件确认、恢复与匿名审计

新增浏览器纯协议合同、controller、storage、card及各自测试，接入App和host-runtime。只处理当前文档当前项目且明确发现能力。预览后点击确认，在请求前持久保存原scope/revision/token/deletionId；ACK未知只读核对，不自动重发删除或更换intent。明确继续必须核同一deletionId及实际控制版本。partial明确正文是否仍保留及共享/归属未明项。策略默认null/null，未实现调度前界面明确策略仅保存。审计导出只有匿名元数据。

测试实际状态机、DOM点击、持久存储失败、文档切换、断线、未知ACK、初始化及数字CAS；包入口不引入Node。

## 单元三：三端能力路由

PC独立presentationGovernanceProxy，实际handler存在才发布presentation-governance.v1。index用真实userData与现有MainService同一项目锁Map，禁止另建锁。Office session和Rust Relay支持完整cap协商及专属路由。仅主PowerPoint连接，不授予TeamOnly连接；拒绝team_context/access_token，七操作白名单，普通presentation.v1不得偷传删除操作。保持原通道预算和已有团队/PDF/保存点合同。

TDD真实PC协议、Office fetch和本地Rust socket路由；错误family、宿主、cap缺失及未知op均拒绝。完成后跨层集成、类型/静态及相关全量回归，再更新工程检查点。无部署、外发或真实资料清理。

## 后续原方案工作

明确保留期的实际到期执行与审计过期需要真实项目活动依据、固定策略版本、时钟及竞态验证，并复用同一冻结/drain/清理事务；不能只依据控制记录createdAt/updatedAt猜内容最近活动。最小禁止复活墓碑必须保留。共享引用归属缺失时保留并报告partial，不能用全局文件删除替代解除项目引用。真实Office与专业任务退出条件分开验收，不能以工程测试数宣称100%。
