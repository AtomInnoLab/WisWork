# 原方案阶段0/O1：字体失败定位与恢复

基线8d0a3cd7。原方案要求任一失败诊断可还原到具体阶段/页面、字体回退或明确失败、持久中断恢复。实际compile与PC映射支持font_unavailable，Office解析/文案也支持，ProjectStore生产页和job事件白名单却缺该值。先以实际PC服务复现失败页无法记账，再修共享源而非降格compile_failed。

1. ProjectStore：presentation-job.ts导出浏览器安全共享PRESENTATION_PRODUCTION_ERRORS常量与类型（旧六种+font_unavailable），job和presentation-store.ts用同一来源严格校验；新真实持久测试覆盖字体失败/重启、原错误和未知拒。无宽泛任意字符串放行。
2. Office：presentation-production.ts复用共享错误集合，验证job/foreground字体失败实际响应接受、工作台中文失败说明和恢复动作；按现有真实输入合同测试，不改确认或QA状态。
3. Root：实际Service前台失败、后台job页失败事件、重启读回、成功字体后恢复同请求保持已编译页，定位真实页ID而不是全局invalid_state；独立交叉复审后完整八工作区、相关类型/静态、Office构建和进度记录。

回退代码即可，无数据迁移、外部写入/部署。保留旧已合法记录；新版新增错误值对旧PC的版本兼容行为遵循既有能力门禁，不假设旧版本能读取新版磁盘。合成字体失败注入证明错误链，真实字体/宿主视觉验收仍待进行。
