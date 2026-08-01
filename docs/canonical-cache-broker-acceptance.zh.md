# Canonical Cache Broker 验收门槛

## 成功定义

Canonical Cache Broker 只有同时证明“单份存储、零重复网络、正确运行、安全隔离、可恢复”才算成功。`transferSize = 0`、前端账本为 `ready` 或单个 Cache Storage 命中都不能单独作为成功证据。

所有测试必须同时采集：

- Chrome 网络事件、Cache Storage、IndexedDB、配额用量和进程生命周期记录。
- Cloudflare Worker 的内容对象请求次数、状态码和下行字节数。
- release manifest、Broker 输出和编辑器实际消费字节的 SHA-256。
- 控制台错误、未处理 Promise、超时、内存和性能数据。

生产协议不得在以下全部门槛通过前启用。

## 固定性能门槛

这些阈值在测试前固定，不能因失败而放宽：

- 资源就绪后的首次编辑器启动，Office segment 网络请求和服务器下行字节均为 `0`。
- Broker 相比纯 HTTP 缓存命中的启动额外耗时，P95 不超过 `max(300 ms, 基线 × 10%)`。
- 单个传输窗口为 `256 KiB–1 MiB`，必须支持背压和取消。
- 禁止把完整 Office Pack 一次读入内存。
- 3 个真实编辑器并发时，Broker 额外在途缓冲总量不超过 `64 MiB`；该上限是协议全局硬限制。
- 连续打开、关闭 100 次后，JS heap、Service Worker 内存和 MessagePort 数量回到稳定区间。

## 分阶段门禁

### G0：存储拓扑原型

- 三个 editor origin 并发安装同一 digest 时，Cloudflare/R2 只成功返回一次。
- canonical 专用 Cache Storage 只有一个对象。
- editor origin 的 Cache Storage 为零，Workbox 通用缓存不含 Office 大资源。
- 清空 HTTP 缓存、浏览器重启和离线后，仍从 canonical 对象读取且服务器计数为零。
- 错误 digest 被拒绝且不持久化。

G0 只证明存储拓扑可行，不代表 Broker 成功。

### G1：协议与字节语义

- Broker 只接受已验证、已激活 release manifest 中的资源 ID，不接受任意 URL 或 Cache key。
- 全资源逐项核对 MIME、长度和 SHA-256。
- Range 的 `200`、`206`、`416`、`Content-Length` 和 `Content-Range` 与原始对象一致。
- WASM 流式编译成功；SDK、字体、字典和 x2t 均可读取。
- 取消 Range 后停止底层读取和传输。
- 分块流支持背压，不发送完整资源 `ArrayBuffer`。

### G2：真实 Office 离线矩阵

- 安装后用 CDP 仅清除 HTTP cache，保留 canonical Cache Storage 与 IndexedDB。
- Aries、Taurus、Gemini 均能启动，content-object Worker 请求数和 R2 下行字节为零。
- 完全退出并重启 Chrome、离线后，新建 Word/Cell/Slide，打开 DOCX/XLSX/PPTX，并完成编辑、保存、重开和 x2t 转换。
- “资源已就绪”要求对象完整、SHA-256 已验证、Broker 读探针成功、Editor SW 协议兼容且没有未完成激活事务。

### G3：3 origin 并发与 12 个产品槽位

- Aries、Taurus、Gemini 三个真实 Office 实例同时运行并复用 canonical 单份资源。
- 相同读取受控合并，不串会话、不死锁、不复制到 editor origin。
- 12 个固定星座槽位、关闭后复用和第 13 个被阻止由轻量状态机、组件及产品集成测试验证，不同时启动 12 个完整 Office 实例。
- 第 13 个编辑器显示中英文提示。

### G4：release 事务和真正增量

- 活动 A 编辑器始终固定 A；B 仅在完整验证和原子激活后用于新编辑器。
- A→B 只下载新增或变化的文件级 CAS 对象，以及大型文件内变化的 FastCDC chunk。
- 未变化 digest 的 Worker 请求和 R2 下行均为零。
- 中断后只补缺失对象；B 校验或激活失败继续使用 A。
- stable 回滚可重新打开 A；垃圾回收不删除活动或保留 manifest 引用的对象。

### G5：恢复、安全和压力

- Broker iframe、Canonical SW、Editor SW、MessagePort、冻结页面和关闭标签页故障均在 30 秒内恢复或进入准确错误状态。
- 已验证对象不重复下载，失败状态不误报 `ready`，未保存文档不因更新刷新。
- 连接使用短期 capability，绑定父 origin、editor origin、release ID 和会话，并拒绝重放。
- 严格允许平台控制的固定 editor origin；所有 `postMessage` 使用精确 target origin。
- 拒绝恶意 origin、伪造消息、跨 release、伪造 digest、路径穿越、超大 Range 和超额并发。
- CSP、`frame-ancestors`、CORS 与跨域响应头只允许预期页面；Broker 不处理 Cookie、文档内容或用户数据。
- 通过固定性能门槛和 100 次生命周期压力测试。

## 增量分块决策

- 常规 SDK、字体、字典和编辑器资源使用文件级 SHA-256 CAS。文件重排或包前插入不会使后续对象失效。
- 对 `>=8 MiB` 的单文件固定使用 FastCDC，平均块 `256 KiB`、最大块 `1 MiB`；这样既支持接近实际变化量的增量更新，也规避 Chrome 将大型网络流直接写入 Cache Storage 时的 `NetworkError`。历史 release 数据用于衡量收益，不再决定是否生成安全分块。
- 不对整个 `.oobpack` 做固定位置分段或全包 FastCDC。冷安装保持“一个完整资源包任务”的产品体验，传输层对每个覆盖计划内容对象的不可变 package segment 只做一次流式 GET，边读边校验并直接写入最终 canonical 内容对象；只含包头的 segment 不下载，完整包和传输 segment 均不得作为 staging 副本持久化。

## 本地与生产门禁

本地矩阵使用真实 Cloudflare Worker 路由、本地 R2、canonical origin 和 Aries、Taurus、Gemini 三个 editor origin。G0 使用最小确定性 release，后续门禁使用真实构建产物和 manifest v5；完整矩阵只在相关协议变化和生产 canary 前运行。

矩阵分成两个不可混淆的层级：

- `synthetic-broker` 是快速协议实验。它使用真实 Worker 和本地 R2，但 `__matrix__` Broker、Editor SW、账本及 2 MiB segment 都是测试实现，只能证明消息、Range、恢复和安全边界，不得作为生产 Broker、生产 Installer 或真实 Office 的通过证据。
- `full-v5` 先执行实际 `pnpm build` 和 `pnpm release:build`，读取 `stable-v5.json`、真实 `manifest.json`、真实 dist、文件级 CAS 和生产 Service Worker/Broker/Installer。它只同时启动 Aries、Taurus、Gemini 三个真实 Office 实例；12 槽、第 13 个阻止和槽位复用继续由轻量测试覆盖。

`full-v5` 必须从 Worker 的 `/__matrix__/content-counters` 同时断言 `workerRequests`、`r2Heads`、`r2Gets` 和 `r2Bytes`，并用 CDP 枚举每个 origin 的 Cache Storage。冷安装时依据真实 manifest 计算覆盖资源内容的 package segment，要求每个 digest 的四项计数分别为 `1/1/1/segment.bytes`，包头专用 segment 和已规划的最终对象均不得被额外拉取。完成安装后，仅清空 HTTP cache，再启动三个编辑器；只有计数完全不变、canonical 只有 `onlyoffice-content-v1` 单份最终对象、没有整包或 segment staging key、editor origin 没有 Office 内容对象时才通过。

“首次成功下载一次”的 full-v5 基线不主动中断传输，以免把用户触发的 abort/retry 混入 exact-once 证据；界面仍检查下载阶段、当前 package segment、进度条和可暂停状态。暂停、取消、超时和断线恢复由聚焦测试另行证明：已验证对象不得重下，当前未完成 segment 允许重试。

同一测试必须使用持久化 Chromium Profile 完成安装，然后真正关闭并重新启动 Chromium。重启后全程离线，依次以每批最多三个真实 Office 实例完成：新建 Word/Cell/Slide、打开现有 DOCX/XLSX/PPTX、编辑表格、保存三种格式、重新打开保存产物并再次保存。最后恢复网络只读取 Worker 计数；内容对象请求、R2 get 和 R2 下行必须相对重启前保持不变。失败时保留 Chromium Profile、保存产物和本地 R2 状态用于复现。

增量矩阵在真实 release A 上生成一个只改变单个有界文件的 v5 B。B 的 manifest、storage-set SHA-256 和变化 blob 是真实 CAS 协议，但 B 是测试派生 release，不是第二次完整生产构建。测试要求 A→B 的 `downloadBytes` 精确等于变化对象、所有未变对象请求为零；更新期间保持一个真实 A 编辑器运行并从 A 读取变化文件，新开的 B 编辑器只能在 B 完整激活后启动；stable 指针切回 A 后要求下载为零并能新开 A 编辑器。全过程真实 Office 实例不超过三个。正式发布前仍须用两个真实历史生产构建复跑同一断言。

本地全部通过后，生产 canary 必须再次同时证明：

1. 服务器 segment 请求计数和 R2 下行字节为零。
2. 三个 editor origin 成功打开真实 Office 文档。
3. Chrome 只持久化 canonical 单份 Office 资源。

## 当前实验状态

截至 2026-07-31，以下结果来自 `synthetic-broker` 最小确定性 release；它们已在真实 Cloudflare Worker、本地 R2 和 Chromium 中证明：

- Aries、Taurus、Gemini 并发安装同一对象时只有一次 segment 请求、一次 R2 get 和一份 canonical Cache Storage 对象。
- 三个 editor origin 的 Cache Storage 均为空，canonical Cache Storage 增量接近一个 2 MiB 对象。
- CDP 仅清空 HTTP cache 后，三端再次读取的 Worker 请求、R2 get 和 R2 下行字节均不增加。
- 断网后仍可读取；错误 digest 不持久化。
- release manifest 白名单、固定 release/session、一次性连接 capability、恶意 origin、capability 重放和跨 release 请求均已验证。
- `200`、`206`、`416`、SHA-256、256 KiB 最大消息窗口、逐块 PULL 背压和取消已在合成 segment 上验证。
- 内容读取已移入 Canonical Service Worker；每个读取使用独立 MessagePort，并由 `ExtendableMessageEvent.waitUntil()` 持有。Editor SW 的 `ReadableStream.pull()` 才请求下一块。
- 主动停止 Canonical 与 Editor Service Worker 后，页面会重建一次性 Broker 通道，已验证对象不重新下载。
- 完全退出 Chromium 后离线重启，Editor SW、Broker 小壳、manifest cache 与 canonical 内容对象可恢复，segment 请求增量为零。
- 合成 release A/B 已验证：活动 A 继续固定 A；B 复用 A 的相同 digest 而不请求网络，只下载 B 新增 digest；A 拒绝读取 B 专属对象。

这仍是合成 segment 协议原型，不代表 Broker 已可生产使用。`full-v5` 现已承担真实构建、生产 Installer/Broker/Editor SW、三 origin、单份存储、清空 HTTP cache 后零请求、A→B/回滚，以及完整 Chromium 重启后的离线新建、打开现有文件、保存和重开检查；在该完整矩阵实际通过并留存报告前，这些项目仍视为未验收。活动 lease/GC、100 次生命周期以及性能/内存门槛仍需单独保留证据。
