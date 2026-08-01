# OnlyOffice 资源增量更新基准

## 结论

OnlyOffice 资源应继续以“一个 Office 资源包”的产品形态呈现，但底层需要拆分冷安装和增量更新两条传输路径：

- 冷安装使用一个可恢复下载的完整包，并保持约 24 MiB 的逻辑分段，避免上千个小请求。
- 增量更新使用独立于 release ID 的文件级 SHA-256 内容对象，只下载新增或变化的资源。
- 安装账本记录 release、资源 SHA-256 和物理位置；同一 release 可以同时引用基础包内的旧资源和独立内容对象中的新资源。
- 后续 release 生成新的完整基础包，用于新用户冷安装；已有用户不必为了整理存储立即重下完整包。
- 不对整个 `.oobpack` 使用 FastCDC。对 `>=8 MiB` 的单个资源固定使用内容定义分块，以保证增量复用并将每个 canonical Cache Storage 写入限制在 `1 MiB` 内。

这同时保留了整体下载体验和真正的增量更新能力，不需要把冷安装退化成约 1600 个资源请求。

## 当前基准模型

当前 `benchmark:incremental` 使用 report v2，并直接读取 release manifest，不再下载完整 `.oobpack` 后对整个包重新执行 FastCDC。默认读取 manifest v5：

- `/releases/<releaseId>/manifest.json`

使用 `--manifest-version 4` 时读取 companion：

- `/releases/<releaseId>/manifest-v4.json`
- 旧 release 没有 companion 时，兼容回退到其 `/manifest.json`

每一对 release 都按稳定路径逐项规划：

1. 路径和完整 SHA-256 都不变：沿用旧映射，网络下载为零。
2. 新增或变化的普通文件：下载该文件的 whole blob。
3. 新增或变化的大文件，且目标 manifest v5 明确包含 FastCDC representation：只下载相同路径的上一版 FastCDC representation 中不存在的 chunk。
4. v4 没有 FastCDC 元数据，不能根据包偏移猜测 chunk 复用；除完整文件未变外，按 whole blob 处理。
5. 下载对象以 SHA-256 去重，报告同时输出 hybrid planner 的下载字节数、下载对象数、冷态对象数以及各路径决策。

是否使用 FastCDC 由固定的 `8 MiB` 发布阈值和 manifest 决定；benchmark 不会临时对整个资源包分块，也不会把没有 FastCDC 元数据的文件误算为可分段复用。这使报告和实际安装器可执行的计划保持一致。

## 历史数据来源

以下数据由 report v1 在 2026-07-30 对线上三个连续的不可变 release 测得，保留用于说明为什么不能对整个资源包采用固定偏移或 FastCDC：

- `v0.5.0-a5938da84b946739`
- `v0.5.1-712f5189cc253569`
- `v0.5.2-19fe97afff092049`

每个资源包约 567.3 MiB。当时的工具下载并校验 release manifest 与完整包 SHA-256，然后分别比较：

- 当前绑定 release 的固定 24 MiB 分段。
- 独立于 release ID、按 SHA-256 寻址的固定 24 MiB 分段。
- manifest 已有资源边界上的文件级 CAS。
- FastCDC 平均 256 KiB、1 MiB 和 4 MiB 的内容定义分块。

历史 FastCDC 数据使用维护中的 `fastcdc` v2020 实现；每个分段都重新计算 SHA-256，不使用文件偏移或 release ID 判断复用。当前 report v2 不再运行这项 whole-pack 实验。

## 历史实测结果：whole-pack 策略对照

### v0.5.0 → v0.5.1

目标 release 有 1642 个路径内容不变、11 个变化、9 个新增、9 个移除。

| 策略                          | 更新下载量 | 更新对象数 | 冷安装对象数 | 复用率 |
| ----------------------------- | ---------: | ---------: | -----------: | -----: |
| 当前 release 绑定 24 MiB 分段 |  567.3 MiB |         24 |           24 |  0.00% |
| 内容寻址固定 24 MiB 分段      |  567.3 MiB |         24 |           24 |  0.00% |
| 文件级 CAS                    |   1.15 MiB |         17 |         1603 | 99.80% |
| FastCDC，平均 256 KiB         |   4.01 MiB |          8 |         1702 | 99.28% |
| FastCDC，平均 1 MiB           |  13.08 MiB |          7 |          429 | 97.69% |
| FastCDC，平均 4 MiB           |  32.27 MiB |          5 |          114 | 94.31% |

### v0.5.1 → v0.5.2

目标 release 有 1647 个路径内容不变、6 个变化、9 个新增、9 个移除。

| 策略                          | 更新下载量 | 更新对象数 | 冷安装对象数 | 复用率 |
| ----------------------------- | ---------: | ---------: | -----------: | -----: |
| 当前 release 绑定 24 MiB 分段 |  567.3 MiB |         24 |           24 |  0.00% |
| 内容寻址固定 24 MiB 分段      |  72.00 MiB |          3 |           24 | 87.31% |
| 文件级 CAS                    |  762.2 KiB |         12 |         1603 | 99.87% |
| FastCDC，平均 256 KiB         |   3.35 MiB |          7 |         1702 | 99.40% |
| FastCDC，平均 1 MiB           |   7.91 MiB |          5 |          429 | 98.60% |
| FastCDC，平均 4 MiB           |  27.10 MiB |          4 |          114 | 95.22% |

固定分段的复用率会受到包内文件偏移变化影响，同样的构建方式在两个升级样本中分别得到 0% 和 87.31%，不可作为稳定的增量协议。whole-pack FastCDC 能抵抗偏移变化，但在这两个真实 release 中仍比文件级 CAS 多下载约 3–32 MiB。这是历史对照结论，不代表当前 hybrid planner 会对完整包做 FastCDC。

## 偏移变化小规模实验

确定性合成资源额外覆盖了固定分段最不稳定的四种变化。这里的“实际变化”不包含文件移动本身：

| 变化场景                  | 实际变化 | 固定 1 MiB | 稳定逻辑组 | 文件级 CAS | FastCDC 256 KiB | FastCDC 1 MiB |
| ------------------------- | -------: | ---------: | ---------: | ---------: | --------------: | ------------: |
| 包头插入                  |  192 KiB |   9.72 MiB |   1.26 MiB |    192 KiB |       846.2 KiB |      1.12 MiB |
| 包头删除                  |      0 B |   9.54 MiB |   1.07 MiB |        0 B |       654.2 KiB |     959.4 KiB |
| 4 MiB 文件内部替换 64 KiB |   64 KiB |   1.00 MiB |   4.00 MiB |   4.00 MiB |       446.4 KiB |      1.93 MiB |
| 文件重排                  |      0 B |   2.00 MiB |        0 B |        0 B |        1.03 MiB |      2.12 MiB |

这验证了两个不同结论：

- 包头插入、删除和文件重排应由文件级 CAS 解决；固定位置分段会因后续偏移整体变化而失效。
- FastCDC 只在单个大型文件内部发生局部变化时明显优于文件级 CAS，因此应作为大文件的第二级分块，而不是整个资源包的统一分块方式。

## 字体边界

字体列表整理不应删除 OnlyOffice 运行时依赖的字体，也不应改写 DocumentServer 原生的缺失字体替代算法。

- 字体选择器只显示经过产品审核的常规中文和西文字体。
- 辅助字体、符号字体和兼容字体可以继续安装，供原生渲染与 fallback 使用，但不进入用户可选列表。
- 新建文档的默认西文字体固定为 Aptos，默认中文字体固定为等线。
- “隐藏字体”是 Host 配置或字体元数据变化，不应导致完整字体资源重新下载。

因此，字体整理应成为验证增量协议的回归场景：当字体二进制没有变化时，升级只能请求变化的 Host/manifest 内容对象。

## 建议协议边界

manifest v5 和混合资源存储遵循以下边界：

1. release manifest 中每个资源继续携带路径、字节数和完整 SHA-256。
2. R2 同时发布冷安装完整包和 `blobs/sha256/<digest>` 增量对象。
3. 安装账本把资源映射到一个或多个内容对象 span；冷安装的 package segment、whole blob 和固定阈值生成的 FastCDC chunk 都可以成为 span 来源，不需要复制成第二份资源。
4. 更新规划器先按路径复用完整未变资源；普通变化选择 whole blob；只有 manifest v5 已声明 FastCDC 的大型变化文件才比较 chunk，并仅获取缺失对象。
5. 失败时保留旧 release 和旧账本；新 release 只有在全部必需资源通过校验后才原子激活。
6. 后续按存储压力和覆盖 release 数量进行后台压实，不把压实作为版本更新的前置条件。

## 复现

默认读取 v5。以下命令只请求两个 manifest，不下载完整资源包，也不运行 Rust：

```bash
pnpm benchmark:incremental -- \
  --release <release-v5-a> \
  --release <release-v5-b> \
  --output /tmp/onlyoffice-incremental-benchmark/report.md
```

要重放上面的三个 v4 历史 release，可显式选择 companion/旧 manifest：

```bash
pnpm benchmark:incremental -- \
  --manifest-version 4 \
  --release v0.5.0-a5938da84b946739 \
  --release v0.5.1-712f5189cc253569 \
  --release v0.5.2-19fe97afff092049 \
  --output /tmp/onlyoffice-incremental-benchmark/report-v4.md
```

v4 报告仍会给出固定 package segment 和 file-level whole CAS 基线；hybrid planner 不会伪造 FastCDC 收益。whole-pack FastCDC 的小规模实验由独立的 `experiment:chunking` 保留，不再属于真实 release 更新计划。
