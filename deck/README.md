# 构建与设计系统

## 页面结构（10页 + 附录）

| # | 页 | 讲什么 |
|---|---|---|
| 1 | 封面 | 每个行业最强的AI，都出自我们的世界 |
| 2 | 一页讲完 | 短期做AI Infra数据，长期覆盖整个AI经济；内核全时段越来越强 |
| 3 | 团队 | 四个平均21岁的量化人；机构标识墙（Jane Street、Citadel、D. E. Shaw、Millennium、Optiver、剑桥） |
| 4 | 核心洞察 | 每次跃迁都来自反馈信号升级，第四代来自世界 |
| 5 | 三层套娃 | 模型 / 世界 / 造世界的引擎，每层配证据；RSI for RSI |
| 6 | 越来越强 | 六个引擎；表象会变，内核只增不减 |
| 7 | 短期 | AI Infra数据：市场、对标、怎么赚钱、为什么是我们 |
| 8 | 长期 | AI经济的每一个阶段、每一个行业；智能体市场与自主经济 |
| 9 | 进展与融资 | 已经做到 → 24个月里程碑 → $20M |
| 10 | 结尾 | 联系方式 |
| A1–A4 | 附录 | Xitadel测试方法、产品清单、数据来源、术语表 |

机构标识：`assets/logos/*-color.png` 由 `tools/logos.py` 从团队提供的原图去底生成（Citadel 为海军蓝字标）。

## 英文版（SimReal-Seed-Deck-EN.pptx）

```bash
cd deck
./build_en.sh     # node build_en.js -> tools/postprocess.py ... en -> ../SimReal-Seed-Deck-EN.pptx
```

`build_en.js` 对应团队在 Google Slides 里改定的中文版（19页 + A1–A4），英文措辞沿用此前英文草稿。
为了在 Google Slides 里导入后不出错：

- 横条图不再翻转坐标轴（Google 会忽略翻转，导致标签和数据错位），改为把数据倒序传入，并使用图表自带的分类标签；
- 图表分类改用单层 `strRef`（Google 读不了 pptxgenjs 默认的多层格式，会显示成 1、2、3）；
- 加粗用同一字体的粗体（Instrument Sans Bold），不用单独的 SemiBold 字体名；三套字体都是 Google Fonts。

## 构建

```bash
cd deck
./build.sh        # node build.js -> tools/postprocess.py -> ../SimReal-BP-ZH.pptx
```

依赖：Node + `pptxgenjs`；Python 3 + `fonttools`；系统里的 Noto CJK 字体
（`fonts-noto-cjk`、`fonts-noto-cjk-extra`，用于中文字体子集）。

- `build.js`：23页的全部版式与内容。字体以占位符（`XSERIF / XSANS / XSANSB / XMONO`）写入。
- `tools/postprocess.py`：把占位符换成「西文 + 中文」字体对；修正图表XML顺序并写入自定义数据标签；
  清理 pptxgenjs 在段落中间多写的 `<a:pPr>`；嵌入字体。
- `tools/fonts.py`：按PPT实际用到的字符对每个字体做子集，中文字体从CFF轮廓转为TrueType
  （PowerPoint只能嵌入TrueType），再封装为 EOT `.fntdata`。
- `tools/eot.py`：EOT写入器。用原版PPT里嵌入的三个字体验证过，输出逐字节一致。

## 设计系统

**颜色**：纸白 `FAF9F6`、墨黑 `111110`、正文 `33322F`、灰 `6E6C66`、分隔线 `E2E0DA`、朱红强调 `C2410C`
（深色底上用 `E27A3F`）。沿用原版品牌色，强调色只用于每页最关键的一处。

**字体**（全部嵌入文件，投资人电脑上没装也能按设计显示）：

| 角色 | 西文 | 中文 |
|---|---|---|
| 标题、大数字、结论句 | Newsreader | 思源宋体 Noto Serif CJK SC |
| 正文 | Instrument Sans | 思源黑体 Noto Sans CJK SC |
| 加重（标签、要点） | Instrument Sans SemiBold | 思源黑体 Medium |
| 编号、单位、来源标注 | IBM Plex Mono | 思源黑体 |

**字号**：原版用了36种字号（7.5–66pt），现在收敛为一套字号阶梯：
54 封面 / 44 结尾 / 60、40 大数字 / 28 页标题 / 22、18 小标题 / 20 结论句 / 15 副标题 / 13 正文 /
11 次要正文 / 10 标签 / 9 来源注释 / 8 页脚。最小可读字号从 7.5pt 提到 9pt。

**网格**：左右边距 0.6"，内容宽 12.13"；页眉（编号 + 栏目）→ 标题 28pt → 副标题 → 内容区 → 结论句 → 页脚，
每页位置一致。

**强调方式**：原版用手工摆放的橙色色块做"荧光笔"，换字体或换电脑就会错位；现在改为直接给关键词上朱红色，
不会漂移。

## 图表（全部为原生、可编辑的 PowerPoint 图表）

- 第7页：头部AI训练数据公司收入（横条图，每条带口径说明）
- 第8页：AI智能体市场 2025–2032年路径（按 Precedence Research 45.8% 年复合增长推算）
- 附录A1：Xitadel公开预览版，人类参考分80 vs 前沿模型最高77.28

## 内容原则

所有数字和事实都来自原版PPT及其数据来源页；叙事按团队定的"短期切入AI Infra数据、长期覆盖整个AI经济"重写。
