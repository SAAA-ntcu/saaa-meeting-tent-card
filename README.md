# SAAA 會議桌牌自動化生成與維護系統 (Meeting Tent Card Skill)

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Format: Word .docx](https://img.shields.io/badge/Format-Word%20.docx-2b579a.svg)](https://www.microsoft.com/word)
[![Platform: Cross-Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-brightgreen.svg)]()

> 專為臺灣學術研討會、政府機關審查會、公務會議量身打造之標準 **A4 四折式立體三角會議桌牌（Tent Card / Name Tent / 立體桌牌）** 自動化生成與維護工具。相容於 Google Antigravity Skill、Claude Projects、Cursor 及獨立 Python CLI。

---

## 視覺與力學架構 (Folding Mechanism)

本工具產出之桌牌採用臺灣公務與學術界標準之 **A4 直式四折立體三角桌牌工學規範**：
* **紙張大小**：標準 A4 直向（`21.0 cm × 29.7 cm`）。
* **版面邊距**：上下左右各 `1.27 cm`（0.5 吋）。
* **單頁 4 段精確高度分割**（總高 `26.29 cm`，完美契合可列印區）：
  * **第一段（Row 0）- 頂部折面**：`7.43 cm`（210.5 pt）— 錨定主辦單位標誌與繪圖向量物件。
  * **第二段（Row 1）- 背面展示面**：`8.45 cm`（239.5 pt）— **180° 倒立面向**（面向會議桌內側）。
  * **第三段（Row 2）- 正面展示面**：`8.45 cm`（239.5 pt）— **0° 正向面向**（面向觀眾或對座）。
  * **第四段（Row 3）- 底部底卡邊**：`1.97 cm`（55.95 pt）— 底部對折黏合或扣合之底邊。

```text
+---------------------------------------------------+
|               Row 0: 頂部折面 (7.43 cm)           |
+---------------------------------------------------+  <-- 折線 1
|                                                   |
|          Row 1: 背面展示面 (8.45 cm, 180° 倒立)   |
|               (與會人員面向會議桌內側)            |
|                                                   |
+---------------------------------------------------+  <-- 折線 2 (頂部山線折)
|                                                   |
|          Row 2: 正面展示面 (8.45 cm, 0° 正立)     |
|               (與會人員面向觀眾或對座)            |
|                                                   |
+---------------------------------------------------+  <-- 折線 3
|               Row 3: 底部卡榫邊 (1.97 cm)         |  <-- 與 Row 0 黏合/卡合底座
+---------------------------------------------------+
```

---

## 核心特色

1. **正反 180° 自動翻轉（Dual-Side 180° Mirroring）**：
   * 現代 Word（DrawingML `Choice`）：使用向量群組與 `rot="10800000"`（180° 精準翻轉）。
   * 舊版 Word（VML `Fallback`）：使用 `style="flip:x y"` 水平垂直雙向鏡射。
   * 折疊成立體三角柱時，**正反兩面皆能正立面向兩側人員**。
2. **智慧字級與排版縮放（Smart Typography Auto-fit）**：
   * **服務單位**：標楷體粗體 + 向量光暈。長單位自動依「大學/系所、局處/科室」行政層級智慧切分雙行；依長度自動調節字級（38pt → 35pt → 30pt → 26pt），絕不爆字。
   * **姓名職稱**：標楷體置中。自適應 54pt、48pt、40pt，自動保留臺灣公務習慣之姓名與職稱間距。
3. **絕對單頁防跑版（Zero-Spill Isolation）**：
   * 每張桌牌以獨立 Section Break（節符號）隔離，每位貴賓嚴格佔據 1 頁，絕不跨頁或溢出。
4. **既有桌牌檢測與修訂（Inspect, Update & Append）**：
   * 可直接讀取現有桌牌 `.docx`，解析列出所有貴賓清單。
   * 支援依「頁碼」或「姓名」修正錯字、換單位、改職稱，不破壞全檔其餘版面。
   * 會議前夕若有追加貴賓，可直接追加插入於空白備用頁之前。
5. **空白備用頁保護（Blank Cards Padding）**：
   * 批次生成時，預設自動在文檔末端附加 2～3 張空白桌牌（保留抬頭與校徽，文字框留空），供現場報到手寫應急。

---

## 目錄結構

```text
meeting-tent-card/
├── SKILL.md                            # Agent Skill 規範文件
├── README.md                           # GitHub 專案說明文件
├── .gitignore                          # Git 忽略配置
├── templates/
│   └── template_a4_tent.docx           # 單頁 Master 模板（內含高解析度向量徽章與樣式）
├── scripts/
│   └── tent_card_engine.py             # 核心 Python 引擎 (generate, inspect, update, append, export)
└── examples/
    ├── sample_roster.xlsx              # 範例名冊 (Excel)
    ├── sample_roster.csv               # 範例名冊 (CSV)
    └── sample_cards_output.docx        # 實測生成之完整桌牌範例文檔
```

---

## 快速開始與使用教學

### 方式 1：作為 Antigravity Agent Skill 使用
已安裝至 `~/.gemini/config/skills/meeting-tent-card`。在對話中直接指示 Agent 即可：
> 「請根據這份名單幫我產出 25 位貴賓的會議桌牌 docx，並在最後留 2 張空白備用。」
> 「幫我把桌牌第 5 頁的職稱改成『候用校長』。」
> 「檢查這份會議桌牌.docx，幫我匯出所有貴賓清單到 Excel。」

### 方式 2：使用 Python 獨立命令列 (CLI)

#### 1. 從名冊檔案批次生成桌牌（支援 `.xlsx`, `.csv`, `.tsv`, `.json`）
```bash
python scripts/tent_card_engine.py generate \
  -i examples/sample_roster.xlsx \
  -o 會議桌牌_output.docx \
  -b 2
```

#### 2. 直接由純文字名冊字串生成（極速模式）
```bash
python scripts/tent_card_engine.py generate \
  -r "國立臺灣大學 心理學系 | 陳台大 | 講座教授\n國立政治大學 教育學院 | 李政大 | 院長\n臺北市政府教育局 | 王小華 | 督學" \
  -o 今日會議桌牌.docx \
  -b 2
```

#### 3. 更換活動抬頭與主辦單位名稱
```bash
python scripts/tent_card_engine.py generate \
  -i examples/sample_roster.csv \
  -o 研討會桌牌.docx \
  --event-title "2026年全國教育創新研討會" \
  --sub-title "國立臺灣大學 主辦"
```

#### 4. 檢測既有桌牌內容
```bash
python scripts/tent_card_engine.py inspect -f 會議桌牌.docx
```

#### 5. 修改特定人員資訊（免重新排版）
```bash
# 依頁碼修改
python scripts/tent_card_engine.py update -f 會議桌牌.docx -p 5 -u "新北市政府教育局中等教育科" -n "黃裕斌" -t "候用校長"

# 依姓名搜尋修改
python scripts/tent_card_engine.py update -f 會議桌牌.docx -m "張善貿" -t "特聘教授"
```

#### 6. 臨時追加新貴賓（自動插入在空白備用頁之前）
```bash
python scripts/tent_card_engine.py append -f 會議桌牌.docx -u "國家教育研究院 測驗及評量研究中心" -n "陳建宏" -t "副研究員"
```

#### 7. 將既有桌牌名冊反向匯出為 Excel 或 CSV
```bash
python scripts/tent_card_engine.py export -f 會議桌牌.docx -o 與會名冊.xlsx
```

---

## 字級與斷行規則對照表

| 項目 | 條件 | 自動處理機制 | 預設字級 |
| :--- | :--- | :--- | :--- |
| **服務單位** | 短單位（≤ 12 字） | 單行置中靠左縮排，標楷體粗體 + 光暈 | 38 pt (`val="76"`) |
| **服務單位** | 中單位（13～14 字） | 自動於機關/學校名（如「大學」、「教育局」）切分雙行 | 35 pt (`val="70"`) |
| **服務單位** | 長單位（15～16 字） | 切分雙行並縮放字級 | 30 pt (`val="60"`) |
| **服務單位** | 特長單位（> 16 字） | 切分雙行並自適應降至最小安全字級 | 26 pt (`val="52"`) |
| **姓名職稱** | 標準姓名（2~3字）+ 職稱（2~4字） | 置中對齊，字元間隙最適化 | 54 pt (`val="108"`) |
| **姓名職稱** | 長職稱或複姓（8~10 字） | 置中對齊，自適應微調 | 48 pt (`val="96"`) |
| **姓名職稱** | 特長姓名職稱（> 10 字） | 置中對齊，自適應微調 | 40 pt (`val="80"`) |

---

## 授權條款 (License)

本專案採用 [MIT License](LICENSE) 授權。
