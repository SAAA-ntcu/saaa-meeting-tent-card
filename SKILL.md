---
name: meeting-tent-card
description: "Generate, inspect, modify, and export professional Taiwan A4 4-fold triangular meeting tent cards (會議桌牌 / 立體桌牌 / 桌上名牌) in Word (.docx). Use this skill when the user wants to create meeting tent cards from participant lists/Excel/CSV/text, inspect existing tent card docx files, add/modify participants, or customize event headers and branding."
---

# 會議桌牌生成與維護助手 (Meeting Tent Card Skill)

專門用於自動化製作與維護臺灣學術研討會、政府機關審查會、公務會議標準 **A4 四折式立體三角會議桌牌（Tent Card / Name Tent / 桌上名牌）**。

---

## 核心功能與特色

1. **一鍵批次生成（Batch Generation）**：
   - 支援由 **Excel (.xlsx)**、**CSV**、**JSON** 或**直接在對話中貼上純文字名冊**（例如 `單位 | 姓名 | 職稱`），1 秒鐘生成排版完成、立即可列印的 `.docx` 桌牌。
2. **正反 180° 自動翻轉（Dual-Side 180° Mirroring）**：
   - 正面（0°）與背面（180° 倒立）自動雙向同步。
   - 採用微軟 Office `Choice`（DrawingML 向量群組）與 `Fallback`（VML 鏡射）雙層架構，無論在最新版 Office、舊版 Word、WPS 或 Google Docs 檢視，均保證折疊成立體三角柱時雙面正立。
3. **智慧字級與排版縮放（Smart Typography Auto-fit）**：
   - **單位**：標楷體粗體帶光暈。自動判定行政/學術層級斷行（如學校/系所、局處/科室），依長度動態調整字級（38pt → 35pt → 30pt → 26pt），絕不爆字。
   - **姓名職稱**：標楷體置中。自適應 54pt、48pt、40pt，自動保留臺灣公務習慣之姓名與職稱間距。
4. **絕對單頁防跑版保證（Zero-Spill Isolation）**：
   - 每一張桌牌固定劃分 4 段精密折疊高度（`7.43cm` + `8.45cm` + `8.45cm` + `1.97cm` = `26.29cm`），並以獨立 Section Break（節區隔符號）隔離，每位貴賓嚴格佔據 1 頁，絕不跨頁或溢出。
5. **既有桌牌檢測與編輯（Inspect, Update & Append）**：
   - **檢視名冊**：可直接讀取現有桌牌 `.docx`，自動解析並列出全體與會人員、頁碼與空白頁數。
   - **精準修改**：直接依「頁碼」或「姓名」修正錯字、換單位、改職稱，不破壞全檔其餘版面。
   - **臨時追加**：會議前夕若有追加貴賓，可直接追加插入於空白備用頁之前。
   - **名冊匯出**：能將桌牌中的名單反向萃取並匯出為 Excel / CSV / JSON。
6. **空白備用頁保護（Blank Cards Padding）**：
   - 預設於檔案末端自動生成 2～3 張同款式之空白桌牌（保留抬頭與校徽，文字框留空），供會議當天現場報到應急手寫。

---

## 版面力學與折疊規格 (Folding Specifications)

本 Skill 採用的紙張力學規格如下：
* **紙張大小**：標準 A4 直向（`21.0 cm × 29.7 cm`）。
* **版面邊距**：上下左右各 `1.27 cm`（0.5 吋）。
* **單頁 4 段折疊高度劃分**：
  * **第一段（Row 0）- 頂部折面**：`7.43 cm`（210.5 pt）— 錨定主辦單位徽章與繪圖向量物件。
  * **第二段（Row 1）- 背面展示面**：`8.45 cm`（239.5 pt）— **180° 倒立面向**。
  * **第三段（Row 2）- 正面展示面**：`8.45 cm`（239.5 pt）— **0° 正向面向**。
  * **第四段（Row 3）- 底部底卡邊**：`1.97 cm`（55.95 pt）— 底部對折黏合或扣合之底邊。

```
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

## 常用指令與操作方式

核心引擎路徑：
`~/.gemini/config/skills/meeting-tent-card/scripts/tent_card_engine.py`

### 1. 從名冊檔案批次生成桌牌
支援 Excel (`.xlsx`)、CSV、TSV、JSON 或純文字檔：
```powershell
python "C:\Users\adminn\.gemini\config\skills\meeting-tent-card\scripts\tent_card_engine.py" generate `
  -i "C:\path\to\attendees.xlsx" `
  -o "C:\path\to\會議桌牌_output.docx" `
  -b 2
```

### 2. 直接在對話中貼上名單生成（極速模式）
直接使用 `-r` 參數傳入字串（支援 `\n`、分號、直線 `|` 或逗號）：
```powershell
python "C:\Users\adminn\.gemini\config\skills\meeting-tent-card\scripts\tent_card_engine.py" generate `
  -r "國立臺灣大學 心理學系 | 陳台大 | 講座教授\n國立政治大學 教育學院 | 李政大 | 院長\n臺北市政府教育局 | 王小華 | 督學" `
  -o "C:\path\to\今日會議桌牌.docx" `
  -b 2
```

### 3. 更換活動抬頭與主辦單位名稱
若要套用至不同研討會或主辦單位：
```powershell
python "C:\Users\adminn\.gemini\config\skills\meeting-tent-card\scripts\tent_card_engine.py" generate `
  -i "C:\path\to\roster.csv" `
  -o "C:\path\to\研討會桌牌.docx" `
  --event-title "2026年全國教育創新研討會" `
  --sub-title "國立臺灣大學 主辦"
```

### 4. 檢測現有桌牌檔案內容
快速檢查 docx 內共有多少位貴賓、頁碼與名冊：
```powershell
python "C:\Users\adminn\.gemini\config\skills\meeting-tent-card\scripts\tent_card_engine.py" inspect `
  -f "C:\Users\adminn\Downloads\會議桌牌.docx"
```

### 5. 修改特定人員資訊（免重新排版）
* **依頁碼修改**（例如修改第 5 頁）：
  ```powershell
  python "C:\Users\adminn\.gemini\config\skills\meeting-tent-card\scripts\tent_card_engine.py" update `
    -f "C:\path\to\會議桌牌.docx" `
    -p 5 `
    -u "新北市政府教育局中等教育科" `
    -n "黃裕斌" `
    -t "候用校長"
  ```
* **依姓名搜尋修改**（例如更換職稱）：
  ```powershell
  python "C:\Users\adminn\.gemini\config\skills\meeting-tent-card\scripts\tent_card_engine.py" update `
    -f "C:\path\to\會議桌牌.docx" `
    -m "張善貿" `
    -t "特聘教授"
  ```

### 6. 臨時追加新貴賓（自動插入在空白備用頁之前）
```powershell
python "C:\Users\adminn\.gemini\config\skills\meeting-tent-card\scripts\tent_card_engine.py" append `
  -f "C:\path\to\會議桌牌.docx" `
  -u "國家教育研究院 測驗及評量研究中心" `
  -n "陳建宏" `
  -t "副研究員"
```

### 7. 將既有桌牌名冊反向匯出為 Excel 或 CSV
```powershell
python "C:\Users\adminn\.gemini\config\skills\meeting-tent-card\scripts\tent_card_engine.py" export `
  -f "C:\Users\adminn\Downloads\會議桌牌.docx" `
  -o "C:\path\to\與會名冊.xlsx"
```

---

## 智慧字級與斷行規則 (Typography Auto-fit Details)

| 項目 | 條件 | 自動處理機制 | 預設字級 |
| :--- | :--- | :--- | :--- |
| **服務單位** | 短單位（≤ 12 字） | 單行置中靠左縮排，標楷體粗體 + 光暈 | 38 pt (`val="76"`) |
| **服務單位** | 中單位（13～14 字） | 自動於機關/學校名（如「大學」、「教育局」）切分雙行，或字級微調 | 35 pt (`val="70"`) |
| **服務單位** | 長單位（15～16 字） | 切分雙行並縮放字級 | 30 pt (`val="60"`) |
| **服務單位** | 特長單位（> 16 字） | 切分雙行並自適應降至最小安全字級 | 26 pt (`val="52"`) |
| **姓名職稱** | 標準姓名（2~3字）+ 職稱（2~4字） | 置中對齊，字元間隙最適化 | 54 pt (`val="108"`) |
| **姓名職稱** | 長職稱或複姓（8~10 字） | 置中對齊，自適應微調 | 48 pt (`val="96"`) |
| **姓名職稱** | 特長姓名職稱（> 10 字） | 置中對齊，自適應微調 | 40 pt (`val="80"`) |

---

## 檔案結構說明

```
C:\Users\adminn\.gemini\config\skills\meeting-tent-card\
├── SKILL.md                            # Skill 指引與規範文檔
├── templates/
│   └── template_a4_tent.docx           # 單頁 Master 模板（萃取自標準原始檔，內含向量圖與樣式）
├── scripts/
│   └── tent_card_engine.py             # 核心 Python 引擎（generate, inspect, update, append, export）
└── examples/
    ├── sample_roster.csv               # 範例名冊 (CSV)
    ├── sample_roster.xlsx              # 範例名冊 (Excel)
    └── test_generated_cards.docx       # 實測生成之完整桌牌 docx
```
