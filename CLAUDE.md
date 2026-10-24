# 上市櫃公司永續報告書聯絡名單

用途：整理台灣上市、上櫃公司永續報告書上的聯絡窗口（email、聯絡人、電話），寄送 IFRS 文宣。
使用者不是工程師，回覆一律用繁體中文、白話說明，不要只丟技術細節。

## 交檔規則（必須遵守）

**每次產出或更新 Excel，交給使用者之前都要先自我檢查：**

1. 執行 `python3 scripts/qa_check.py output/上市櫃公司永續報告書聯絡email.xlsx`，問題數必須為 0；
   有問題就修正後重跑，誤報要確認原因並修正檢查規則，不可直接忽略。
2. 隨機抽 15 家以上，把 H～L、F、U 欄和報告書原文逐一比對（內容正確性 qa_check 抓不到）。
3. 有改到擷取邏輯時，比對新舊版差異（哪些公司的 email／電話／聯絡人變了），逐筆確認變動合理。
4. 回覆時簡要說明檢查了什麼、發現並修正了什麼。

## 資料流程

| 步驟 | 指令 | 說明 |
|---|---|---|
| 1 | Open API 下載 `L.json`、`O.json` | 證交所 `t187ap03_L`、櫃買 `mopsfin_t187ap03_O` 公司基本資料 |
| 2 | `python3 scripts/fetch_reports.py reports.json <年度>` | ESG InfoHub 永續報告書申報資料（含聯絡資訊、PDF 檔案 id） |
| 3 | `node scripts/extract_contacts.mjs reports.json contacts.jsonl 10` | 用 HTTP Range 只讀 PDF 聯絡資訊頁；沒找到的用 `DEEP=1` 再跑第二輪 |
| 4 | `python3 scripts/trim_snippets.py ...` | 裁出聯絡段落，**逐份人工判讀**部門／職稱／姓名，存成 parse/manual_*.tsv、deep_*.tsv |
| 5 | `python3 scripts/merge_contacts.py ...` | 合併判讀結果，擷取電話、挑最高職級聯絡人、產生稱呼 → `data/contacts_parsed.tsv` |
| 6 | `python3 scripts/build_excel.py <資料夾> output/...xlsx` | 產出 Excel |
| 7 | `python3 scripts/qa_check.py output/...xlsx` | 交檔前自我檢查（見上方規則） |

環境需求：網路需允許 `twse.com.tw`、`tpex.org.tw`；`npm i pdfjs-dist@4.10.38 undici@6.21.1`、`pip install pypinyin openpyxl`。
ESG InfoHub 會擋過多連線，平行數不要超過 10。

## 使用者確認過的規則

- **F 欄寄送 email**：以報告書內文 email（N 欄，和聯絡人同一段）為主，沒有才用申報資料（M 欄），再沒有用公司信箱。
- **聯絡人**：多位時，先用 email 帳號的姓氏拼音比對（candy.chiu → 邱、davishan → 韓），email 對得上的那位優先；
  寄送 email 是公用信箱（esg@、public@）而申報資料／內文有某位聯絡人的個人信箱時，改留那位並寄個人信箱（只有一位聯絡人也適用；名字縮寫加姓 shpan → 潘信宏 也算；看不出姓的如 cally@ 不換）。
  都對不上才留職級最高者。一人多職稱只留最高的。被移除的用 I、J 欄「註解」標示（含比對依據），不另開欄位。
- **職級**：董事長 > 總經理/執行長 > 副總經理/財務長/永續長 > 協理/總監/處長 > 經理/副理 > 課長/主任/組長 > 專員/秘書/特助 > 發言人/公司治理主管/執行秘書等角色。
- **稱呼（K 姓、L 職稱）**：職稱用全稱（副總經理，不寫副總；特別助理，不寫特助）；執行秘書 → 王執行秘書；
  報告書寫「劉小姐」照用；英文名 K 填名字、L 空白（信件用「Wendy 您好」）；
  只有姓名或只有角色職稱（發言人、公司治理主管）→ 依序用 ① 報告書原文寫明 ② email 英文名 ③ 中文名字（逐一人工判讀，存在 parse/gender_by_name.tsv）
  判斷先生/小姐；③ 沒把握的仍填最可能的，但 L 欄黃底、W 欄標「🟡 先生／小姐依名字推測，請確認」。L 欄註解寫明判斷依據。
  沒有人名才 K、L 空白（信件用「您好」）。
- **U 欄**：報告書上聯絡人的電話（含分機、排除傳真與客服專線），V 欄才是總機。
- **S、T 欄**：超連結。
- **W 欄需確認事項**：🔴 必須處理（紅底）、🟡 建議確認（黃底）。程式判斷不了的疑慮寫在 `data/manual_flags.tsv`；W 欄有內容的公司會集中列在 Excel 的「需確認清單」分頁，並寫明要確認什麼。
- 使用者公司無法上傳可編輯檔案到 Claude，只能貼文字或上傳 PDF。

## 檔案位置

- `data/`：公司基本資料、申報資料、擷取的報告書文字、合併結果（`contacts_parsed.tsv`）、`manual_flags.tsv`。
- `parse/`：逐家人工判讀的結果（`manual_*.tsv`、`deep_*.tsv`、`gender_by_name.tsv`）與裁出的聯絡段落。
- 不連網也能重現：`merge_contacts.py` → `build_excel.py` → `qa_check.py`，結果應與 `output/` 的 Excel 相同。
