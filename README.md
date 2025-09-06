# Auto GUI Tool Professional

GUI 自動化ツール本体は `src/Auto_gui_tool.py`。  
動作環境はWindows 環境を前提としています。

## 要件
- Windows 11
- Python 3.12
- `pip install -r requirements.txt`

## 起動

### 通常起動（GUIモード）
```bash
python .\src\Auto_gui_tool.py
```

### バッチ実行（引数指定）
```bash
python .\src\Auto_gui_tool.py .\config\設定ファイル.json
```

**バッチ実行の特徴:**
- 10秒のカウントダウン後に自動実行開始
- 実行中はウィンドウが最小化される
- 実行完了後5秒でプログラム終了
- リターンコードで実行結果を判定可能
  - `0`: 正常終了
  - `1`: 実行エラー（画像認識失敗等）
  - `2`: 起動エラー（設定ファイル不正等）

**リターンコード確認例（Windows）:**
```cmd
python .\src\Auto_gui_tool.py .\config\機能デモ用.json
echo 実行結果: %ERRORLEVEL%
```

## 構成
```
.
├─ README.md
├─ LICENSE
├─ requirements.txt
├─ .gitignore
├─ .gitattributes
└─ src/
   └─ Auto_gui_tool.py
```

## 運用メモ
- 生成物（`logs/`, `backups/`, `temp/`）は Git 管理外（`.gitignore` 済み）
- 改行は `.gitattributes` で正規化（CRLF 警告の抑止）
- ライセンスは MIT


## デモ動画（Youtube）

[![デモ: 動作 イメージ](https://img.youtube.com/vi/wOKixLcCThY/hqdefault.jpg)](https://www.youtube.com/watch?v=wOKixLcCThY "Auto GUI Tool - 動作 イメージ")

## デモ動画 設定方法（Youtube）

[![デモ: 設定方法 イメージ](https://img.youtube.com/vi/-LTGj1-l5VA/hqdefault.jpg)](https://www.youtube.com/watch?v=-LTGj1-l5VA "Auto GUI Tool - 設定方法 イメージ")



