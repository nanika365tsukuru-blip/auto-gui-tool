# Auto GUI Tool Professional (Minimal Production Setup)

本番運用向けの最小構成（GitHub 用）。GUI 自動化ツール本体は `src/Auto_gui_tool.py`。  
Windows 最適化（他 OS でも動く部分あり）。

## 要件
- Windows 11
- Python 3.12
- `pip install -r requirements.txt`

## 起動
```bash
python .\src\Auto_gui_tool.py
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

[![デモ: 動作イメージ](https://img.youtube.com/vi/wOKixLcCThY/hqdefault.jpg)](https://www.youtube.com/watch?v=wOKixLcCThY "Auto GUI Tool - 動作イメージ")

