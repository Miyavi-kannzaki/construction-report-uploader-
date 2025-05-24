# construction-report-uploader（匿名化版）

このツールは、工事完了報告書を自動で処理・リネーム・Boxへ格納する Python製の業務効率化アプリです。

## ✅ 特徴
- Excelのマクロを自動実行し、ファイル名と保存先を構築
- CSV/Excel をBoxの適切なフォルダに自動で振り分け
- GUI付きで誰でも操作可能
- `.exe化` すれば非エンジニアでも使用可能

## 📦 ファイル構成
- `怠惰の極み乙女ver1.2.py`：メイン処理スクリプト
- `README.md`：この説明ファイル

## 🧑‍💻 補足
- 実行には Windows + Python + Excel が必要です
- `win32com.client` を使用します（`pip install pywin32`）
- 個人情報・企業情報は全てマスク済みです

## 🧪 使用例
```bash
python construction-report-uploader.py
