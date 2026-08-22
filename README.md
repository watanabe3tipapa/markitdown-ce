![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Version](https://img.shields.io/badge/version-0.1.0-green.svg)

# 📂 markitdown-ce

Microsoft Office ドキュメントを Markdown ファイルに変換するためのシンプルなローカル向け Web アプリケーションラッパーです。

## 概要

このプロジェクトは、Microsoft のライブラリ「markitdown」を利用して Office ドキュメントを Markdown に変換する用途を想定した、ローカル実行型の Web アプリケーションです。markitdown 自体はクライアントサイドのブラウザで動作する JavaScript ライブラリではないため、本プロジェクトはサーバー側での変換処理を簡単に利用できるようにするものです。

- markitdown（元リポジトリ）: https://github.com/microsoft/markitdown

## 主な内容

- Office ドキュメント → Markdown 変換（markitdown を活用）
- シンプルなローカル Web インターフェース（Flask を使用）

## 必要なもの

- Python
- Flask
- markitdown

※ 既存 README によると「必要なもの」は列挙されていますが、各種導入手順は省略されています。本 README でもインストール手順は追加していません。

リポジトリには requirements.txt が含まれています。

## ローカルでの使い方（既存 README に記載の手順）

1. リポジトリをクローンしてください（詳細は省略）

Windows の場合

2. ワーキングディレクトリで `python app.py` を実行
3. ブラウザで http://localhost:5000/ にアクセス

macOS の場合

2. ワーキングディレクトリで `python app_mac.py` を実行
3. ブラウザで http://localhost:5555/ にアクセス

備考: macOS 用のポートを 5555 にしているのは、AirPlay 等で使用しているポート 5000 との干渉を避けるためと README に記載されています。

## 補足

- markitdown はブラウザ上でのファイル選択や直接的なダウンロード処理を目的としたクライアントサイド JavaScript ライブラリではありません。本プロジェクトは当該ライブラリをサーバー側で利用するためのシンプルな Web アプリケーションとして構成されています。

## 補助資料（スライド）

- Windowsでの使い方（スライド）: https://docs.google.com/presentation/d/e/2PACX-1vTRfpQqMJwOnIXo3CkFWh1v-aMVsH7JdDiaangU0wrZ_8mQHw1VSmWABBno9BCdyGlX3Tw7FXzKEwQQ/pub?start=true&loop=false&delayms=3000
- macOSでの使い方（スライド）: https://docs.google.com/presentation/d/e/2PACX-1vTO8aVfIsTfUJhmIJUGmDKja_W7-bSpcjrI5bAGMLJ0mZzr1lExGxUq0sQEtHNWy_672wXPeBeINf5n/pub?start=false&loop=false&delayms=3000

## リポジトリ構成（主要ファイル）

このリポジトリのルートに含まれる主なファイル・ディレクトリ:

- .gitignore
- README.md
- app.py
- app_mac.py
- downloads/
- public/
- requirements.txt
- templates/
- uploads/

## 開発・保守状態

- バージョン表示（既存 README バッジ）: 0.1.0
- 最終更新（リポジトリ情報）: 2025-11-08T03:45:23Z

※ リポジトリはアーカイブされていません（archived: false）。それ以上の活動状況やメンテナンス方針は既存 README に記載がないため、本 README では追記していません。

## ライセンス

- MIT ライセンス（既存 README に従う）
