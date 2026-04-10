# -*- coding: utf-8 -*-
"""
word_to_pdf.py
同じフォルダ内のすべての Word ファイル (.doc / .docx) を PDF に変換し、
変換済みの Word ファイルを 'converted' フォルダに移動します。

【必要なもの】
  - Windows に Microsoft Word がインストールされていること
  - Python がインストールされていること（https://www.python.org/）
  - pywin32 ライブラリ（初回のみ自動インストールを試みます）
"""

import os
import sys
import time
import shutil
import subprocess


def install_pywin32():
    """pywin32 を自動インストールする"""
    print("[情報] pywin32 をインストールしています...")
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "pywin32"],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        print("[情報] インストール成功。変換を開始します。\n")
        return True
    else:
        print("[エラー] インストールに失敗しました。")
        print("手動で以下を実行してください:")
        print("    pip install pywin32")
        return False


def convert_word_to_pdf():
    # このスクリプトと同じフォルダを対象にする
    script_dir = os.path.dirname(os.path.abspath(__file__))
    converted_dir = os.path.join(script_dir, "converted")

    print("=" * 50)
    print("   Word → PDF 一括変換ツール")
    print("=" * 50)
    print(f"対象フォルダ: {script_dir}\n")

    # Word ファイルを探す（一時ファイル ~$ は除外）
    word_files = sorted([
        f for f in os.listdir(script_dir)
        if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')
    ])

    if not word_files:
        print("[警告] Wordファイルが見つかりませんでした。")
        input("\nEnterキーを押すと終了します...")
        return

    print(f"変換対象: {len(word_files)} 件")
    for f in word_files:
        print(f"  - {f}")
    print()

    # pywin32 のインポートを試みる（なければ自動インストール）
    try:
        import win32com.client
    except ImportError:
        if not install_pywin32():
            input("\nEnterキーを押すと終了します...")
            sys.exit(1)
        try:
            import win32com.client
        except ImportError:
            print("[エラー] インストール後も読み込めませんでした。")
            print("一度このウィンドウを閉じて、再実行してください。")
            input("\nEnterキーを押すと終了します...")
            sys.exit(1)

    # converted フォルダを作成（既に存在していても OK）
    os.makedirs(converted_dir, exist_ok=True)

    # Word アプリケーションを起動
    word = None
    success = 0
    failed = 0
    failed_files = []

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        for filename in word_files:
            input_path = os.path.join(script_dir, filename)
            output_path = os.path.splitext(input_path)[0] + ".pdf"

            print(f"変換中: {filename}", end="", flush=True)

            doc = None
            try:
                doc = word.Documents.Open(input_path)
                doc.SaveAs(output_path, FileFormat=17)  # 17 = wdFormatPDF
                doc.Close(SaveChanges=False)
                doc = None

                # 変換成功したら converted フォルダへ移動
                dest_path = os.path.join(converted_dir, filename)
                # 同名ファイルが既にある場合はタイムスタンプを付けてリネーム
                if os.path.exists(dest_path):
                    base, ext = os.path.splitext(filename)
                    dest_path = os.path.join(converted_dir, f"{base}_{int(time.time())}{ext}")
                shutil.move(input_path, dest_path)

                print(" ✓  → converted/ に移動")
                success += 1

            except Exception as e:
                print(f" ✗  ({e})")
                failed += 1
                failed_files.append(filename)
            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass

            time.sleep(0.2)

    except Exception as e:
        print(f"\n[エラー] Word の起動に失敗しました: {e}")
        print("Microsoft Word がインストールされているか確認してください。")
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

    # 結果サマリー
    print()
    print("=" * 50)
    print(f"  完了: 成功 {success} 件 ／ 失敗 {failed} 件")
    if success > 0:
        print(f"  変換済みWordファイルの移動先: converted/")
    if failed_files:
        print("  失敗したファイル（移動していません）:")
        for f in failed_files:
            print(f"    - {f}")
    print("=" * 50)
    input("\nEnterキーを押すと終了します...")


if __name__ == "__main__":
    convert_word_to_pdf()
