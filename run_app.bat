@echo off
REM このスクリプトは、Pythonの環境変数が設定されているWindows環境で、
REM 'tag_classification_app_tkinter.py' アプリケーションを実行します。

REM 必要なライブラリ (pandas) をインストールします。
REM すでにインストールされている場合はスキップされます。
echo 必要なライブラリをインストールしています...
pip install pandas
if %errorlevel% neq 0 (
    echo.
    echo エラー: pandasのインストールに失敗しました。
    echo Pythonが正しくインストールされ、環境変数PATHが設定されているか確認してください。
    echo.
    pause
    exit /b %errorlevel%
)
echo インストール完了。

REM Pythonスクリプトを実行
echo アプリケーションを起動しています...
python tag_classification_app_tkinter.py

REM アプリケーションが閉じられたら、このウィンドウを閉じるのを待つ
pause
