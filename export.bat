@echo off
  
rem
rem bin フォルダの xlsm ファイルから VBA スクリプトを一括エクスポートします。
rem 同じフォルダに、vbac.wsf、bin フォルダが必要です。
rem xlsm ファイルは、bin フォルダに格納します。
rem
  
rem このバッチが存在するフォルダをカレントに設定
pushd %0\..
cls
  
rem エクスポート
cscript vbac.wsf decombine
  
exit