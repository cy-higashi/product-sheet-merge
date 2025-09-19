#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Phase5処理の詳細追跡

基準位置検出失敗ファイルがなぜall_collect.xlsxに含まれるのかを調査
"""

import os
import pandas as pd
from openpyxl import load_workbook

def trace_phase5_processing():
    municipality = "2025-09-09_福岡県北九州市"
    base_dir = f"G:\\共有ドライブ\\★OD\\99_商品管理\\DATA\\Phase3\\HARV\\{municipality}"
    
    print("=== Phase5処理追跡 ===")
    
    # 1. _normalized.xlsxファイルの確認
    if os.path.exists(base_dir):
        files = [f for f in os.listdir(base_dir) 
                if f.startswith("PAT") and f.endswith("_normalized.xlsx") 
                and not f.endswith("_normalized_normalized.xlsx")]
        
        print(f"_normalized.xlsxファイル数: {len(files)}")
        
        # PAT0004_normalized.xlsxの確認
        pat0004_normalized = os.path.join(base_dir, "PAT0004_normalized.xlsx")
        if os.path.exists(pat0004_normalized):
            print(f"PAT0004_normalized.xlsx が存在します")
            
            # ファイルの内容確認
            try:
                df = pd.read_excel(pat0004_normalized)
                print(f"  形状: {df.shape}")
                print(f"  列名: {list(df.columns)[:10]}")  # 最初の10列
                print(f"  サンプルデータ:")
                for i, row in df.head(3).iterrows():
                    print(f"    行{i+1}: {list(row)[:5]}")  # 最初の5列
            except Exception as e:
                print(f"  読み込みエラー: {e}")
        else:
            print("PAT0004_normalized.xlsx が存在しません")
    
    # 2. all_collect.xlsxの確認
    all_collect_file = os.path.join(base_dir, "all_collect.xlsx")
    if os.path.exists(all_collect_file):
        print(f"\nall_collect.xlsx が存在します")
        
        try:
            df = pd.read_excel(all_collect_file)
            print(f"  形状: {df.shape}")
            print(f"  列名: {list(df.columns)[:10]}")  # 最初の10列
            
            # ハマダ関連レコードの確認
            if 'ファイル名' in df.columns:
                hamada_records = df[df['ファイル名'].astype(str).str.contains('ハマダ', na=False)]
                print(f"  ハマダ関連レコード: {len(hamada_records)}件")
                
                if len(hamada_records) > 0:
                    print("  ハマダレコードの詳細:")
                    for i, (idx, row) in enumerate(hamada_records.head(3).iterrows()):
                        print(f"    レコード{i+1}: ファイル名={row.get('ファイル名', 'N/A')}")
                        # 商品名・内容量列をチェック
                        for col in df.columns:
                            if '商品名' in col:
                                print(f"      {col}: {row.get(col, 'N/A')}")
                            if '内容量' in col or '容量' in col:
                                print(f"      {col}: {row.get(col, 'N/A')}")
            else:
                print("  'ファイル名'列が見つかりません")
                
        except Exception as e:
            print(f"  読み込みエラー: {e}")
    else:
        print("\nall_collect.xlsx が存在しません")
    
    # 3. Phase1のファイル別パターンとの照合
    phase1_dir = f"G:\\共有ドライブ\\★OD\\99_商品管理\\DATA\\Phase1\\HARV\\{municipality}"
    file_pattern_file = os.path.join(phase1_dir, f"{municipality}_ファイル別パターン.xlsx")
    
    if os.path.exists(file_pattern_file):
        print(f"\nPhase1ファイル別パターンとの照合")
        
        try:
            df = pd.read_excel(file_pattern_file)
            pat0004_files = df[df['パターン名'] == 'PAT0004']
            
            print(f"  PAT0004ファイル数: {len(pat0004_files)}")
            print("  PAT0004ファイル一覧:")
            for _, row in pat0004_files.iterrows():
                print(f"    {row['ファイル名']}")
                
        except Exception as e:
            print(f"  読み込みエラー: {e}")

def check_phase4_to_phase5_flow():
    """Phase4→Phase5の処理フローを確認"""
    print("\n=== Phase4→Phase5フロー確認 ===")
    
    municipality = "2025-09-09_福岡県北九州市"
    base_dir = f"G:\\共有ドライブ\\★OD\\99_商品管理\\DATA\\Phase3\\HARV\\{municipality}"
    
    # Phase4で生成される_normalized.xlsxファイルの確認
    pat0004_original = os.path.join(base_dir, "PAT0004.xlsx")
    pat0004_normalized = os.path.join(base_dir, "PAT0004_normalized.xlsx")
    
    print(f"PAT0004.xlsx 存在: {os.path.exists(pat0004_original)}")
    print(f"PAT0004_normalized.xlsx 存在: {os.path.exists(pat0004_normalized)}")
    
    if os.path.exists(pat0004_normalized):
        # Phase4の正規化処理の結果を確認
        try:
            df = pd.read_excel(pat0004_normalized)
            print(f"正規化後の形状: {df.shape}")
            
            # ヘッダー行の確認（1行目）
            if len(df) > 0:
                headers = list(df.columns)
                print(f"正規化後のヘッダー: {headers[:10]}")  # 最初の10個
                
                # データの中身確認
                print("正規化後のサンプルデータ:")
                for i, (idx, row) in enumerate(df.head(3).iterrows()):
                    print(f"  行{i+1}: {list(row)[:5]}")  # 最初の5列
                    
        except Exception as e:
            print(f"正規化ファイル読み込みエラー: {e}")

def main():
    trace_phase5_processing()
    check_phase4_to_phase5_flow()

if __name__ == "__main__":
    main()
