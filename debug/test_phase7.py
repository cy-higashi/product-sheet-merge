import os
import shutil

def main():
    # 環境変数からフォルダ名（MUNICIPALITY_NAME）を取得
    municipality = os.environ.get("MUNICIPALITY_NAME")
    if not municipality:
        print("MUNICIPALITY_NAME 環境変数が設定されていません。")
        return

    # Phase3 配下の対象ファイルパスを構築
    source_dir = os.path.join(
        r'G:\共有ドライブ\★OD\99_商品管理\DATA\Phase3\HARV',
        municipality
    )
    source_file = os.path.join(source_dir, "all_collect.xlsx")
    
    if not os.path.exists(source_file):
        print(f"ソースファイルが存在しません: {source_file}")
        return

    # 複製先のディレクトリパス
    dest_dir = r"G:\共有ドライブ\★OD\99_商品管理\DATA\Phase4\HARV"
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir, exist_ok=True)
    
    # 複製後のファイル名を環境変数の値（フォルダ名）に変更
    dest_file = os.path.join(dest_dir, f"{municipality}.xlsx")
    
    try:
        shutil.copy(source_file, dest_file)
        print(f"複製完了: {dest_file}")
    except Exception as e:
        print(f"複製に失敗しました → {e}")

if __name__ == "__main__":
    main()
