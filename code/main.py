import argparse
import sys
import os
import json
import diff

def parse_args():
    parser = argparse.ArgumentParser(
        prog="excel-diff",
        description="Excelファイルの差分を比較するツール"
    )

    # ======================
    # 必須引数
    # ======================
    parser.add_argument(
        "bookPath1",
        help="比較元（Before）のExcelファイル"
    )

    parser.add_argument(
        "bookPath2",
        help="比較先（After）のExcelファイル"
    )

    # ======================
    # オプション
    # ======================

    # 出力ファイル
    parser.add_argument(
        "-o", "--output",
        help="出力JSONファイルパス（デフォルトは標準出力）"
    )

    # 類似度閾値
    parser.add_argument(
        "-t", "--threshold",
        type=float,
        default=0.5,
        help="行マッチの類似度閾値（0.0〜1.0）"
    )

    parser.add_argument(
        "-d", "--data",
        action="store_true",
        help="数式ではなく計算結果（data_only=True）で比較する"
    )

    return parser.parse_args()

def validate_args(args):
    # ファイル存在チェック
    if not os.path.exists(args.bookPath1):
        print(f"エラー: ファイルが存在しません: {args.bookPath1}")
        sys.exit(1)

    if not os.path.exists(args.bookPath2):
        print(f"エラー: ファイルが存在しません: {args.bookPath2}")
        sys.exit(1)

    # thresholdチェック
    if not (0.0 <= args.threshold <= 1.0):
        print("エラー: thresholdは0.0〜1.0で指定してください")
        sys.exit(1)

def main():
    sys.stdout.reconfigure(encoding="utf-8")
    args = parse_args()
    validate_args(args)

    data = diff.exceldiff(path1=args.bookPath1, path2=args.bookPath2, dataonly=args.data, threshold=args.threshold)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    else:
        json.dump(data, sys.stdout, indent=2, ensure_ascii=False)

if __name__ == "__main__":
    main()