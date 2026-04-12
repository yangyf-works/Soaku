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
        default="diff.json",
        help="出力JSONファイルパス（デフォルト: diff.json）"
    )

    # 類似度閾値
    parser.add_argument(
        "-t", "--threshold",
        type=float,
        default=0.5,
        help="行マッチの類似度閾値（0.0〜1.0）"
    )

    # シート指定
    parser.add_argument(
        "-s", "--sheet",
        help="対象シート名（未指定なら全シート）"
    )

    # 詳細ログ
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="詳細ログを表示"
    )

    # ドライラン（実際には出力しない）
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="処理内容だけ表示（ファイル出力しない）"
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
    args = parse_args()
    validate_args(args)

    if args.verbose:
        print("=== 入力パラメータ ===")
        print(f"Before: {args.bookPath1}")
        print(f"After : {args.bookPath2}")
        print(f"DataOnly: {args.data}")
        print(f"DryRun: {args.dry_run}")
        print(f"Output: {args.output}")
        print(f"Threshold: {args.threshold}")
        print(f"Sheet: {args.sheet}")
        print(f"DryRun: {args.dry_run}")

    data = diff.diff(path1=args.bookPath1, path2=args.bookPath2, dataonly=args.data, threshold=args.threshold)

    if args.dry_run:
        print("DryRunのため出力は行いません")
        return

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    else:
        json.dump(data, sys.stdout, indent=2, ensure_ascii=False)

if __name__ == "__main__":
    main()