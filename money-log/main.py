import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from money_log import create


def main():
    parser = argparse.ArgumentParser(description="가계부 Excel 생성기")
    parser.add_argument("--year", type=int, default=2026, help="가계부 연도 (기본값: 2026)")
    parser.add_argument("--output", type=str, default=None, help="출력 파일 경로")
    args = parser.parse_args()

    os.makedirs("data", exist_ok=True)
    path = create(year=args.year, output_path=args.output)
    print(f"가계부 생성 완료: {path}")


if __name__ == "__main__":
    main()
