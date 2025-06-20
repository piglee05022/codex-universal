import sys

from .filing import create_filing
from .example import sample_case

MENU = """\
Legal Document Generator
=======================
1. 產生範例起訴狀
2. 離開
"""


def main() -> None:
    print(MENU)
    choice = input("請輸入選項: ").strip()
    if choice == "1":
        create_filing(sample_case, "法律文書_起訴狀.docx")
        print("產生完成: 法律文書_起訴狀.docx")
    else:
        print("Bye")


if __name__ == "__main__":
    main()
