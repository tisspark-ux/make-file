from money_log import MoneyLog


def main():
    log = MoneyLog()
    log.save("data/가계부.xlsx")
    print("가계부가 생성되었습니다: data/가계부.xlsx")


if __name__ == "__main__":
    main()
