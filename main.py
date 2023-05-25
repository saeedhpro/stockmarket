import pandas as pd


def run():
    pe_list = pd.read_excel("./pe شرکت ها.xlsx", skiprows=0)
    pe_list = pe_list[pe_list.columns.intersection(["P/E", "نام شرکت"])]
    pe_list = pe_list.dropna(subset=["P/E"])
    pe_list = pe_list[pe_list["P/E"] < 9.94166666666667]

    print("pe done")

    mali_list = pd.read_excel("./نسبت‌های مالی-مجمع-29-12-1401 .xlsx", skiprows=0)
    mali_list = mali_list[mali_list.columns.intersection(["ردیف", "نسبت بدهی به ارزش ویژه", "نسبت جاری", "نام شرکت"])]
    mali_list = mali_list.dropna(subset=["ردیف"])
    mali_list = mali_list[mali_list["نام شرکت"].isin(pe_list["نام شرکت"])]
    mali_list = mali_list[mali_list["نسبت بدهی به ارزش ویژه"] < 1]
    mali_list = mali_list[mali_list["نسبت جاری"] > 1]

    print("mali done")

    taraz_list = pd.read_excel("./ترازنامه-مجمع-29-12-1401.xlsx", skiprows=0)
    taraz_list = taraz_list[taraz_list.columns.intersection(["ردیف", "جمع کل دارایی‌ها", "جمع کل بدهی‌ها", "شرکت"])]
    taraz_list = taraz_list.dropna(subset=["ردیف"])
    taraz_list = taraz_list[taraz_list["شرکت"].isin(mali_list["نام شرکت"])]
    taraz_list = taraz_list[(taraz_list["جمع کل بدهی‌ها"] < 2 * (taraz_list["جمع کل دارایی‌ها"] - taraz_list["جمع کل بدهی‌ها"]))]

    print("taraz done")

    name_list = taraz_list["شرکت"]

    name_list.to_csv("./result.csv", sep='\t', encoding='utf-8')

    print("see result name in result.csv")


if __name__ == '__main__':
    run()
