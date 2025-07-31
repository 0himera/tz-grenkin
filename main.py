import datetime as dt
from pathlib import Path

import pandas as pd
from dateutil.relativedelta import relativedelta

# ---------- CONFIG ----------
ARENDA_FILE = Path("arenda.xlsx")
BANK_FILE   = Path("print 2.xlsx")
REPORT_FILE = Path(f"report_{dt.datetime.now():%Y-%m-%d}.xlsx")
# ----------------------------

def parse_arenda() -> pd.DataFrame:
    df = pd.read_excel(ARENDA_FILE, skiprows=1, header=None)
    # порядок колонок в файле: Гараж, Сумма, Первоначальная дата
    df.columns = ["Гараж", "Сумма", "Первоначальная_дата"]

    # явно задаём формат, чтобы не было предупреждений
    df["День_оплаты"] = pd.to_datetime(
        df["Первоначальная_дата"], format="%Y-%m-%d %H:%M:%S", errors="coerce"
    ).dt.day

    return df[["Гараж", "Сумма", "День_оплаты"]]

def _clean_amount(series: pd.Series) -> pd.Series:
    """Оставляет положительные значения в колонке сумму и преобразует к int."""
    return (
        series.dropna().astype(str)
        .loc[lambda x: x.str.startswith("+")]
        .str.replace("+", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
        .pipe(pd.to_numeric, errors="coerce")
        .dropna()
        .astype(int)
    )


def parse_bank() -> dict[int, dt.date]:
    """Возвращает словарь {сумма: последняя_дата_платежа}."""
    df = pd.read_excel(BANK_FILE, header=None)

    # в выписке дата операции в колонке A (index 0), сумма в колонке E (index 4)
    df = df[[0, 4]].rename(columns={0: "date_raw", 4: "amount_raw"})

    # очищаем суммы
    df["amount"] = _clean_amount(df["amount_raw"])
    df = df.dropna(subset=["amount"])

    # парсим дату: берём первые 10 символов вида 09.06.2025
    df["date"] = pd.to_datetime(df["date_raw"].astype(str).str.slice(0, 10), format="%d.%m.%Y", errors="coerce")
    df = df.dropna(subset=["date"])

    # оставляем только положительные платежи (по условию уже +)
    last_dates = (
        df.groupby("amount")["date"].max().dt.date.to_dict()
    )
    return last_dates


def expected_pay_date(day: int, month: int, year: int) -> dt.date:
    """
    Возвращает последний день месяца, если день выходит за его пределы.
    """
    try:
        return dt.date(year, month, day)
    except ValueError:
        # последний день месяца
        return dt.date(year, month, 1) + relativedelta(months=1) - dt.timedelta(days=1)

def build_report() -> None:
    arenda = parse_arenda()
    bank_payments = parse_bank()  # {amount: last_date}

    today = dt.date.today()
    rows = []

    for _, row in arenda.iterrows():
        garage = row["Гараж"]
        amount = int(row["Сумма"])
        day_pay = int(row["День_оплаты"])

        due_date = expected_pay_date(day_pay, today.month, today.year)
        grace_date = due_date + dt.timedelta(days=3)

        last_payment_date = bank_payments.get(amount)

        if last_payment_date:
            # платеж был
            if last_payment_date <= grace_date:
                status = "получен"
            elif last_payment_date <= today:
                status = "получен с задержкой"
            else:
                # 
                status = "неизвестно"
        else:
            
            if today <= grace_date:
                status = "срок не наступил"
            else:
                status = "просрочен"

        rows.append({
            "Гараж": garage,
            "Дата платежа (срок)": due_date.strftime("%d.%m.%Y"),
            "Сумма": amount,
            "Последний платёж": last_payment_date.strftime("%d.%m.%Y") if last_payment_date else "-",
            "Статус": status,
        })



    pd.DataFrame(rows).to_excel(REPORT_FILE, index=False)
    print(f"Отчёт сохранён: {REPORT_FILE}")

if __name__ == "__main__":
    build_report()