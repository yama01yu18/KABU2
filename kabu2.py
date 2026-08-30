import os
import argparse
import smtplib

from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
from email.mime.text import MIMEText

import pandas as pd
import yfinance as yf


# ============================================================
# 基本設定
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILE = BASE_DIR / "customers.xlsx"

SHEET_NAME = "customers"

JST = ZoneInfo("Asia/Tokyo")


# ============================================================
# Gmail設定
# ============================================================

# GitHub ActionsではSecretsから取得
#
# GMAIL_ADDRESS
# GMAIL_APP_PASSWORD
# ADMIN_EMAIL

GMAIL_ADDRESS = os.environ["GMAIL_ADDRESS"]

GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]

ADMIN_EMAIL = os.environ["ADMIN_EMAIL"]


# ============================================================
# メール送信
# ============================================================

def send_email(to_email, subject, body):

    try:

        msg = MIMEText(
            body,
            "plain",
            "utf-8"
        )

        msg["Subject"] = subject

        msg["From"] = (
            f"KABU監視 <{GMAIL_ADDRESS}>"
        )

        msg["To"] = to_email


        with smtplib.SMTP_SSL(
            "smtp.gmail.com",
            465
        ) as server:

            server.login(
                GMAIL_ADDRESS,
                GMAIL_APP_PASSWORD
            )

            server.send_message(msg)


        print(
            f"✅ メール送信成功: "
            f"{to_email}"
        )

        return True


    except Exception as e:

        print(
            f"❌ メール送信失敗: "
            f"{to_email}"
        )

        print(e)

        return False


# ============================================================
# 運営者への異常通知
# ============================================================

def send_admin_alert(subject, body):

    success = send_email(
        ADMIN_EMAIL,
        f"🚨 KABU監視システム異常：{subject}",
        body
    )

    if not success:

        print(
            "❌ 運営者への"
            "異常通知も失敗しました"
        )


# ============================================================
# Excel読み込み
# ============================================================

def load_customers():

    try:

        print()
        print("Excelファイル:")
        print(EXCEL_FILE)


        df = pd.read_excel(
            EXCEL_FILE,
            sheet_name=SHEET_NAME,
            dtype={
                "customer_id": str,
                "email": str,
                "name": str,
                "ticker": str
            }
        )


        required_columns = [
            "customer_id",
            "email",
            "name",
            "ticker",
            "purchase_date",
            "average_price",
            "start_rate",
            "drop_rate",
            "enabled"
        ]


        missing_columns = []

        for column in required_columns:

            if column not in df.columns:

                missing_columns.append(
                    column
                )


        if missing_columns:

            raise ValueError(
                "Excelに必要な列がありません: "
                f"{missing_columns}"
            )


        print(
            f"Excel読込成功: "
            f"{len(df)}行"
        )


        return df


    except Exception as e:

        print()
        print(
            "❌ Excel読み込みエラー"
        )

        print(e)


        send_admin_alert(
            "Excel読み込みエラー",
            f"""
customers.xlsx の読み込みに失敗しました。


Excelファイル:

{EXCEL_FILE}


シート名:

{SHEET_NAME}


エラー:

{e}


実行日時:

{datetime.now(JST)}
"""
        )

        raise


# ============================================================
# enabled 判定
# ============================================================

def is_enabled(value):

    text = str(
        value
    ).strip().lower()


    return text in [
        "true",
        "1",
        "yes",
        "y",
        "on"
    ]


# ============================================================
# 株価取得
# ============================================================

def get_stock_data(
    ticker,
    purchase_date
):

    stock = yf.Ticker(
        ticker
    )


    data = stock.history(
        start=purchase_date,
        auto_adjust=False
    )


    if data.empty:

        raise ValueError(
            "株価データを"
            "取得できませんでした"
        )


    if "Close" not in data.columns:

        raise ValueError(
            "Close列がありません"
        )


    close = data[
        "Close"
    ].dropna()


    if close.empty:

        raise ValueError(
            "終値データがありません"
        )


    return close


# ============================================================
# 1銘柄の監視
# ============================================================

def check_stock(row):


    # --------------------------------------------------------
    # 顧客情報
    # --------------------------------------------------------

    customer_id = row[
        "customer_id"
    ]

    email = row[
        "email"
    ]

    name = row[
        "name"
    ]

    ticker = row[
        "ticker"
    ]


    # Excelの日付を
    # YYYY-MM-DDに統一
    purchase_date = (
        pd.to_datetime(
            row["purchase_date"]
        )
        .strftime("%Y-%m-%d")
    )


    average_price = float(
        row["average_price"]
    )

    start_rate = float(
        row["start_rate"]
    )

    drop_rate = float(
        row["drop_rate"]
    )


    print()
    print(
        "=" * 60
    )

    print(
        f"{customer_id} / "
        f"{name} / "
        f"{ticker}"
    )

    print(
        "=" * 60
    )


    try:

        # ====================================================
        # 株価取得
        # ====================================================

        price = get_stock_data(
            ticker,
            purchase_date
        )


        # ====================================================
        # 最新終値
        # ====================================================

        current_price = float(
            price.iloc[-1]
        )


        # ====================================================
        # 購入日以降の最高終値
        # ====================================================

        highest_price = float(
            price.max()
        )


        # ====================================================
        # 監視開始ライン
        # ====================================================
        #
        # average_price = 100
        # start_rate = 30
        #
        # → 130円
        #
        # ====================================================

        start_price = (
            average_price
            * (
                1
                + start_rate / 100
            )
        )


        # ====================================================
        # 現在の損益率
        # ====================================================

        profit_rate = (
            current_price
            / average_price
            - 1
        ) * 100


        # ====================================================
        # 表示
        # ====================================================

        print(
            f"購入日: "
            f"{purchase_date}"
        )

        print(
            f"平均取得価格: "
            f"{average_price:.2f}"
        )

        print(
            f"最新終値: "
            f"{current_price:.2f}"
        )

        print(
            f"損益率: "
            f"{profit_rate:+.2f}%"
        )

        print(
            f"監視開始ライン: "
            f"{start_price:.2f}"
        )

        print(
            f"購入日以降最高終値: "
            f"{highest_price:.2f}"
        )


        # ====================================================
        # 監視対象判定
        # ====================================================
        #
        # 現在値が、
        # 平均取得価格から
        # 指定率以上上昇しているか
        #
        # ====================================================

        if current_price < start_price:

            print()
            print(
                "⚪ 監視対象外"
            )


            remaining_rate = (
                start_price
                / current_price
                - 1
            ) * 100


            print(
                f"あと "
                f"{remaining_rate:.1f}% "
                f"上昇すると監視対象"
            )


            return


        print()
        print(
            "🟢 監視対象"
        )


        # ====================================================
        # 通知ライン
        # ====================================================
        #
        # highest_price = 200
        # drop_rate = 10
        #
        # → 180円
        #
        # ====================================================

        alert_price = (
            highest_price
            * (
                1
                - drop_rate / 100
            )
        )


        # ====================================================
        # 最高値からの下落率
        # ====================================================

        actual_drop_rate = (
            1
            - current_price
            / highest_price
        ) * 100


        print()
        print(
            "【下落判定】"
        )

        print(
            f"最高終値: "
            f"{highest_price:.2f}"
        )

        print(
            f"通知ライン: "
            f"{alert_price:.2f}"
        )

        print(
            f"最高値からの下落率: "
            f"{actual_drop_rate:.2f}%"
        )


        # ====================================================
        # 通知条件
        # ====================================================

        if current_price <= alert_price:

            print()
            print(
                "🚨 通知条件到達"
            )


            subject = (
                "【KABU監視】"
                f"{name} が"
                "設定条件に到達しました"
            )


            body = f"""
KABU監視システムからのお知らせです。


【登録銘柄】

{name}（{ticker}）


【購入情報】

購入日：
{purchase_date}

平均取得価格：
{average_price:,.2f}


【現在の状況】

最新終値：
{current_price:,.2f}

平均取得価格からの損益率：
{profit_rate:+.1f}%


【監視条件】

監視開始条件：
平均取得価格から +{start_rate:.1f}%

購入日以降の最高終値：
{highest_price:,.2f}

最高値からの現在の下落率：
-{actual_drop_rate:.1f}%

設定された通知条件：
最高値から -{drop_rate:.1f}%


設定された条件に到達しています。


本通知は売却を推奨するものではありません。

実際の売却・保有については、
ご自身でご判断ください。
"""


            success = send_email(
                email,
                subject,
                body
            )


            # =================================================
            # 通知失敗
            # =================================================

            if not success:

                send_admin_alert(
                    "条件到達メール送信失敗",
                    f"""
条件到達メールの送信に失敗しました。


顧客ID:

{customer_id}


メール:

{email}


銘柄:

{name}


Ticker:

{ticker}


最新終値:

{current_price}


最高終値:

{highest_price}


実行日時:

{datetime.now(JST)}
"""
                )


        else:

            print()
            print(
                "通知条件には未到達"
            )


            remaining_drop = (
                1
                - alert_price
                / current_price
            ) * 100


            print(
                f"現在値からあと "
                f"{remaining_drop:.1f}% "
                f"下落すると通知"
            )


    # ========================================================
    # 監視処理異常
    # ========================================================

    except Exception as e:

        print()
        print(
            f"❌ 銘柄監視エラー: "
            f"{e}"
        )


        send_admin_alert(
            f"{ticker} 株価監視エラー",
            f"""
株価監視処理でエラーが発生しました。


顧客ID:

{customer_id}


メール:

{email}


銘柄:

{name}


Ticker:

{ticker}


購入日:

{purchase_date}


エラー:

{e}


実行日時:

{datetime.now(JST)}
"""
        )


# ============================================================
# 通信試験
# ============================================================

def send_weekly_health_check(df):

    today = datetime.now(
        JST
    )


    active_df = df[
        df["enabled"].apply(
            is_enabled
        )
    ]


    print()
    print(
        "=" * 60
    )

    print(
        "🟢 通信試験"
    )

    print(
        "=" * 60
    )


    # ========================================================
    # メールアドレスごとに1通
    # ========================================================

    for email, group in (
        active_df.groupby(
            "email"
        )
    ):

        stock_count = len(
            group
        )


        stock_names = "\n".join(
            f"・{row['name']} "
            f"（{row['ticker']}）"
            for _, row
            in group.iterrows()
        )


        subject = (
            "🟢【KABU監視】"
            "監視システム稼働確認"
        )


        body = f"""
KABU監視システムは正常に稼働しています。


確認日時：

{today.strftime("%Y/%m/%d %H:%M")}


現在の登録銘柄数：

{stock_count}銘柄


登録銘柄：

{stock_names}


設定条件に到達した場合は、
別途メールでお知らせします。
"""


        success = send_email(
            email,
            subject,
            body
        )


        if not success:

            send_admin_alert(
                "通信試験メール送信失敗",
                f"""
通信試験メールの送信に失敗しました。


送信先:

{email}


登録銘柄数:

{stock_count}


実行日時:

{today}
"""
            )


# ============================================================
# 初回テストメール
# ============================================================

def send_test_email(
    df,
    customer_id
):


    # ========================================================
    # 指定顧客だけ抽出
    # ========================================================

    rows = df[
        df["customer_id"]
        == str(customer_id)
    ]


    # enabled TRUEのみ
    rows = rows[
        rows["enabled"].apply(
            is_enabled
        )
    ]


    if rows.empty:

        print()
        print(
            f"❌ customer_id="
            f"{customer_id} "
            "が見つかりません"
        )

        return


    # ========================================================
    # 同一customer_idに
    # 複数メールアドレスがないか確認
    # ========================================================

    emails = (
        rows["email"]
        .dropna()
        .unique()
    )


    if len(emails) != 1:

        print()
        print(
            "❌ 同じcustomer_idに"
            "複数のメールアドレスがあります"
        )

        print(
            emails
        )

        return


    email = emails[0]


    stock_count = len(
        rows
    )


    stock_names = "\n".join(
        f"・{row['name']} "
        f"（{row['ticker']}）"
        for _, row
        in rows.iterrows()
    )


    subject = (
        "【KABU監視】"
        "登録テストメール"
    )


    body = f"""
KABU監視システムへの登録テストメールです。


このメールを受信できていれば、
通知先メールアドレスは正常に登録されています。


登録銘柄数：

{stock_count}銘柄


登録銘柄：

{stock_names}


今後、設定された条件に到達した場合に
メールでお知らせします。


また、定期的に
監視システムの稼働確認メールをお送りします。
"""


    success = send_email(
        email,
        subject,
        body
    )


    if success:

        print()
        print(
            "✅ テストメール送信完了"
        )


    else:

        print()
        print(
            "❌ テストメール送信失敗"
        )


        send_admin_alert(
            "テストメール送信失敗",
            f"""
テストメールの送信に失敗しました。


customer_id:

{customer_id}


送信先:

{email}


実行日時:

{datetime.now(JST)}
"""
        )


# ============================================================
# 通常監視
# ============================================================

def run_monitor(df):

    now = datetime.now(
        JST
    )


    print()
    print(
        "=" * 70
    )

    print(
        "KABU監視システム"
    )

    print(
        f"実行日時: {now}"
    )

    print(
        "=" * 70
    )


    active_df = df[
        df["enabled"].apply(
            is_enabled
        )
    ]


    print()
    print(
        f"有効監視銘柄数: "
        f"{len(active_df)}"
    )


    # ========================================================
    # 各銘柄監視
    # ========================================================

    for _, row in (
        active_df.iterrows()
    ):

        check_stock(
            row
        )


    # ========================================================
    # 月曜日だけ通信試験
    #
    # Python weekday
    #
    # 月 = 0
    # 火 = 1
    # 水 = 2
    # 木 = 3
    # 金 = 4
    # 土 = 5
    # 日 = 6
    # ========================================================

    if now.weekday() == 0:

        send_weekly_health_check(
            df
        )


    print()
    print(
        "=" * 70
    )

    print(
        "KABU監視終了"
    )

    print(
        "=" * 70
    )


# ============================================================
# プログラム開始
# ============================================================

if __name__ == "__main__":


    parser = argparse.ArgumentParser()


    parser.add_argument(
        "--test-customer",
        type=str,
        help=(
            "指定したcustomer_idへ "
            "登録テストメールを送信"
        )
    )


    args = parser.parse_args()


    # ========================================================
    # Excel読込
    # ========================================================

    df = load_customers()


    # ========================================================
    # テストメール
    # ========================================================

    if args.test_customer:

        send_test_email(
            df,
            args.test_customer
        )


    # ========================================================
    # 通常監視
    # ========================================================

    else:

        run_monitor(
            df
        )