import os
import argparse
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import pandas as pd
import yfinance as yf
import resend
import sys
print("使用Python:", sys.executable)
print("Python version:", sys.version)


# ============================================================
# 基本設定
# ============================================================

# このPythonファイルがあるフォルダ
BASE_DIR = Path(__file__).resolve().parent

# Excelファイル
EXCEL_FILE = BASE_DIR / "customers.xlsx"

# Excelのシート名
SHEET_NAME = "customers"

# Resend送信元
FROM_EMAIL = "KABU監視 <onboarding@resend.dev>"

# ============================================================
# ローカルテスト用
# ============================================================
# ★ここに自分のResend APIキーを入れる
RESEND_API_KEY = "RESEND_API"

# ★システム異常通知を受け取る自分のGmail
ADMIN_EMAIL = "yama01yu18@gmail.com"


# ============================================================
# GitHub Actionsに移行するときは
# 上の2行を削除して、こちらを使う
# ============================================================

# RESEND_API_KEY = os.environ["RESEND_API"]
# ADMIN_EMAIL = os.environ["ADMIN_EMAIL"]


# Resend APIキー設定
resend.api_key = RESEND_API_KEY

# 日本時間
JST = ZoneInfo("Asia/Tokyo")


# ============================================================
# メール送信
# ============================================================

def send_email(to_email, subject, body):

    try:

        resend.Emails.send({
            "from": FROM_EMAIL,
            "to": to_email,
            "subject": subject,
            "text": body
        })

        print(f"✅ メール送信成功: {to_email}")

        return True

    except Exception as e:

        print(f"❌ メール送信失敗: {to_email}")
        print(e)

        return False


# ============================================================
# 運営者への異常通知
# ============================================================

def send_admin_alert(subject, body):

    try:

        resend.Emails.send({
            "from": FROM_EMAIL,
            "to": ADMIN_EMAIL,
            "subject": f"🚨 KABU監視システム異常：{subject}",
            "text": body
        })

    except Exception as e:

        # Resend自体が止まっている可能性もあるので
        # ここでさらにメールを送ろうとはしない

        print("❌ 運営者への異常通知も失敗しました")
        print(e)


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
                "ticker": str,
                "purchase_date": str
            }
        )

        # 必須列
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

        # 足りない列を確認
        missing_columns = []

        for column in required_columns:

            if column not in df.columns:
                missing_columns.append(column)

        if missing_columns:

            raise ValueError(
                f"Excelに必要な列がありません: {missing_columns}"
            )

        print(f"Excel読込成功: {len(df)}行")

        return df

    except Exception as e:

        print()
        print("❌ Excel読み込みエラー")
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
# enabled判定
# ============================================================

def is_enabled(value):

    text = str(value).strip().lower()

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

def get_stock_data(ticker, purchase_date):

    stock = yf.Ticker(ticker)

    data = stock.history(
        start=purchase_date,
        auto_adjust=False
    )

    if data.empty:

        raise ValueError(
            "株価データを取得できませんでした"
        )

    if "Close" not in data.columns:

        raise ValueError(
            "Close列がありません"
        )

    close = data["Close"].dropna()

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
    # Excelから顧客情報取得
    # --------------------------------------------------------

    customer_id = row["customer_id"]
    email = row["email"]

    name = row["name"]
    ticker = row["ticker"]

   
    purchase_date = pd.to_datetime(
    row["purchase_date"]).strftime("%Y-%m-%d")

    average_price = float(row["average_price"])
    start_rate = float(row["start_rate"])
    drop_rate = float(row["drop_rate"])


    print()
    print("=" * 60)
    print(f"{customer_id} / {name} / {ticker}")
    print("=" * 60)


    try:

        # ====================================================
        # 株価取得
        # ====================================================

        price = get_stock_data(
            ticker,
            purchase_date
        )


        # ----------------------------------------------------
        # 最新終値
        # ----------------------------------------------------

        current_price = float(
            price.iloc[-1]
        )


        # ----------------------------------------------------
        # 購入日以降の最高終値
        # ----------------------------------------------------

        highest_price = float(
            price.max()
        )


        # ====================================================
        # 監視開始ライン計算
        # ====================================================

        # 例
        #
        # 平均取得価格 100円
        # start_rate 30
        #
        # 100 × 1.30
        #
        # → 130円

        start_price = (
            average_price
            * (1 + start_rate / 100)
        )


        # ====================================================
        # 現在の損益率
        # ====================================================

        profit_rate = (
            current_price / average_price - 1
        ) * 100


        # ====================================================
        # 現在状況表示
        # ====================================================

        print(f"購入日: {purchase_date}")

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
        # ★重要
        #
        # 過去に監視ラインを超えたかではなく、
        #
        # 「現在値」が平均取得価格から
        # 指定率以上上昇しているか
        #
        # で判定
        #
        # ====================================================

        if current_price < start_price:

            print()
            print("⚪ 監視対象外")

            remaining_rate = (
                start_price / current_price - 1
            ) * 100

            print(
                f"あと {remaining_rate:.1f}% "
                f"上昇すると監視対象"
            )

            return


        print()
        print("🟢 監視対象")


        # ====================================================
        # 通知ライン
        # ====================================================

        # 例
        #
        # 最高値200円
        # drop_rate 10
        #
        # 200 × 0.90
        #
        # → 180円

        alert_price = (
            highest_price
            * (1 - drop_rate / 100)
        )


        # ====================================================
        # 最高値からの実際の下落率
        # ====================================================

        actual_drop_rate = (
            1 - current_price / highest_price
        ) * 100


        print()
        print("【下落判定】")

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
        # メール通知判定
        # ====================================================

        if current_price <= alert_price:

            print()
            print("🚨 通知条件到達")


            subject = (
                f"【KABU監視】"
                f"{name} が設定条件に到達しました"
            )


            body = f"""
KABU監視システムからのお知らせです。


【登録銘柄】

{name}（{ticker})


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

実際の売却・保有については
ご自身でご判断ください。
"""


            success = send_email(
                email,
                subject,
                body
            )


            # ------------------------------------------------
            # 条件通知メールが送れなかった場合
            # 運営者へ通知
            # ------------------------------------------------

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
            print("通知条件には未到達")


            # 現在値から通知価格までの距離

            remaining_drop = (
                1 - alert_price / current_price
            ) * 100


            print(
                f"現在値からあと "
                f"{remaining_drop:.1f}%下落すると通知"
            )


    # ========================================================
    # 株価監視処理で異常があった場合
    # ========================================================

    except Exception as e:

        print()
        print(f"❌ 銘柄監視エラー: {e}")


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
# 日曜日の通信試験
# ============================================================

def send_weekly_health_check(df):

    today = datetime.now(JST)


    # --------------------------------------------------------
    # Python weekday
    #
    # 月 = 0
    # 火 = 1
    # 水 = 2
    # 木 = 3
    # 金 = 4
    # 土 = 5
    # 日 = 6
    # --------------------------------------------------------

    if today.weekday() != 6:

        return


    print()
    print("=" * 60)
    print("🟢 日曜日：通信試験")
    print("=" * 60)


    # ========================================================
    # enabledがTRUEの銘柄のみ
    # ========================================================

    active_df = df[
        df["enabled"].apply(is_enabled)
    ]


    # ========================================================
    # メールアドレスごとにまとめる
    #
    # 同じ人が5銘柄登録していても
    # メールは1通だけ
    # ========================================================

    for email, group in active_df.groupby("email"):


        stock_count = len(group)


        stock_names = "\n".join(
            f"・{row['name']}（{row['ticker']}）"
            for _, row in group.iterrows()
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


        # ----------------------------------------------------
        # 通信試験メールが送れなかった
        # ----------------------------------------------------

        if not success:

            send_admin_alert(
                "日曜通信試験メール送信失敗",
                f"""
日曜通信試験メールの送信に失敗しました。


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

def send_test_email(df, customer_id):


    # ========================================================
    # 指定customer_idだけ抽出
    # ========================================================

    rows = df[
        df["customer_id"]
        == str(customer_id)
    ]


    # ========================================================
    # enabledがTRUEのみ
    # ========================================================

    rows = rows[
        rows["enabled"].apply(is_enabled)
    ]


    # ========================================================
    # customer_idが存在しない
    # ========================================================

    if rows.empty:

        print()
        print(
            f"❌ customer_id={customer_id} "
            "が見つかりません"
        )

        return


    # ========================================================
    # メールアドレス
    # ========================================================

    email = rows.iloc[0]["email"]


    # ========================================================
    # 登録銘柄数
    # ========================================================

    stock_count = len(rows)


    # ========================================================
    # 登録銘柄一覧
    # ========================================================

    stock_names = "\n".join(
        f"・{row['name']}（{row['ticker']}）"
        for _, row in rows.iterrows()
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


また、毎週日曜日に
監視システムの稼働確認メールをお送りします。
"""


    success = send_email(
        email,
        subject,
        body
    )


    if success:

        print()
        print("✅ テストメール送信完了")


    else:

        print()
        print("❌ テストメール送信失敗")


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
# 通常の株価監視
# ============================================================

def run_monitor(df):


    now = datetime.now(JST)


    print()
    print("=" * 70)
    print("KABU監視システム")
    print(f"実行日時: {now}")
    print("=" * 70)


    # ========================================================
    # enabled TRUEだけ取得
    # ========================================================

    active_df = df[
        df["enabled"].apply(is_enabled)
    ]


    print()
    print(
        f"有効監視銘柄数: "
        f"{len(active_df)}"
    )


    # ========================================================
    # 各銘柄チェック
    # ========================================================

    for _, row in active_df.iterrows():

        check_stock(row)


    # ========================================================
    # 日曜日なら通信試験
    # ========================================================

    send_weekly_health_check(df)


    print()
    print("=" * 70)
    print("KABU監視終了")
    print("=" * 70)


# ============================================================
# プログラム開始
# ============================================================

if __name__ == "__main__":


    # ========================================================
    # コマンドライン引数
    # ========================================================

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
    # Excel読み込み
    # ========================================================

    df = load_customers()


    # ========================================================
    # テストメールモード
    # ========================================================

    if args.test_customer:

        send_test_email(
            df,
            args.test_customer
        )


    # ========================================================
    # 通常監視モード
    # ========================================================

    else:

        run_monitor(df)