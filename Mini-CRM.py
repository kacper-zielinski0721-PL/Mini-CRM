import os
import pandas as pd

BASE_PATH = os.path.dirname(__file__)
INPUT_FILE = os.path.join(BASE_PATH, "dane klientów Mini-CRM.xlsx")
OUTPUT_FILE = os.path.join(BASE_PATH, "crm_report.xlsx")

def load_data(path):
    df = pd.read_excel(path, engine="openpyxl")

    if len(df.columns) == 1:
        df = df.iloc[:, 0].astype(str).str.split(expand=True)
        df.columns = ["client", "amount", "city"]

    df.columns = df.columns.str.strip().str.lower()

    df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    df = df.dropna(subset=["amount"])

    return df

def add_categoty(df):
    def get_category(amount):
        if amount > 1500:
            return "Enterprise"
        elif amount > 850:
            return "Mid-market"
        else:
            return "SMB"

    df["category"] = df["amount"].apply(get_category)
    return df


def calculate_matrics(df):
    total = df["amount"].sum()
    average = df["amount"].mean()
    max_val = df["amount"].max()

    top_clients = df.groupby("client")["amount"].sum().sort_values(ascending=False)

    top_client = top_clients.idxmax()
    top_client_sum = top_clients.max()

    return{
        "total": total,
        "average": average,
        "max": max_val,
        "top_client": top_client,
        "top_client_sum": top_client_sum,
        "top_clients_table": top_clients.reset_index()

        }

def build_summary(metrics):
    return pd.DataFrame([
        ["Total revenue", metrics["total"]],
        ["Average deal", metrics["average"]],
        ["Max deal", metrics["max"]],
        ["Top client", metrics["top_client"]],
        ["top client revenue", metrics["top_client_sum"]],
    ], columns=["Metric", "Value"])

def build_city_report(df):
    return df.groupby("city")["amount"].sum().reset_index().sort_values(by="amount", ascending=False)


def export_report(df, metrics):
    summary_df = build_summary(metrics)
    city_df = build_city_report(df)
    top_clients_df = metrics["top_clients_table"]

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        df.sort_values("amount", ascending=False).to_excel(
        writer, sheet_name="Report", index=False
        )

    summary_df.to_excel(writer, sheet_name="Summary", index=False)

    top_clients_df.to_excel(writer, sheet_name="Top client", index=False)

    city_df.to_excel(writer, sheet_name="By City", index=False)

    workbook = writer.book

    for sheet in ["Report", "Summary", "Top Clients", "By City"]:
        worksheet = writer.sheets["Report"]
        worksheet.set_column("A:D", 20)

print(f" Report saved: {OUTPUT_FILE}")

def main():
    df =load_data(INPUT_FILE)
    df = add_categoty(df)

    metrics = calculate_matrics(df)

    export_report(df, metrics)

    print("\n KPI:")
    print("Total:", metrics["total"])
    print("Average:", round(metrics["average"], 2))
    print("Top client:", metrics["top_client"])

if __name__ == "__main__":
    main()
