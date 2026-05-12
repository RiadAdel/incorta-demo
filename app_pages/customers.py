import streamlit as st

df = st.session_state.filtered_df

st.subheader("Top customers")

top = (
    df[["AccountNumber", "CustomerID", "City", "StateProvinceName",
        "Sales_Order_Count", "Average_Sales_Order", "Total_Sales"]]
    .sort_values("Total_Sales", ascending=False)
    .reset_index(drop=True)
    .head(100)
)

headers = ["#", "Account #", "Customer ID", "City", "State / Province",
           "Orders", "Avg Order Value", "Total Sales"]
right_cols = {5, 6, 7}
amber_cols = {7}

head_html = "".join(
    f'<th class="right">{h}</th>' if i in right_cols else f"<th>{h}</th>"
    for i, h in enumerate(headers)
)

rows_html = []
for rank, (_, row) in enumerate(top.iterrows(), start=1):
    cells = [
        ("", f"{rank}"),
        ("", str(row["AccountNumber"])),
        ("", str(row["CustomerID"])),
        ("", str(row["City"])),
        ("", str(row["StateProvinceName"])),
        ("right", f"{int(row['Sales_Order_Count']):,}"),
        ("right", f"${row['Average_Sales_Order']:,.0f}"),
        ("right amber", f"${row['Total_Sales']:,.0f}"),
    ]
    cell_html = "".join(
        f'<th class="{cls}">{val}</th>' if i == 0 else f'<td class="{cls}">{val}</td>'
        for i, (cls, val) in enumerate(cells)
    )
    rows_html.append(f"<tr>{cell_html}</tr>")

table_html = f"""
<div class="customer-table-wrap">
  <table class="customer-table">
    <thead><tr>{head_html}</tr></thead>
    <tbody>{''.join(rows_html)}</tbody>
  </table>
</div>
"""

st.markdown(table_html, unsafe_allow_html=True)
st.caption(f"Showing top 100 of {len(df):,} customers · ranked by total sales")
