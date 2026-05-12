import streamlit as st
import plotly.express as px

from utils import INCORTA_SCALE, apply_chart_style

df = st.session_state.filtered_df

left, right = st.columns(2)

with left.container(border=True, height="stretch"):
    st.subheader("Sales vs. order count")
    fig_scatter = px.scatter(
        df, x="Sales_Order_Count", y="Total_Sales",
        hover_data=["City", "StateProvinceName"],
        labels={"Sales_Order_Count": "Order Count", "Total_Sales": "Total Sales ($)"},
        color="Total_Sales", color_continuous_scale=INCORTA_SCALE,
    )
    apply_chart_style(fig_scatter)
    st.plotly_chart(fig_scatter, use_container_width=True)

with right.container(border=True, height="stretch"):
    st.subheader("Avg order value by state (top 10)")
    avg_df = (
        df.groupby("StateProvinceName", as_index=False)["Average_Sales_Order"]
        .mean().nlargest(10, "Average_Sales_Order")
        .rename(columns={"StateProvinceName": "State", "Average_Sales_Order": "Avg Order Value"})
    )
    fig_avg = px.bar(
        avg_df, x="Avg Order Value", y="State", orientation="h",
        labels={"Avg Order Value": "Avg Order Value ($)", "State": ""},
        color="Avg Order Value", color_continuous_scale=INCORTA_SCALE,
    )
    apply_chart_style(fig_avg, categorical_y=True)
    st.plotly_chart(fig_avg, use_container_width=True)
