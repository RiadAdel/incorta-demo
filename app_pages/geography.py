import streamlit as st
import plotly.express as px

from utils import INCORTA_SCALE, US_STATE_ABBREV, apply_chart_style

df = st.session_state.filtered_df

state_df = (
    df.groupby("StateProvinceName", as_index=False)["Total_Sales"]
    .sum()
    .rename(columns={"StateProvinceName": "State", "Total_Sales": "Total Sales"})
)
state_df["code"] = state_df["State"].map(US_STATE_ABBREV)
map_df = state_df.dropna(subset=["code"])

if not map_df.empty:
    st.subheader("Sales by state")
    fig_map = px.choropleth(
        map_df,
        locations="code",
        locationmode="USA-states",
        color="Total Sales",
        scope="usa",
        hover_name="State",
        color_continuous_scale=INCORTA_SCALE,
        labels={"Total Sales": "Total Sales ($)"},
    )
    fig_map.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        geo={
            "bgcolor": "rgba(0,0,0,0)",
            "lakecolor":    "rgba(0,0,0,0)",
            "landcolor":    "rgba(36,28,85,0.4)",
            "subunitcolor": "rgba(139,92,246,0.3)",
        },
        font={"color": "#e8e6f0"},
        coloraxis_showscale=True,
        margin={"r": 0, "t": 0, "l": 0, "b": 0},
    )
    st.plotly_chart(fig_map, use_container_width=True)

left, right = st.columns(2)

with left.container(border=True, height="stretch"):
    st.subheader("Top 10 cities by sales")
    city_df = (
        df.groupby("City", as_index=False)["Total_Sales"]
        .sum().nlargest(10, "Total_Sales")
    )
    fig_city = px.bar(
        city_df, x="Total_Sales", y="City", orientation="h",
        labels={"Total_Sales": "Total Sales ($)", "City": ""},
        color="Total_Sales", color_continuous_scale=INCORTA_SCALE,
    )
    apply_chart_style(fig_city, categorical_y=True)
    st.plotly_chart(fig_city, use_container_width=True)

with right.container(border=True, height="stretch"):
    st.subheader("Top 10 states by total sales")
    state_sales_df = (
        df.groupby("StateProvinceName", as_index=False)["Total_Sales"]
        .sum().nlargest(10, "Total_Sales")
        .rename(columns={"StateProvinceName": "State", "Total_Sales": "Total Sales"})
    )
    fig_state = px.bar(
        state_sales_df, x="Total Sales", y="State", orientation="h",
        labels={"Total Sales": "Total Sales ($)", "State": ""},
        color="Total Sales", color_continuous_scale=INCORTA_SCALE,
    )
    apply_chart_style(fig_state, categorical_y=True)
    st.plotly_chart(fig_state, use_container_width=True)
