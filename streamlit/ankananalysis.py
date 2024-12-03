import streamlit as st
import pandas as pd
import openai
import openpyxl

# Set Streamlit page configuration
st.set_page_config(page_title="Fashion Sales Dashboard", page_icon="🛍️", layout="wide")

# Load data
data = pd.read_excel('data_with_classified_categories.xlsx')

# Ensure Quantity_Sold and Quantity_In_Stock columns are numeric
data['Quantity_Sold'] = pd.to_numeric(data['Quantity_Sold'], errors='coerce').fillna(0)
data['Quantity_In_Stock'] = pd.to_numeric(data['Quantity_In_Stock'], errors='coerce').fillna(0)

# Initialize OpenAI API
openai.api_key = "API KEY NEEDED"  # Replace with your OpenAI API key

# Sidebar setup for selecting filters and styling
st.sidebar.title("🏮 Filter Options")
page = st.sidebar.selectbox("Select Page", ["Sales Dashboard", "Executive Insights and Recommendations"])
marketing_group = st.sidebar.selectbox("Select Marketing Group", data['Marketing_Group'].unique())
classified_category = st.sidebar.selectbox("Select Classified Category", data['Classified_Category'].unique())

# Filter data by the selected marketing group and classified category
filtered_data = data[(data['Marketing_Group'] == marketing_group) & (data['Classified_Category'] == classified_category)]

# Month selection slider
months_available = filtered_data['Month_Name'].unique()
selected_months = st.sidebar.multiselect("Select Months", options=months_available, default=months_available)
filtered_data = filtered_data[filtered_data['Month_Name'].isin(selected_months)]

# Category filter
categories_available = filtered_data['Category'].unique()
selected_categories = st.sidebar.multiselect("Select Categories", options=categories_available, default=categories_available)
filtered_data = filtered_data[filtered_data['Category'].isin(selected_categories)]

# Sales Dashboard Page
if page == "Sales Dashboard":
    st.markdown("<h1 style='text-align: center;'>Sales Dashboard</h1>", unsafe_allow_html=True)

    # Sales Trend Analysis (Monthly Sales Summary by Brand/Category)
    st.markdown("<div style='border: 1px solid #cccccc; border-radius: 10px; padding: 15px; margin-bottom: 20px;'>", unsafe_allow_html=True)
    st.markdown("<h2>Sales Trend Analysis by Month</h2>", unsafe_allow_html=True)
    monthly_sales = filtered_data.groupby(['Month_Name', 'Brand'])['Quantity_Sold'].sum().reset_index()
    if not monthly_sales.empty:
        st.markdown("<h3>Sales Trend Graph</h3>", unsafe_allow_html=True)
        monthly_sales_pivot = monthly_sales.pivot(index='Month_Name', columns='Brand', values='Quantity_Sold').fillna(0)
        st.line_chart(monthly_sales_pivot)
        st.dataframe(monthly_sales)
    else:
        st.write("No data available for the selected filters.")
    st.markdown("</div>", unsafe_allow_html=True)

    # Top 10 Brands by Categories
    st.markdown("<div style='border: 1px solid #cccccc; border-radius: 10px; padding: 15px; margin-bottom: 20px;'>", unsafe_allow_html=True)
    st.markdown("<h2>Top 10 Brands by Category</h2>", unsafe_allow_html=True)
    top_10_brands = filtered_data.groupby('Brand')['Quantity_Sold'].sum().nlargest(10).reset_index()
    if not top_10_brands.empty:
        st.dataframe(top_10_brands)
    else:
        st.write("No data available for the selected filters.")
    st.markdown("</div>", unsafe_allow_html=True)

    # Bottom 10 Brands by Categories
    st.markdown("<div style='border: 1px solid #cccccc; border-radius: 10px; padding: 15px; margin-bottom: 20px;'>", unsafe_allow_html=True)
    st.markdown("<h2>Bottom 10 Brands by Category</h2>", unsafe_allow_html=True)
    bottom_10_brands = filtered_data.groupby('Brand')['Quantity_Sold'].sum().nsmallest(10).reset_index()
    if not bottom_10_brands.empty:
        st.dataframe(bottom_10_brands)
    else:
        st.write("No data available for the selected filters.")
    st.markdown("</div>", unsafe_allow_html=True)

    # Inventory Turnover Analysis
    st.markdown("<div style='border: 1px solid #cccccc; border-radius: 10px; padding: 15px; margin-bottom: 20px;'>", unsafe_allow_html=True)
    st.markdown("<h2>Inventory Turnover Analysis</h2>", unsafe_allow_html=True)
    filtered_data['Average_Inventory'] = (filtered_data['Quantity_Sold'] + filtered_data['Quantity_In_Stock']) / 2
    filtered_data['Inventory_Turnover'] = filtered_data['Quantity_Sold'] / filtered_data['Average_Inventory']
    inventory_turnover = filtered_data[['Brand', 'Category', 'Inventory_Turnover']].dropna().reset_index(drop=True)
    if not inventory_turnover.empty:
        st.dataframe(inventory_turnover)
    else:
        st.write("No data available for the selected filters.")
    st.markdown("</div>", unsafe_allow_html=True)

    # Sell-Through Rate Analysis
    st.markdown("<div style='border: 1px solid #cccccc; border-radius: 10px; padding: 15px; margin-bottom: 20px;'>", unsafe_allow_html=True)
    st.markdown("<h2>Sell-Through Rate Analysis</h2>", unsafe_allow_html=True)
    filtered_data['Sell_Through_Rate'] = (filtered_data['Quantity_Sold'] / (filtered_data['Quantity_Sold'] + filtered_data['Quantity_In_Stock'])) * 100
    sell_through_rate = filtered_data[['Brand', 'Category', 'Sell_Through_Rate']].dropna().reset_index(drop=True)
    if not sell_through_rate.empty:
        st.markdown("<h3>Sell-Through Rate Graph</h3>", unsafe_allow_html=True)
        sell_through_rate_pivot = sell_through_rate.pivot_table(index='Brand', columns='Category', values='Sell_Through_Rate', fill_value=0)
        st.bar_chart(sell_through_rate_pivot)
        st.dataframe(sell_through_rate)
    else:
        st.write("No data available for the selected filters.")
    st.markdown("</div>", unsafe_allow_html=True)

    # Stock-Out Alerts
    st.markdown("<div style='border: 1px solid #cccccc; border-radius: 10px; padding: 15px; margin-bottom: 20px;'>", unsafe_allow_html=True)
    st.markdown("<h2>Stock-Out Alerts</h2>", unsafe_allow_html=True)
    stock_out_alerts = filtered_data[filtered_data['Quantity_In_Stock'] == 0]
    if not stock_out_alerts.empty:
        st.dataframe(stock_out_alerts[['Brand', 'Category', 'Quantity_In_Stock']])
    else:
        st.write("No stock-out alerts for the selected filters.")
    st.markdown("</div>", unsafe_allow_html=True)

    # Sales Contribution Analysis by Brand
    st.markdown("<div style='border: 1px solid #cccccc; border-radius: 10px; padding: 15px; margin-bottom: 20px;'>", unsafe_allow_html=True)
    st.markdown("<h2>Sales Contribution by Brand</h2>", unsafe_allow_html=True)
    total_sales = filtered_data['Quantity_Sold'].sum()
    if total_sales > 0:
        filtered_data['Sales_Contribution'] = (filtered_data['Quantity_Sold'] / total_sales) * 100
        sales_contribution = filtered_data[['Brand', 'Category', 'Sales_Contribution']].dropna().reset_index(drop=True)
        st.dataframe(sales_contribution)
    else:
        st.write("No data available for the selected filters.")
    st.markdown("</div>", unsafe_allow_html=True)

# Executive Insights and Recommendations Page
elif page == "Executive Insights and Recommendations":
    st.markdown("<h1 style='text-align: center;'>Executive Insights and Recommendations</h1>", unsafe_allow_html=True)

    # Generate OpenAI-driven insights for key analyses
    def generate_data_driven_analysis(context, prompt):
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=1500,
                temperature=0.7
            )
            analysis = response['choices'][0]['message']['content'].strip()
            st.markdown(f"<div style='font-size: 15px;'>{analysis}</div>", unsafe_allow_html=True)
        except Exception as e:
            st.write("An error occurred while generating the analysis:", e)

    # Define questions for executive insights
    questions = [
        "Which items reach 75% and 50% sold, including days to sell out.",
        "Identify the weekly, monthly, and quarterly best-selling items.",
        "Describe non-moving products and their aging quantities.",
        "Describe slow-moving sizes within specific categories.",
        "Provide insights on variances and suggest strategies for improvement in general, daily use language.",
        "Recommend which products from our stock to prioritize for online sales.",
        "Describe unique products to enhance our online portfolio and what products are searched on Google as per Google Trends in the last 30 days in Madhya Pradesh.",
        "Identify the top 20% of products contributing to 80% of sales.",
        "Suggest strategies to reduce inventory of low-performing items."
    ]

    # Generate insights for each question
    for question in questions:
        st.markdown(f"<h2>{question}</h2>", unsafe_allow_html=True)
        if not filtered_data.empty:
            prompt = f"""
            You are a fashion retail expert. Based on the following data, {question}
            Data:
            {filtered_data.to_string(index=False)}
            Please provide a detailed analysis and recommendations.
            """
            generate_data_driven_analysis(filtered_data.to_string(index=False), prompt)
