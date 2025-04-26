import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import plotly.express as px
from datetime import datetime
from fpdf import FPDF
# Setting page config
st.set_page_config(page_title="Sales Insights Dashboard", layout="wide", initial_sidebar_state="expanded")

# Apply a light pastel Seaborn theme
sns.set_theme(style="whitegrid", palette="pastel")

# Title
st.title("📊 Sales Insights Dashboard")

# # Load your data
# @st.cache_data
# def load_data():
#     df = pd.read_excel("Amazon Target Combined Sales - Jan'24.xlsx")
#     # Ensure column names match the dataset
#     df.columns = df.columns.str.strip()  # Remove extra spaces
#     df['order_date'] = pd.to_datetime(df['Order Date'], errors='coerce')  # Adjust column name
#     df['total_revenue'] = pd.to_numeric(df['Invoice Amount'], errors='coerce')  # Adjust column name
#     return df

# df = load_data()
# Load your data
@st.cache_data
def load_data(uploaded_file):
    if uploaded_file is not None:
        # Read the uploaded Excel file
        df = pd.read_excel(uploaded_file)
        # Ensure column names match the dataset
        df.columns = df.columns.str.strip()  # Remove extra spaces
        df['order_date'] = pd.to_datetime(df['Order Date'], errors='coerce')  # Adjust column name
        df['total_revenue'] = pd.to_numeric(df['Invoice Amount'], errors='coerce')  # Adjust column name
        return df
    else:
        st.warning("Please upload an Excel file to proceed.")
        return None

# File uploader widget
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

# Load the data if a file is uploaded
df = load_data(uploaded_file)
# Check if data is loaded
if df is None:
    st.stop()  # Stop execution until a file is uploaded

# Sidebar Filters
st.sidebar.header("🔎 Filters")
date_range = st.sidebar.date_input("Select Date Range", [df['order_date'].min(), df['order_date'].max()])
products = st.sidebar.multiselect("Select Products", options=df['Item Description'].unique())  # Adjust column name
states = st.sidebar.multiselect("Select States", options=df['Ship To State'].unique())  # Adjust column name
fulfillment = st.sidebar.multiselect("Fulfillment Channel", options=df['Fulfillment Channel'].unique())  # Adjust column name

# Apply Filters
filtered_df = df.copy()
if date_range:
    filtered_df = filtered_df[(filtered_df['order_date'] >= pd.to_datetime(date_range[0])) & (filtered_df['order_date'] <= pd.to_datetime(date_range[1]))]
if products:
    filtered_df = filtered_df[filtered_df['Item Description'].isin(products)]
if states:
    filtered_df = filtered_df[filtered_df['Ship To State'].isin(states)]
if fulfillment:
    filtered_df = filtered_df[filtered_df['Fulfillment Channel'].isin(fulfillment)]

# Tabs
tab1, tab2, tab3, tab4 = st.tabs(["📈 Overview", "🛒 Products", "🚚 Shipping", "🔄 Returns"])

# 1. Overview Tab
with tab1:
    st.subheader("Daily Sales Trend")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Revenue", f"₹{df['Invoice Amount'].sum():,.2f}")
    col2.metric("Total Orders", df['Order Id'].nunique())
    col3.metric("Total Units Sold", df['Quantity'].sum())
    daily_sales = filtered_df.groupby('order_date')['total_revenue'].sum().reset_index()
    fig, ax = plt.subplots(figsize=(10, 5))
    sns.lineplot(data=daily_sales, x='order_date', y='total_revenue', ax=ax)
    ax.set_ylabel("Revenue ($)")
    st.pyplot(fig)


    st.subheader("State-wise Sales")
    # Sales grouped by state
    state_sales = df.groupby('Ship To State').agg({
        'Invoice Amount': 'sum',
        'Order Id': 'nunique',
        'Quantity': 'sum'
    }).sort_values('Invoice Amount', ascending=False).reset_index()
    
    # Metrics for Top 1 State
    top_state = state_sales.iloc[0] # Taking top 1 state based on sales
    
    st.markdown(f"### 🥇 Top State: **{top_state['Ship To State']}**")
    col4, col5, col6 = st.columns(3)
    col4.metric("Revenue", f"₹{top_state['Invoice Amount']:,.2f}")
    col5.metric("Orders", int(top_state['Order Id']))
    col6.metric("Units Sold", int(top_state['Quantity']))

    # Plotting Top 10 States
    fig2 = px.bar(
        state_sales.head(10), 
        x='Invoice Amount', 
        y='Ship To State', 
        orientation='h', 
        color='Ship To State', 
        title="Top 10 States by Sales"
    )
    st.plotly_chart(fig2, use_container_width=True)
    # Correct Insights
    top_state_name = top_state['Ship To State']
    top_product_name = filtered_df['Item Description'].mode()[0] if not filtered_df.empty else "N/A"

    st.markdown("#### Insights")
    st.info(
        f"🏆 Highest sales were recorded in **{top_state_name}**, "
        f"driven mainly by **{top_product_name}** during the selected period."
    )
# 2. Products Tab
with tab2:
    st.subheader("Top Products by Revenue")
    top_products = df.groupby('Item Description')['Invoice Amount'].sum().nlargest(10).reset_index()
    st.dataframe(top_products)
    top_products = filtered_df.groupby('Item Description')['total_revenue'].sum().sort_values(ascending=False).head(10)
    st.bar_chart(top_products)
    
# 3. Shipping Tab
with tab3:
    st.subheader("Shipping Amount Trend")
    st.metric("Total Shipping Amount", f"₹{df['Shipping Amount'].sum():,.2f}")
    if 'Shipping Amount' in filtered_df.columns:
        shipping_trend = filtered_df.groupby('order_date')['Shipping Amount'].sum()
        st.line_chart(shipping_trend)

# 4. Returns Tab
with tab4:
    # First, make sure 'Invoice Date' is in datetime format
    filtered_df['Invoice Date'] = pd.to_datetime(filtered_df['Invoice Date'], errors='coerce')
    filtered_df['Invoice Day'] = filtered_df['Invoice Date'].dt.date

    # -----------------------------
    # 8. Top Returned Products
    st.subheader("\U0001F501 Top Returned Products")
    returns_df = df[df['Credit Note No'].notnull()]
    st.write(f"Total Returns: {returns_df.shape[0]}")
    st.dataframe(returns_df[['Credit Note No', 'Credit Note Date', 'Invoice Number', 'Item Description', 'Invoice Amount']])

    returned_df = filtered_df[filtered_df['Quantity'] < 0]
    returned_products = returned_df['Item Description'].value_counts().reset_index().head(10)
    returned_products.columns = ['Item Description', 'Return Count']

    fig8, ax8 = plt.subplots()
    sns.barplot(data=returned_products, x='Return Count', y='Item Description', ax=ax8)
    st.pyplot(fig8)

    st.markdown("**Insight:** Highlights most returned products and return value trends.")
    st.markdown("**Use Case:** Product quality checks, return policy adjustments.")

    # -----------------------------
    # 9. Returns Over Time
    st.subheader("\U0001F4C9 Returns Over Time")
    returns_daily = filtered_df[filtered_df['Credit Note No'].notnull()].groupby('Invoice Day')['Invoice Amount'].sum().reset_index()

    fig9, ax9 = plt.subplots()
    sns.lineplot(data=returns_daily, x='Invoice Day', y='Invoice Amount', marker='o', color='red', ax=ax9)
    st.pyplot(fig9)

    st.markdown("**Insight:** Spot return spikes and their cause (could be faulty batches or delivery issues).")
    st.markdown("**Use Case:** Prevent future losses by correcting root problems.")


# Download Buttons
st.sidebar.header("⬇️ Downloads")

def convert_df(df):
    return df.to_csv(index=False).encode('utf-8')

csv = convert_df(filtered_df)
st.sidebar.download_button(
    label="Download Filtered Data as CSV",
    data=csv,
    file_name=f"filtered_data_{datetime.now().strftime('%Y%m%d')}.csv",
    mime='text/csv',
)


def generate_pdf():
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=14)
    pdf.cell(0, 10, "📄 Sales & Tax Insights Summary", ln=True, align='C')
    pdf.ln(10)

    # Total Revenue
    total_revenue = filtered_df['total_revenue'].sum()
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, f"🧾 Total Revenue: ₹{total_revenue:,.2f}", ln=True)
    pdf.ln(5)

    # Top State
    if not state_sales.empty:
        top_state_name = state_sales.loc[state_sales['Invoice Amount'].idxmax(), 'Ship To State']
        top_state_revenue = state_sales['Invoice Amount'].max()
        pdf.cell(0, 10, f"📍 Top State by Revenue: {top_state_name} (₹{top_state_revenue:,.2f})", ln=True)
    else:
        pdf.cell(0, 10, "📍 Top State by Revenue: No Data", ln=True)
    pdf.ln(5)

    # Top Product
    if not top_products.empty:
        top_product_name = top_products.index[0]
        top_product_revenue = top_products.iloc[0]
        pdf.cell(0, 10, f"🏆 Top Product: {top_product_name} (₹{top_product_revenue:,.2f})", ln=True)
    else:
        pdf.cell(0, 10, "🏆 Top Product: No Data", ln=True)
    pdf.ln(5)

    # Filters Applied
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "📅 Filters Applied:", ln=True)
    pdf.set_font("Arial", '', 12)
    pdf.multi_cell(0, 8, f"Date Range: {date_range[0]} to {date_range[1]}")
    pdf.multi_cell(0, 8, f"Selected Products: {', '.join(products) if products else 'All'}")
    pdf.multi_cell(0, 8, f"Selected States: {', '.join(states) if states else 'All'}")
    pdf.multi_cell(0, 8, f"Fulfillment Channels: {', '.join(fulfillment) if fulfillment else 'All'}")

    pdf.ln(5)
    pdf.set_font("Arial", 'I', 10)
    pdf.cell(0, 10, f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align='C')

    return pdf.output(dest='S').encode('latin-1')
