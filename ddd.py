import pandas as pd
import streamlit as st
import plotly.express as px
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io

# Load the data
file_path = 'exp1.xlsx'  # Replace with your file path
sheet_name = 'bng '  # Specify the sheet to load
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Preprocess data
data.columns = data.columns.str.strip()  # Clean column names
data['Price'] = pd.to_numeric(data['Price'], errors='coerce')  # Convert Price to numeric
data['Date'] = data['Date'].str.strip()  # Clean Date column if it contains extra spaces
data['Module'] = data['Module'].str.strip()  # Clean Module column

# Dashboard Title
st.title("Interactive Expense Dashboard with PDF Export")

# --- Sidebar Filters ---
st.sidebar.header("Filters")

# 1. Price Range Filter
min_price = int(data['Price'].min())
max_price = int(data['Price'].max())
price_range = st.sidebar.slider("Select Price Range", min_value=min_price, max_value=max_price, value=(min_price, max_price))

# 2. Month Filter
unique_months = data['Date'].dropna().unique()  # Get unique months
selected_months = st.sidebar.multiselect("Select Month(s)", options=unique_months, default=unique_months)

# 3. Module Filter
unique_modules = data['Module'].dropna().unique()  # Get unique modules
selected_modules = st.sidebar.multiselect("Select Module(s)", options=unique_modules, default=unique_modules)

# --- Apply Filters ---
filtered_data = data[
    (data['Price'] >= price_range[0]) & 
    (data['Price'] <= price_range[1]) & 
    (data['Date'].isin(selected_months)) & 
    (data['Module'].isin(selected_modules))
]

# Display the selected filters in the sidebar
st.sidebar.write(f"Selected Price Range: {price_range[0]} to {price_range[1]}")
st.sidebar.write(f"Selected Months: {', '.join(selected_months)}")
st.sidebar.write(f"Selected Modules: {', '.join(selected_modules)}")

# --- Visualizations ---
# 1. Total Monthly Expense
st.header("Total Monthly Expense")
monthly_expense = filtered_data.groupby('Date')['Price'].sum().reset_index()
fig1 = px.bar(monthly_expense, x='Date', y='Price', title="Monthly Expense", labels={'Price': 'Total Expense', 'Date': 'Month'})
st.plotly_chart(fig1)

# 2. Category Breakdown
st.header("Category Breakdown")
category_expense = filtered_data.groupby('Module')['Price'].sum().reset_index()
fig2 = px.pie(category_expense, values='Price', names='Module', title="Expense by Category")
st.plotly_chart(fig2)

# 3. Priority Analysis
st.header("Priority Analysis")
priority_expense = filtered_data.groupby('Priority')['Price'].sum().reset_index()
fig3 = px.bar(priority_expense, x='Priority', y='Price', title="Expense by Priority", labels={'Price': 'Total Expense', 'Priority': 'Priority Level'})
st.plotly_chart(fig3)

# 4. Top Categories
st.header("Top 5 Categories by Expense")
top_categories = category_expense.nlargest(5, 'Price')
fig4 = px.bar(top_categories, x='Price', y='Module', orientation='h', title="Top 5 Categories", labels={'Price': 'Total Expense', 'Module': 'Category'})
st.plotly_chart(fig4)

# 5. Filtered Data Table
st.header("Filtered Expense Data Table")
st.dataframe(filtered_data)

# --- PDF Export Function ---
def create_pdf(data):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.drawString(100, 750, "Expense Report")
    c.drawString(100, 730, f"Price Range: {price_range[0]} - {price_range[1]}")
    c.drawString(100, 710, f"Selected Months: {', '.join(selected_months)}")
    c.drawString(100, 690, f"Selected Modules: {', '.join(selected_modules)}")

    # Add top expenses to the PDF
    c.drawString(100, 650, "Top 5 Categories:")
    y = 630
    for index, row in top_categories.iterrows():
        c.drawString(120, y, f"{row['Module']}: {row['Price']}")
        y -= 20

    c.drawString(100, y - 20, "Thank you for using the Expense Dashboard!")
    c.save()
    buffer.seek(0)
    return buffer

# --- PDF Download Button ---
if st.button("Download Report as PDF"):
    pdf_buffer = create_pdf(filtered_data)
    st.download_button(
        label="Download PDF",
        data=pdf_buffer,
        file_name="expense_report.pdf",
        mime="application/pdf"
    )
