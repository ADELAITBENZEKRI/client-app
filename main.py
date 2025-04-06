import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def create_pdf(order_details, selected_order):
    # Create a PDF in memory
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)

    # Set up fonts and styles for the header
    c.setFont("Helvetica-Bold", 16)
    c.drawString(15, 750, f"Invoice for Order ID: {selected_order}")
    c.setFont("Helvetica", 12)
    c.drawString(15, 730, f"Order Date: {order_details['Order Date'].values[0]}")
    c.drawString(15, 710, f"Ship Date: {order_details['Ship Date'].values[0]}")
    c.drawString(15, 690, f"Shipping Mode: {order_details['Ship Mode'].values[0]}")

    # Product Details Section
    y_position = 650
    c.setFont("Helvetica-Bold", 14)
    c.drawString(15, y_position, "Product Details")
    
    # Draw a line under the title
    c.setStrokeColorRGB(0, 0, 0)  # Black color for the line
    c.setLineWidth(1)
    c.line(1, y_position - 5, 500, y_position - 5)  # Draw line
    
    y_position -= 20  # Move down for the table

    # Set table headers with better column spacing
    c.setFont("Helvetica-Bold", 10)
    c.drawString(15, y_position, "Product Name")
    c.drawString(300, y_position, "Quantity")
    c.drawString(350, y_position, "Unit Price")
    c.drawString(450, y_position, "Discount")
    c.drawString(550, y_position, "Profit")
    
    # Draw a line under the headers
    c.line(1, y_position - 5, 600, y_position - 5)  # Draw line
    
    # Move down for product rows
    y_position -= 20

    # Draw product rows with better spacing
    for _, row in order_details.iterrows():
        c.setFont("Helvetica", 9)
        c.drawString(15, y_position, str(row['Product Name']))
        c.drawString(300, y_position, str(row['Quantity']))
        c.drawString(350, y_position, f"${row['Sales'] / row['Quantity']:.2f}")
        c.drawString(450, y_position, f"{row['Discount'] * 1}%")
        c.drawString(550, y_position, f"${row['Profit']:.2f}")
        y_position -= 20  # Move to the next line
        
        # If the table reaches the bottom of the page, create a new page
        if y_position < 15:
            c.showPage()  # Create new page
            c.setFont("Helvetica-Bold", 12)
            c.drawString(15, 750, f"Invoice for Order ID: {selected_order}")
            c.setFont("Helvetica", 10)
            c.drawString(15, 730, f"Order Date: {order_details['Order Date'].values[0]}")
            c.drawString(15, 710, f"Ship Date: {order_details['Ship Date'].values[0]}")
            c.drawString(15, 690, f"Shipping Mode: {order_details['Ship Mode'].values[0]}")
            y_position = 650  # Reset y_position for the new page

    # Add total calculations section with proper spacing
    c.setFont("Helvetica-Bold", 12)
    c.drawString(15, y_position - 20, f"Total Sales: ${order_details['Sales'].sum():.2f}")
    c.drawString(15, y_position - 40, f"Total Discount: ${order_details['Discount'].sum():.2f}")
    c.drawString(15, y_position - 60, f"Total Profit: ${order_details['Profit'].sum():.2f}")

    # Finalize the PDF
    c.save()

    # Return the buffer to be downloaded
    buffer.seek(0)
    return buffer

st.title("Purchase Order Management and Analysis")

# Upload the file
uploaded_file = st.file_uploader("Choose a excel file", type="xlsx")

if uploaded_file is not None:
    # Load the dataset
    df = pd.read_excel(uploaded_file)

    # Data Preview
    st.markdown("<h2 style='color:#4CAF50;'>Data Preview</h2>", unsafe_allow_html=True)
    st.write(df.head())

    # Data Summary
    st.markdown("<h2 style='color:#FF5722;'>Data Summary</h2>", unsafe_allow_html=True)
    st.write(df.describe())

    # Add space between sections
    st.markdown("<br><br>", unsafe_allow_html=True)

    # Filter by customer
    st.markdown("<h2 style='color:#2196F3;'>Filter Data by Customer</h2>", unsafe_allow_html=True)
    customer_names = df['Customer Name'].unique()
    selected_customer = st.selectbox("Select Customer", customer_names)
    customer_data = df[df['Customer Name'] == selected_customer]
    st.write(customer_data)

    # Add space between sections
    st.markdown("<br><br>", unsafe_allow_html=True)

    # Purchase Order Invoice Generation
    st.markdown("<h2 style='color:#9C27B0;'>Generate Purchase Order Invoice</h2>", unsafe_allow_html=True)
    order_data = customer_data[['Order ID', 'Order Date', 'Ship Date', 'Ship Mode','Product Name', 'Sales', 'Quantity', 'Discount', 'Profit']]
    
    # Allow user to select an Order ID
    selected_order = st.selectbox("Select Order ID", order_data['Order ID'].unique())
    order_details = order_data[order_data['Order ID'] == selected_order]
    
    if not order_details.empty:
       # Display invoice header
        st.write(f"### Invoice for Order ID: {selected_order}")
        st.write(f"**Order Date:** {order_details['Order Date'].values[0]}")
        st.write(f"**Ship Date:** {order_details['Ship Date'].values[0]}")
        st.write(f"**Shipping Mode:** {order_details['Ship Mode'].values[0]}")

        # Prepare the data for the invoice table
        invoice_table = order_details[['Product Name', 'Quantity', 'Sales', 'Discount', 'Profit']]
    
        # Calculate unit price and add it to the table
        invoice_table.loc[:, 'Unit Price'] = invoice_table['Sales'] / invoice_table['Quantity']
    
        # Reorder columns to have unit price first
        invoice_table_1 = invoice_table[['Product Name', 'Quantity', 'Unit Price', 'Discount', 'Profit']]
    
        # Display the invoice as a table
        st.dataframe(invoice_table_1)
        
        # Calculate total sales after discount
        total_sales = order_details['Sales'].sum()
        total_discount = order_details['Discount'].sum()
        total_profit = order_details['Profit'].sum()

        st.write(f"### Total Sales: ${total_sales:.2f}")
        st.write(f"### Total Discount: ${total_discount:.2f}")
        st.write(f"### Total Profit: ${total_profit:.2f}")
        # Button to download invoice as PDF
        pdf_button = st.button("Download Invoice as PDF")

        if pdf_button:
            # Create PDF from order details
            pdf_buffer = create_pdf(order_details, selected_order)

            # Provide PDF for download
            st.download_button(
                label="Download PDF",
                data=pdf_buffer,
                file_name=f"Invoice_{selected_order}.pdf",
                mime="application/pdf"
            )
    # Add space between sections
    st.markdown("<br><br>", unsafe_allow_html=True)

    # Sales Analysis by Category
    st.markdown("<h2 style='color:#3F51B5;'>Sales Analysis by Category</h2>", unsafe_allow_html=True)
    category_sales = df.groupby('Category')['Sales'].sum().reset_index()
    
    category_sales_chart = category_sales.plot(kind='bar', x='Category', y='Sales', legend=False)
    category_sales_chart.set_title('Total Sales by Category')
    category_sales_chart.set_ylabel('Sales')
    
    st.pyplot(category_sales_chart.figure)

    # Add space between sections
    st.markdown("<br><br>", unsafe_allow_html=True)

    # Sales vs Profit Visualization
    st.markdown("<h2 style='color:#8BC34A;'>Sales vs Profit Analysis</h2>", unsafe_allow_html=True)
    sales_profit_fig, ax = plt.subplots()
    ax.scatter(df['Sales'], df['Profit'], alpha=0.6, color='green')
    ax.set_xlabel('Sales')
    ax.set_ylabel('Profit')
    ax.set_title('Sales vs Profit')
    st.pyplot(sales_profit_fig)
     # Add space between sections
    st.markdown("<br><br>", unsafe_allow_html=True)
    # Optimized Quantity Distribution by City on Map (United States Only)
    st.markdown("<h2 style='color:#FF9800;'>Sales and Profits by Region or City</h2>", unsafe_allow_html=True)

    # Region or City Selector
    st.subheader("View Sales and Profit by Region or City")
    
    # Create a list of regions and cities
    regions = df['Region'].unique().tolist()
    cities = df['City'].unique().tolist()
    
    # User selects either Region or City
    selection_type = st.selectbox("Choose selection type", ["Region", "City"])

    if selection_type == "Region":
        selected_region = st.selectbox("Select Region", regions)
        filtered_data = df[df['Region'] == selected_region]
    else:
        selected_city = st.selectbox("Select City", cities)
        filtered_data = df[df['City'] == selected_city]

    # Display sales and profit summary for the selected region/city
    if not filtered_data.empty:
        total_sales = filtered_data['Sales'].sum()
        total_profit = filtered_data['Profit'].sum()
        
        st.write(f"**Total Sales in {selection_type}: {selected_region if selection_type == 'Region' else selected_city}**: ${total_sales:,.2f}")
        st.write(f"**Total Profit in {selection_type}: {selected_region if selection_type == 'Region' else selected_city}**: ${total_profit:,.2f}")

        # Plotting the sales and profit for selected region/city
        fig, ax = plt.subplots()
        ax.bar(['Sales', 'Profit'], [total_sales, total_profit], color=['blue', 'green'])
        ax.set_title(f"Sales and Profit in {selected_region if selection_type == 'Region' else selected_city}")
        ax.set_ylabel('Amount in USD')
        st.pyplot(fig)

    else:
        st.write(f"No data found for the selected {selection_type.lower()}.")

else:
    st.write("Waiting on file upload...")
