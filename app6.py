import pandas as pd
import streamlit as st
from fpdf import FPDF
import tempfile

# Streamlit app setup
st.title("Generate Address Labels PDF")
st.write("Upload an Excel file with columns: Name, Name of School, Coordinator Name, Address, Contact Number.")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)
    
    # Display the uploaded data
    st.write("Preview of Uploaded Data:")
    st.dataframe(df)

    # Check if required columns are present
    required_columns = ["Name", "Name of School", "Coordinator Name", "Address", "Contact Number"]
    if all(col in df.columns for col in required_columns):
        
        # Function to generate the PDF
        def generate_pdf(data):
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            count = 0
            for _, row in data.iterrows():
                # Centralize the text with cell width equal to the page width
                pdf.cell(0, 10, txt=f"Name: {row['Name']}", ln=True, align="C")
                pdf.cell(0, 10, txt=f"School: {row['Name of School']}", ln=True, align="C")
                pdf.cell(0, 10, txt=f"Coordinator: {row['Coordinator Name']}", ln=True, align="C")
                pdf.cell(0, 10, txt=f"Address: {row['Address']}", ln=True, align="C")
                pdf.cell(0, 10, txt=f"Contact: {row['Contact Number']}", ln=True, align="C")
                pdf.cell(0, 10, txt=" ", ln=True)  # Empty line for spacing
                count += 1

                # Add a new page after every 20 addresses
                if count % 20 == 0:
                    pdf.add_page()

            # Save the PDF to a temporary file
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            pdf.output(temp_file.name)
            return temp_file.name

        # Generate the PDF
        if st.button("Generate PDF"):
            pdf_file_path = generate_pdf(df)
            st.success("PDF generated successfully!")
            with open(pdf_file_path, "rb") as pdf_file:
                st.download_button(
                    label="Download PDF",
                    data=pdf_file,
                    file_name="Address_Labels.pdf",
                    mime="application/pdf",
                )
    else:
        st.error(f"The uploaded file must contain these columns: {', '.join(required_columns)}")
