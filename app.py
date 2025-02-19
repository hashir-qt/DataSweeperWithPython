import streamlit as st  # type: ignore
import pandas as pd  # type: ignore
import os
import pdfkit  # type: ignore
from io import BytesIO
from docx import Document  # type: ignore
import tempfile

# Set up PDFKit configuration
config = pdfkit.configuration(wkhtmltopdf=r"D:\wkhtmltopdf\bin\wkhtmltopdf.exe")



st.set_page_config(page_title="Data Formatter", layout='wide')
st.title("✨ Data Formatter ✨")


st.markdown(
    """
<style>
    /* Center content */
    .main {
        background-color: #0a192f;  /* Deep Navy Blue */
        color: white;
        display: flex;
        justify-content: center;
        align-items: center;
        min-height: 100vh;
        text-align: center;
    }

    /* Main container */
    .block-container {
        max-width: 90%;
        width: 100%;
        padding: 2rem;
        border-radius: 12px;
        background-color: #112240;  /* Darker Navy Blue */
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
        margin: auto;
    }

    /* Headings */
    h1, h2, h3, h4, h5, h6 {
        color: white;
        font-weight: bold;
        letter-spacing: 0.5px;
    }

    /* Paragraphs and text */
    p, span, div {
        color: #C0C0C0 !important;
    }

    /* Buttons */
    .stButton>button {
        border: none;
        border-radius: 8px;
        background: linear-gradient(135deg, #0078D7, #004f8b);
        color: white;
        padding: 0.75rem 1.5rem;
        font-size: 1rem;
        font-weight: 600;
        transition: all 0.3s ease-in-out;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.4);
        width: 100%;
        max-width: 300px;
        margin: 10px auto;
    }

    .stButton>button:hover {
        background: linear-gradient(135deg, #005a9e, #003f6b);
        transform: scale(1.05);
        cursor: pointer;
    }

    /* Dataframe & Tables */
    .stDataFrame, .stTable {
        border-radius: 10px;
        overflow: hidden;
        background-color: #102a43;
        color: white;
    }

    /* Adjust font sizes for small screens */
    @media (max-width: 768px) {
        .block-container {
            padding: 1.5rem;
        }

        h1 {
            font-size: 1.8rem;
        }

        h2 {
            font-size: 1.5rem;
        }

        .stButton>button {
            font-size: 0.9rem;
            padding: 0.6rem 1.2rem;
        }
    }
</style>
    """,
    unsafe_allow_html=True  # 'unsafe_allow_html' permits raw HTML/CSS embedding in the Streamlit app
)
st.markdown("### Easily convert your files between CSV, Excel, and Word formats with built-in data cleaning, visualization, and seamless PDF export!")
# Upload files (CSV, Excel, Word)
uploaded_files = st.file_uploader(
    "📂 Upload your files (CSV, Excel, or Word):", 
    type=["csv", "xlsx", "docx"], 
    accept_multiple_files=True
)

if uploaded_files:
    for file in uploaded_files:
        file_ext = os.path.splitext(file.name)[-1].lower()

        # Handle CSV and Excel files
        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext == ".xlsx":
            df = pd.read_excel(file)
        elif file_ext == ".docx":
            doc = Document(file)
        else:
            st.error("❌ Unsupported file format. Please upload CSV, Excel, or Word files.")
            continue

        st.markdown(f"### 📄 File: `{file.name}`")
        st.write(f"**File Size:** {file.getbuffer().nbytes / 1024:.2f} KB")

        # Preview data (for CSV/Excel)
        if file_ext in [".csv", ".xlsx"]:
            st.write("🔍 **Data Preview**")
            st.dataframe(df.head())

            # Data Cleaning Options
            st.subheader("🛠 Data Cleaning Options")
            if st.checkbox(f"🧹 Clean Data for {file.name}"):
                col1, col2 = st.columns(2)

                with col1:
                    if st.button(f"🗑 Remove Duplicates", key=f"dup_{file.name}"):
                        df.drop_duplicates(inplace=True)
                        st.success("✔ Duplicates removed successfully!")

                with col2:
                    if st.button(f"🔄 Fill Missing Values", key=f"fill_{file.name}"):
                        numeric_cols = df.select_dtypes(include=['number']).columns
                        df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())  # type: ignore
                        st.success("✔ Missing values filled successfully!")

            # Column Selection
            st.subheader("📝 Select Columns to Convert")
            columns = st.multiselect(f"📌 Choose Columns", df.columns, default=df.columns)
            df = df[columns]

            # Data Visualization
            st.subheader("📊 Data Visualization")
            if st.checkbox(f"📈 Show Visualization for {file.name}"):
                st.bar_chart(df.select_dtypes(include='number').iloc[:, :2])

        # Conversion Options
        st.subheader("🔄 Convert & Download")
        if file_ext in [".csv", ".xlsx"]:
            conversion_type = st.radio(f"Convert {file.name} to:", ["CSV", "Excel"], key=file.name)
        elif file_ext == ".docx":
            conversion_type = "PDF"

        if st.button(f"💾 Convert {file.name}", key=f"convert_{file.name}"):
            buffer = BytesIO()

            # Convert CSV/Excel files
            if file_ext in [".csv", ".xlsx"]:
                if conversion_type == "CSV":
                    df.to_csv(buffer, index=False)
                    file_name = file.name.replace(file_ext, ".csv")
                    mime_type = "text/csv"
                elif conversion_type == "Excel":
                    df.to_excel(buffer, index=False)
                    file_name = file.name.replace(file_ext, ".xlsx")
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            # Convert Word to PDF
            elif file_ext == ".docx":
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    html_content = "<html><body>"
                    for para in doc.paragraphs:
                        html_content += f"<p>{para.text}</p>"
                    html_content += "</body></html>"

                    pdfkit.from_string(html_content, temp_pdf.name, configuration=config)

                    with open(temp_pdf.name, "rb") as f:
                        pdf_data = f.read()

                buffer = BytesIO(pdf_data)
                file_name = file.name.replace(".docx", ".pdf")
                mime_type = "application/pdf"

            buffer.seek(0)

            # Download button
            st.download_button(
                label=f"⬇️ Download {file_name}",
                data=buffer,
                file_name=file_name,
                mime=mime_type,
                key=f"download_{file.name}"
            )
