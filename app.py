import pandas as pd
import streamlit as st
import os
import requests
from datetime import datetime

# Page Configuration
st.set_page_config(page_title="DAV Sector - 3", layout="centered")

EXCEL_FILE = '1.xlsx'
CORRECTIONS_FILE = 'corrections.csv'

# Google Drive Folder ID
GDRIVE_FOLDER_ID = '1x9XLO3npiQ5u249fsLbqAFxwsdKz6X85'

# Load Excel Data
@st.cache_data(ttl=60)
def load_data():
    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip()  # Extra spaces hatayein
    
    # Adm. No. clean string conversion
    if 'Adm. No.' in df.columns:
        df['Adm. No. Clean'] = (
            pd.to_numeric(df['Adm. No.'], errors='coerce')
            .fillna(0)
            .astype(int)
            .astype(str)
            .str.strip()
        )
    return df

df = load_data()

st.title("📇 Student ID Card Verification Portal")
st.write("Apne bachhe ki ID Card details check karein aur galat detail hone par correction submit karein.")

# Navigation (Parent View vs Admin View)
menu = st.sidebar.radio("Navigation", ["Parent Portal", "Admin Panel"])

if menu == "Parent Portal":
    st.subheader("🔍 Search Student Profile")
    
    adm_no_input = st.text_input("Enter Admission Number (Adm. No.):")
    
    if adm_no_input:
        clean_search_no = str(adm_no_input).strip()
        
        # Search student by Adm. No.
        student = df[df['Adm. No. Clean'] == clean_search_no]
        
        if not student.empty:
            row = student.iloc[0]
            st.success(f"Record Found: **{row['Student Name']}**")
            
            # Photo Code
            photo_name = str(row['photo']).strip() if pd.notna(row['photo']) else ""
            
            col1, col2 = st.columns([1, 2])
            
            with col1:
                if photo_name:
                    # Google Drive Direct Thumbnail Link
                    photo_url = f"https://drive.google.com/thumbnail?id={photo_name}&sz=w500"
                    st.image(
                        photo_url, 
                        caption=f"Photo Code: {photo_name}", 
                        width=180
                    )
                else:
                    st.warning("Photo Code Not Found")
            
            with col2:
                st.markdown(f"**Sr No:** {row.get('Sr No', 'N/A')}")
                st.markdown(f"**Class & Section:** {row.get('CLASS', '')} - {row.get('SECTION', '')}")
                st.markdown(f"**Father's Name:** {row.get('Father\'s Name', '')}")
                st.markdown(f"**Mother's Name:** {row.get('Mother\'s Name', '')}")
                st.markdown(f"**House / Colour:** {row.get('colour', '')}")
                st.markdown(f"**Date of Birth (DOB):** {row.get('DOB', '')}")
                
                # Phone formatting
                phone_val = row.get('Phone No.', '')
                phone_str = str(int(phone_val)) if pd.notna(phone_val) and str(phone_val).replace('.','').isdigit() else str(phone_val)
                st.markdown(f"**Phone No.:** {phone_str}")
                
                st.markdown(f"**Address:** {row.get('Address', '')}")
                st.markdown(f"**Mode:** {row.get('MODE', '')}")

            st.divider()
            
            # Correction Form
            st.subheader("✏️ Request Correction / Update")
            st.info("Agar kisi detail me galti hai toh sahi details niche darj karein:")
            
            with st.form(key="correction_form"):
                new_name = st.text_input("Correct Student Name", value=str(row['Student Name']))
                new_dob = st.text_input("Correct DOB", value=str(row.get('DOB', '')))
                new_father = st.text_input("Correct Father Name", value=str(row.get("Father's Name", '')))
                new_mother = st.text_input("Correct Mother Name", value=str(row.get("Mother's Name", '')))
                new_colour = st.text_input("Correct House / Colour", value=str(row.get('colour', '')))
                new_contact = st.text_input("Correct Phone Number", value=phone_str)
                new_address = st.text_area("Correct Address", value=str(row.get('Address', '')))
                
                new_photo = st.file_uploader("Upload New Photo (Optional)", type=['jpg', 'jpeg', 'png'])
                
                submit_btn = st.form_submit_button("Submit Corrections")
                
                if submit_btn:
                    correction_data = {
                        'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'Adm. No.': row['Adm. No. Clean'],
                        'Student Name': new_name,
                        'DOB': new_dob,
                        'Father Name': new_father,
                        'Mother Name': new_mother,
                        'Colour': new_colour,
                        'Phone No.': new_contact,
                        'Address': new_address,
                        'Photo Code': photo_name,
                        'New Photo Uploaded': 'Yes' if new_photo is not None else 'No'
                    }
                    
                    corr_df = pd.DataFrame([correction_data])
                    
                    if not os.path.exists(CORRECTIONS_FILE):
                        corr_df.to_csv(CORRECTIONS_FILE, index=False)
                    else:
                        corr_df.to_csv(CORRECTIONS_FILE, mode='a', header=False, index=False)
                        
                    st.success("✅ Correction request successfully submit ho gayi hai!")
        else:
            st.error("Admission Number nahi mila. Kripya sahi Adm. No. enter karein.")

elif menu == "Admin Panel":
    st.subheader("🔒 Admin Dashboard")
    password = st.text_input("Enter Admin Password:", type="password")
    
    if password == "admin123":
        st.success("Welcome Admin!")
        
        if os.path.exists(CORRECTIONS_FILE):
            corrections = pd.read_csv(CORRECTIONS_FILE)
            st.write(f"Total Corrections Received: **{len(corrections)}**")
            st.dataframe(corrections)
            
            st.download_button(
                label="📥 Download Corrections CSV",
                data=corrections.to_csv(index=False),
                file_name="student_corrections.csv",
                mime="text/csv"
            )
        else:
            st.info("Abhi tak koi correction request nahi aayi hai.")
