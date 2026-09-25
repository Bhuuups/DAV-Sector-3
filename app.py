import pandas as pd
import streamlit as st
import os
from datetime import datetime

# Page Configuration
st.set_page_config(page_title="ID Card Verification Portal", layout="centered")

EXCEL_FILE = '1.xlsx'
CORRECTIONS_FILE = 'corrections.csv'
PHOTOS_DIR = 'static/photos'

# Load Excel Data
@st.cache_data(ttl=60)
def load_data():
    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip()  # Column names ke extra spaces hatayein
    
    # Adm. No. ko Clean String me Convert Karein
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

# Multi-navigation (Parent View vs Admin View)
menu = st.sidebar.radio("Navigation", ["Parent Portal", "Admin Panel"])

if menu == "Parent Portal":
    st.subheader("🔍 Search Student Profile")
    
    adm_no_input = st.text_input("Enter Admission Number (Adm. No.):")
    
    if adm_no_input:
        clean_search_no = str(adm_no_input).strip()
        
        # Search by Admission No
        student = df[df['Adm. No. Clean'] == clean_search_no]
        
        if not student.empty:
            row = student.iloc[0]
            st.success(f"Record Found: **{row['Student Name']}**")
            
            # Photo display logic
            photo_name = str(row['photo']).strip() if pd.notna(row['photo']) else ""
            photo_path = os.path.join(PHOTOS_DIR, f"{photo_name}.jpg")
            
            col1, col2 = st.columns([1, 2])
            
            with col1:
                if photo_name and os.path.exists(photo_path):
                    st.image(photo_path, caption=f"Photo Code: {photo_name}", width=180)
                else:
                    st.warning(f"Photo Not Found ({photo_name})")
            
            with col2:
                st.markdown(f"**Sr No:** {row.get('Sr No', 'N/A')}")
                st.markdown(f"**Class & Section:** {row.get('CLASS', '')} - {row.get('SECTION', '')}")
                st.markdown(f"**Father's Name:** {row.get('Father\'s Name', '')}")
                st.markdown(f"**Mother's Name:** {row.get('Mother\'s Name', '')}")
                st.markdown(f"**House / Colour:** {row.get('colour', '')}")
                st.markdown(f"**Date of Birth (DOB):** {row.get('DOB', '')}")
                
                # Phone number formatting
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
                    
                    # Save image if uploaded
                    if new_photo is not None:
                        os.makedirs(PHOTOS_DIR, exist_ok=True)
                        save_path = os.path.join(PHOTOS_DIR, f"{photo_name}_updated.jpg")
                        with open(save_path, "wb") as f:
                            f.write(new_photo.getbuffer())
                    
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