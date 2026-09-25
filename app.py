import pandas as pd
import streamlit as st
import os
import io
import zipfile
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
                    p_jpg = os.path.join(PHOTOS_DIR, f"{photo_name}.jpg")
                    p_JPG = os.path.join(PHOTOS_DIR, f"{photo_name}.JPG")
                    p_jpeg = os.path.join(PHOTOS_DIR, f"{photo_name}.jpeg")
                    p_png = os.path.join(PHOTOS_DIR, f"{photo_name}.png")
                    
                    if os.path.exists(p_jpg):
                        st.image(p_jpg, caption=f"Photo Code: {photo_name}", width=180)
                    elif os.path.exists(p_JPG):
                        st.image(p_JPG, caption=f"Photo Code: {photo_name}", width=180)
                    elif os.path.exists(p_jpeg):
                        st.image(p_jpeg, caption=f"Photo Code: {photo_name}", width=180)
                    elif os.path.exists(p_png):
                        st.image(p_png, caption=f"Photo Code: {photo_name}", width=180)
                    else:
                        st.info(f"Photo Code: {photo_name}")
                        st.warning("Photo File Not Found in Server Folder")
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
                    has_new_photo = 'No'
                    saved_photo_filename = ""
                    
                    # Track changes made by parent (Cleaning commas and quotes for CSV safety)
                    changes = []
                    if str(new_name).strip() != str(row['Student Name']).strip():
                        changes.append(f"Name: {row['Student Name']} -> {new_name}")
                    if str(new_dob).strip() != str(row.get('DOB', '')).strip():
                        changes.append(f"DOB: {row.get('DOB', '')} -> {new_dob}")
                    if str(new_father).strip() != str(row.get("Father's Name", '')).strip():
                        changes.append(f"Father: {row.get('Father\'s Name', '')} -> {new_father}")
                    if str(new_mother).strip() != str(row.get("Mother's Name", '')).strip():
                        changes.append(f"Mother: {row.get('Mother\'s Name', '')} -> {new_mother}")
                    if str(new_colour).strip() != str(row.get('colour', '')).strip():
                        changes.append(f"Colour: {row.get('colour', '')} -> {new_colour}")
                    if str(new_contact).strip() != phone_str.strip():
                        changes.append(f"Phone: {phone_str} -> {new_contact}")
                    if str(new_address).strip() != str(row.get('Address', '')).strip():
                        clean_addr_old = str(row.get('Address', '')).replace('\n', ' ').replace(',', ' ')
                        clean_addr_new = str(new_address).replace('\n', ' ').replace(',', ' ')
                        changes.append(f"Address: {clean_addr_old} -> {clean_addr_new}")
                    
                    # Save image if uploaded
                    if new_photo is not None:
                        os.makedirs(PHOTOS_DIR, exist_ok=True)
                        has_new_photo = 'Yes'
                        saved_photo_filename = f"{photo_name}_updated.jpg"
                        save_path = os.path.join(PHOTOS_DIR, saved_photo_filename)
                        with open(save_path, "wb") as f:
                            f.write(new_photo.getbuffer())
                        changes.append("New Photo Uploaded")
                    
                    change_summary = " | ".join(changes) if changes else "No Text Change"
                    
                    correction_data = {
                        'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'Adm. No.': row['Adm. No. Clean'],
                        'Student Name': str(new_name).replace(',', ' '),
                        'Changed Fields': change_summary,
                        'DOB': str(new_dob).replace(',', ' '),
                        'Father Name': str(new_father).replace(',', ' '),
                        'Mother Name': str(new_mother).replace(',', ' '),
                        'Colour': str(new_colour).replace(',', ' '),
                        'Phone No.': str(new_contact).replace(',', ' '),
                        'Address': str(new_address).replace('\n', ' ').replace(',', ' '),
                        'Photo Code': photo_name,
                        'New Photo Uploaded': has_new_photo,
                        'Saved Photo Filename': saved_photo_filename
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
            try:
                corrections = pd.read_csv(CORRECTIONS_FILE, on_bad_lines='skip')
            except Exception as e:
                st.error("Corrupted CSV structure detected. Resetting view logic.")
                corrections = pd.read_csv(CORRECTIONS_FILE, engine='python', on_bad_lines='skip')
                
            st.write(f"Total Corrections Received: **{len(corrections)}**")
            st.dataframe(corrections)
            
            # 1. Download Correction Excel File (.xlsx)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                corrections.to_excel(writer, index=False, sheet_name='Corrections')
            
            st.download_button(
                label="📊 Download Correction Excel (With Change Log)",
                data=buffer.getvalue(),
                file_name="student_corrections.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.divider()
            st.subheader("🖼️ Download Photos")
            col1, col2 = st.columns(2)
            
            with col1:
                # 2. Download ONLY NEWLY Uploaded Photos ZIP
                zip_buffer_new = io.BytesIO()
                new_photos_count = 0
                
                with zipfile.ZipFile(zip_buffer_new, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for idx, c_row in corrections.iterrows():
                        saved_file = str(c_row.get('Saved Photo Filename', '')).strip()
                        if saved_file and saved_file != 'nan' and os.path.exists(os.path.join(PHOTOS_DIR, saved_file)):
                            filepath = os.path.join(PHOTOS_DIR, saved_file)
                            zip_file.write(filepath, arcname=saved_file)
                            new_photos_count += 1

                st.download_button(
                    label=f"📸 Download ONLY New Uploaded Photos ({new_photos_count})",
                    data=zip_buffer_new.getvalue(),
                    file_name="newly_uploaded_photos.zip",
                    mime="application/zip",
                    disabled=(new_photos_count == 0)
                )
            
            with col2:
                # 3. Download ALL Photos of Correction Students
                zip_buffer_all = io.BytesIO()
                all_photos_count = 0
                
                with zipfile.ZipFile(zip_buffer_all, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for idx, c_row in corrections.iterrows():
                        photo_code = str(c_row.get('Photo Code', '')).strip()
                        saved_file = str(c_row.get('Saved Photo Filename', '')).strip()
                        
                        if saved_file and saved_file != 'nan' and os.path.exists(os.path.join(PHOTOS_DIR, saved_file)):
                            filepath = os.path.join(PHOTOS_DIR, saved_file)
                            zip_file.write(filepath, arcname=saved_file)
                            all_photos_count += 1
                        elif photo_code and photo_code != 'nan':
                            for ext in ['.jpg', '.JPG', '.jpeg', '.png']:
                                orig_path = os.path.join(PHOTOS_DIR, f"{photo_code}{ext}")
                                if os.path.exists(orig_path):
                                    zip_file.write(orig_path, arcname=f"{photo_code}{ext}")
                                    all_photos_count += 1
                                    break

                st.download_button(
                    label=f"📁 Download All Correction Student Photos ({all_photos_count})",
                    data=zip_buffer_all.getvalue(),
                    file_name="all_correction_photos.zip",
                    mime="application/zip",
                    disabled=(all_photos_count == 0)
                )
        else:
            st.info("Abhi tak koi correction request nahi aayi hai.")
