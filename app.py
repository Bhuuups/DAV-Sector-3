import pandas as pd
import streamlit as st
import os
import io
import zipfile
from datetime import datetime
from PIL import Image

# Page Configuration
st.set_page_config(page_title="ID Card Verification Portal", layout="centered")

EXCEL_FILE = '1.xlsx'
CORRECTIONS_FILE = 'corrections.csv'
PHOTOS_DIR = 'static/photos'

# Helper Function: Compress Image under 1 MB and save as JPEG
def process_and_save_image(image_file, save_path, max_size_mb=1.0):
    try:
        img = Image.open(image_file)
        # Convert RGBA/P mode to RGB for JPEG compatibility
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
            
        # Resize if dimensions are unnecessarily huge (e.g. max 1920px width/height)
        max_dim = 1920
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            
        # Quality compression loop to ensure file size <= max_size_mb
        quality = 85
        buffer = io.BytesIO()
        
        while quality >= 20:
            buffer.seek(0)
            buffer.truncate(0)
            img.save(buffer, format="JPEG", quality=quality, optimize=True)
            size_mb = len(buffer.getvalue()) / (1024 * 1024)
            if size_mb <= max_size_mb:
                break
            quality -= 10
            
        with open(save_path, "wb") as f:
            f.write(buffer.getvalue())
        return True
    except Exception as e:
        st.error(f"Image compression/saving error: {e}")
        return False

# Load Excel Data
@st.cache_data(ttl=60)
def load_data():
    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip()
    
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

# Navigation
menu = st.sidebar.radio("Navigation", ["Parent Portal", "Admin Panel"])

if menu == "Parent Portal":
    st.subheader("🔍 Search Student Profile")
    
    adm_no_input = st.text_input("Enter Admission Number (Adm. No.):")
    
    if adm_no_input:
        clean_search_no = str(adm_no_input).strip()
        student = df[df['Adm. No. Clean'] == clean_search_no]
        
        if not student.empty:
            row = student.iloc[0]
            st.success(f"Record Found: **{row['Student Name']}**")
            
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
                
                phone_val = row.get('Phone No.', '')
                phone_str = str(int(phone_val)) if pd.notna(phone_val) and str(phone_val).replace('.','').isdigit() else str(phone_val)
                st.markdown(f"**Phone No.:** {phone_str}")
                
                st.markdown(f"**Address:** {row.get('Address', '')}")
                st.markdown(f"**Mode:** {row.get('MODE', '')}")

            st.divider()
            
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
                
                st.markdown("---")
                st.markdown("📷 **Upload / Click New Photo (Optional - Auto Compress Under 1 MB)**")
                photo_method = st.radio("Photo Upload Mode:", ["File Upload (Gallery)", "Camera Input (Live Photo)"], horizontal=True)
                
                new_photo = None
                if photo_method == "File Upload (Gallery)":
                    new_photo = st.file_uploader("Upload New Photo", type=['jpg', 'jpeg', 'png'])
                else:
                    new_photo = st.camera_input("Take Live Photo")
                
                submit_btn = st.form_submit_button("Submit Corrections")
                
                if submit_btn:
                    has_new_photo = 'No'
                    saved_photo_filename = ""
                    
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
                    
                    # Compress and save image if provided
                    if new_photo is not None:
                        os.makedirs(PHOTOS_DIR, exist_ok=True)
                        saved_photo_filename = f"{photo_name}_updated.jpg"
                        save_path = os.path.join(PHOTOS_DIR, saved_photo_filename)
                        
                        success = process_and_save_image(new_photo, save_path, max_size_mb=1.0)
                        if success:
                            has_new_photo = 'Yes'
                            changes.append("New Photo Uploaded (Compressed < 1MB)")
                    
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
        
        tab1, tab2 = st.tabs(["📊 Corrections List", "➕ Add New Student Data"])
        
        with tab1:
            if os.path.exists(CORRECTIONS_FILE):
                try:
                    corrections = pd.read_csv(CORRECTIONS_FILE, on_bad_lines='skip')
                except Exception:
                    corrections = pd.read_csv(CORRECTIONS_FILE, engine='python', on_bad_lines='skip')
                
                if 'Adm. No.' in corrections.columns:
                    unique_corrections = corrections.drop_duplicates(subset=['Adm. No.'], keep='last')
                else:
                    unique_corrections = corrections
                    
                st.write(f"Total Submissions: **{len(corrections)}** | Unique Students: **{len(unique_corrections)}**")
                st.dataframe(unique_corrections)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    unique_corrections.to_excel(writer, index=False, sheet_name='Corrections')
                
                st.download_button(
                    label="📊 Download Correction Excel (Unique Records)",
                    data=buffer.getvalue(),
                    file_name="student_corrections.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.divider()
                st.subheader("🖼️ Download Photos")
                col1, col2 = st.columns(2)
                
                with col1:
                    zip_buffer_new = io.BytesIO()
                    new_photos_count = 0
                    
                    with zipfile.ZipFile(zip_buffer_new, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for idx, c_row in unique_corrections.iterrows():
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
                    zip_buffer_all = io.BytesIO()
                    all_photos_count = 0
                    
                    with zipfile.ZipFile(zip_buffer_all, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for idx, c_row in unique_corrections.iterrows():
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

                st.divider()
                if st.button("🗑️ Clear / Reset All Correction Data"):
                    os.remove(CORRECTIONS_FILE)
                    st.success("All correction records cleared!")
                    st.rerun()
            else:
                st.info("Abhi tak koi correction request nahi aayi hai.")
                
        with tab2:
            st.subheader("➕ Naya Student Record Add Karein")
            st.info("Yahan se naya student data add karne par wo direct '1.xlsx' main database file me save ho jayega.")
            
            with st.form(key="add_student_form"):
                add_sr_no = st.number_input("Sr No", value=len(df)+1, step=1)
                add_adm_no = st.text_input("Adm. No. (Admission Number)*")
                add_name = st.text_input("Student Name*")
                add_father = st.text_input("Father's Name")
                add_mother = st.text_input("Mother's Name")
                add_class = st.text_input("CLASS (e.g. VI, X, XII)")
                add_section = st.text_input("SECTION (e.g. A1, B)")
                add_photo_code = st.text_input("Photo Code (e.g. KR-1500)")
                add_colour = st.text_input("House / Colour (e.g. HANSRAJ)")
                add_dob = st.text_input("DOB (DD/MM/YYYY)")
                add_phone = st.text_input("Phone No.")
                add_address = st.text_area("Address")
                add_mode = st.selectbox("MODE", ["SELF", "SCHOOL BUS", "OTHER"])
                
                st.markdown("---")
                st.markdown("📷 **Student Photo (Auto Compress Under 1 MB)**")
                add_photo_method = st.radio("Photo Source:", ["File Upload (Gallery)", "Camera Input (Live Photo)"], key="add_photo_mode", horizontal=True)
                
                uploaded_photo_file = None
                if add_photo_method == "File Upload (Gallery)":
                    uploaded_photo_file = st.file_uploader("Upload Student Photo", type=['jpg', 'jpeg', 'png'], key="add_file_up")
                else:
                    uploaded_photo_file = st.camera_input("Take Live Photo", key="add_cam_up")
                
                add_submit = st.form_submit_button("➕ Save Student Data")
                
                if add_submit:
                    if not add_adm_no.strip() or not add_name.strip():
                        st.error("Adm. No. aur Student Name fill karna zaroori hai.")
                    else:
                        final_photo_code = add_photo_code.strip() if add_photo_code.strip() else f"NEW_{add_adm_no.strip()}"
                        
                        if uploaded_photo_file is not None:
                            os.makedirs(PHOTOS_DIR, exist_ok=True)
                            photo_save_path = os.path.join(PHOTOS_DIR, f"{final_photo_code}.jpg")
                            process_and_save_image(uploaded_photo_file, photo_save_path, max_size_mb=1.0)
                        
                        new_student_dict = {
                            'Sr No': add_sr_no,
                            'Adm. No.': add_adm_no.strip(),
                            'Student Name': add_name.strip(),
                            "Father's Name": add_father.strip(),
                            "Mother's Name": add_mother.strip(),
                            'CLASS': add_class.strip(),
                            'SECTION': add_section.strip(),
                            'photo': final_photo_code,
                            'colour': add_colour.strip(),
                            'DOB': add_dob.strip(),
                            'Phone No.': add_phone.strip(),
                            'Address': add_address.strip(),
                            'MODE': add_mode
                        }
                        
                        try:
                            current_excel_df = pd.read_excel(EXCEL_FILE)
                            new_row_df = pd.DataFrame([new_student_dict])
                            updated_excel_df = pd.concat([current_excel_df, new_row_df], ignore_index=True)
                            updated_excel_df.to_excel(EXCEL_FILE, index=False)
                            
                            st.cache_data.clear()
                            st.success(f"✅ Student **{add_name.strip()}** successfully Excel database me add ho gaya hai!")
                        except Exception as e:
                            st.error(f"Excel file update karne me error aaya: {e}")
