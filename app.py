import pandas as pd
import streamlit as st
import os
import io
import zipfile
from datetime import datetime
from PIL import Image

# Page Configuration
st.set_page_config(page_title="ID Card Verification & Printing Portal", layout="wide")

EXCEL_FILE = '1.xlsx'
CORRECTIONS_FILE = 'corrections.csv'
PRINT_STATUS_FILE = 'print_status.csv'
PHOTOS_DIR = 'static/photos'

# Helper Function: Compress Image under 1 MB and save as JPEG
def process_and_save_image(image_file, save_path, max_size_mb=1.0):
    try:
        img = Image.open(image_file)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
            
        max_dim = 1920
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            
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
        st.error(f"Image compression error: {e}")
        return False

# Helper Function: Load Print Status CSV
def load_print_status():
    if os.path.exists(PRINT_STATUS_FILE):
        try:
            return pd.read_csv(PRINT_STATUS_FILE, dtype={'Adm. No. Clean': str})
        except Exception:
            return pd.DataFrame(columns=['Adm. No. Clean', 'Is_Printed', 'Updated_At'])
    return pd.DataFrame(columns=['Adm. No. Clean', 'Is_Printed', 'Updated_At'])

# Helper Function: Save Print Status
def save_print_status(status_df):
    status_df.to_csv(PRINT_STATUS_FILE, index=False)

# Load Master Excel Data
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
    
    for col in ['CLASS', 'SECTION']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            
    return df

df = load_data()

st.title("📇 Student ID Card Portal")

# Navigation
menu = st.sidebar.radio("Navigation Menu", ["Parent Portal", "Teachers Checklist (Printing Status)", "Admin Panel"])

# ---------------------------------------------------------
# 1. PARENT PORTAL
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# 2. TEACHERS CHECKLIST (PRINTING STATUS PORTAL)
# ---------------------------------------------------------
elif menu == "Teachers Checklist (Printing Status)":
    st.subheader("👩‍🏫 Teachers Class-wise ID Card Printing Checklist")
    st.info("Class & Section filter select karke har student ke aage 'Printed ✅' status tick karein.")

    print_df = load_print_status()

    # Class & Section Filters
    all_classes = sorted([c for c in df['CLASS'].dropna().unique() if str(c).strip() != ''])
    selected_class = st.selectbox("Select Class:", ["Select Class"] + all_classes)

    if selected_class != "Select Class":
        class_students = df[df['CLASS'] == selected_class]
        all_sections = sorted([s for s in class_students['SECTION'].dropna().unique() if str(s).strip() != ''])
        selected_section = st.selectbox("Select Section:", ["All Sections"] + all_sections)

        if selected_section != "All Sections":
            filtered_df = class_students[class_students['SECTION'] == selected_section].copy()
        else:
            filtered_df = class_students.copy()

        total_students = len(filtered_df)

        # Merge with current printed status
        if not print_df.empty:
            merged_df = pd.merge(filtered_df, print_df, on='Adm. No. Clean', how='left')
            merged_df['Is_Printed'] = merged_df['Is_Printed'].fillna(False).astype(bool)
        else:
            merged_df = filtered_df.copy()
            merged_df['Is_Printed'] = False

        printed_count = merged_df['Is_Printed'].sum()
        percentage = (printed_count / total_students * 100) if total_students > 0 else 0

        st.markdown(f"### 📊 Printing Progress: **{printed_count} / {total_students}** Printed ({percentage:.1f}%)")
        st.progress(printed_count / total_students if total_students > 0 else 0.0)

        st.divider()

        # Display Students with Small Photos and Unique Key Checkbox
        updated_status = False

        for idx, s_row in merged_df.reset_index(drop=True).iterrows():
            adm_no = str(s_row['Adm. No. Clean'])
            photo_code = str(s_row.get('photo', '')).strip() if pd.notna(s_row.get('photo')) else ""
            is_printed = bool(s_row['Is_Printed'])

            c_photo, c_details, c_check = st.columns([1, 4, 1.5])

            with c_photo:
                photo_found = False
                if photo_code:
                    for ext in ['.jpg', '.JPG', '.jpeg', '.png']:
                        p_path = os.path.join(PHOTOS_DIR, f"{photo_code}{ext}")
                        if os.path.exists(p_path):
                            st.image(p_path, width=70)
                            photo_found = True
                            break
                if not photo_found:
                    st.caption("No Photo")

            with c_details:
                st.markdown(f"**{s_row.get('Student Name', '')}** (Adm No: `{adm_no}`) | Sr No: {s_row.get('Sr No', 'N/A')}")
                st.caption(f"Father: {s_row.get('Father\'s Name', '')} | DOB: {s_row.get('DOB', '')} | Ph: {s_row.get('Phone No.', '')}")

            with c_check:
                # Unique key using Adm No + Index to prevent Duplicate Key Error
                new_val = st.checkbox("Printed ✅", value=is_printed, key=f"chk_{adm_no}_{idx}")
                if new_val != is_printed:
                    if adm_no in print_df['Adm. No. Clean'].values:
                        print_df.loc[print_df['Adm. No. Clean'] == adm_no, 'Is_Printed'] = new_val
                        print_df.loc[print_df['Adm. No. Clean'] == adm_no, 'Updated_At'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    else:
                        new_rec = pd.DataFrame([{
                            'Adm. No. Clean': adm_no,
                            'Is_Printed': new_val,
                            'Updated_At': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }])
                        print_df = pd.concat([print_df, new_rec], ignore_index=True)
                    updated_status = True

            st.markdown("<hr style='margin: 5px 0;'>", unsafe_allow_html=True)

        if updated_status:
            save_print_status(print_df)
            st.rerun()

# ---------------------------------------------------------
# 3. ADMIN PANEL
# ---------------------------------------------------------
elif menu == "Admin Panel":
    st.subheader("🔒 Admin Dashboard")
    password = st.text_input("Enter Admin Password:", type="password")
    
    if password == "admin123":
        st.success("Welcome Admin!")
        
        tab1, tab2, tab3 = st.tabs(["📊 Corrections List", "➕ Add New Student Data", "🖨️ Class-wise Print Report"])
        
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
            st.info("Yahan se naya student data add karne par wo '1.xlsx' aur 'Corrections List' dono me save ho jayega.")
            
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
                        saved_photo_filename = ""
                        has_new_photo = 'No'
                        
                        if uploaded_photo_file is not None:
                            os.makedirs(PHOTOS_DIR, exist_ok=True)
                            saved_photo_filename = f"{final_photo_code}.jpg"
                            photo_save_path = os.path.join(PHOTOS_DIR, saved_photo_filename)
                            process_and_save_image(uploaded_photo_file, photo_save_path, max_size_mb=1.0)
                            has_new_photo = 'Yes'
                        
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
                        except Exception as e:
                            st.error(f"Excel file update karne me error aaya: {e}")
                        
                        correction_entry = {
                            'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            'Adm. No.': add_adm_no.strip(),
                            'Student Name': add_name.strip().replace(',', ' '),
                            'Changed Fields': "New Student Added via Admin Panel",
                            'DOB': str(add_dob).strip().replace(',', ' '),
                            'Father Name': str(add_father).strip().replace(',', ' '),
                            'Mother Name': str(add_mother).strip().replace(',', ' '),
                            'Colour': str(add_colour).strip().replace(',', ' '),
                            'Phone No.': str(add_phone).strip().replace(',', ' '),
                            'Address': str(add_address).strip().replace('\n', ' ').replace(',', ' '),
                            'Photo Code': final_photo_code,
                            'New Photo Uploaded': has_new_photo,
                            'Saved Photo Filename': saved_photo_filename
                        }
                        
                        corr_df = pd.DataFrame([correction_entry])
                        if not os.path.exists(CORRECTIONS_FILE):
                            corr_df.to_csv(CORRECTIONS_FILE, index=False)
                        else:
                            corr_df.to_csv(CORRECTIONS_FILE, mode='a', header=False, index=False)
                            
                        st.cache_data.clear()
                        st.success(f"✅ Student **{add_name.strip()}** Excel database aur Corrections List dono me add ho gaya hai!")

        with tab3:
            st.subheader("🖨️ Class-wise Printing Summary Report")
            print_df = load_print_status()

            if not print_df.empty:
                full_merged = pd.merge(df, print_df, on='Adm. No. Clean', how='left')
                full_merged['Is_Printed'] = full_merged['Is_Printed'].fillna(False).astype(bool)

                summary = full_merged.groupby(['CLASS', 'SECTION']).agg(
                    Total_Students=('Adm. No. Clean', 'count'),
                    Printed_Count=('Is_Printed', 'sum')
                ).reset_index()

                summary['Pending_Count'] = summary['Total_Students'] - summary['Printed_Count']
                summary['Progress (%)'] = (summary['Printed_Count'] / summary['Total_Students'] * 100).round(1)

                st.dataframe(summary, use_container_width=True)

                buf_rep = io.BytesIO()
                with pd.ExcelWriter(buf_rep, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='Print_Summary')

                st.download_button(
                    label="📥 Download Class-wise Print Summary Excel",
                    data=buf_rep.getvalue(),
                    file_name="classwise_print_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Abhi tak teachers ne kisi record ko tick mark nahi kiya hai.")
