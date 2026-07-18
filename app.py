import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import base64

# Page configuration for a clean enterprise feel
st.set_page_config(page_title="DOD FastCash Engine", layout="centered")

st.title("DOD FastCash Engine")
st.caption("FastCash Upload Files Generator")
st.markdown("---")

# Form inputs wrapped beautifully
with st.form("engine_form"):
    st.subheader("Configuration Parameters")
    
    col1, col2 = st.columns(2)
    with col1:
        contra_acc = st.text_input("Contra Account", placeholder="Alphanumeric account code")
        tran_type = st.selectbox("Tran_Type (D/C)", ["D", "C"], index=0)
        contra_amt = st.number_input("Contra Amount", min_value=0.0, step=0.01, format="%.2f")
    
    with col2:
        out_sheet_name = st.text_input("Output Sheet Name", value="FastCash_Batch")
        ref_prefix = st.text_input("Ref. Prefix", placeholder="e.g., FAST//")
        contra_narr = st.text_input("Contra Narration", max_chars=40, placeholder="Max 40 chars")
        
    inc_pos = st.checkbox("Include cell position details (e.g., E12)", value=True)
    
    st.markdown("---")
    st.subheader("File Ingestion")
    # Streamlit natively displays the uploaded file name cleanly right underneath!
    uploaded_file = st.file_uploader("Upload Excel Template", type=["xlsx", "xls", "xlsb"])
    
    submit_btn = st.form_submit_button("Process Data & Generate Output", use_container_width=True)

# Processing Logic execution
if submit_btn:
    if not uploaded_file:
        st.error("❌ No file detected. Please upload an Excel file first.")
    elif len(contra_narr) > 40:
        st.error(f"❌ Contra Narration exceeds 40 characters (Currently {len(contra_narr)}).")
    else:
        with st.spinner("⏳ Processing file... Please wait."):
            try:
                file_name = uploaded_file.name.lower()
                file_bytes = uploaded_file.read()
                
                if file_name.endswith('.xlsb'):
                    df_in = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, engine='pyxlsb')
                else:
                    df_in = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
                
                df_in['_excel_row_idx'] = df_in.index + 2
                col_e_name = df_in.columns[4]
                df_in[col_e_name] = pd.to_numeric(df_in[col_e_name], errors='coerce')
                df_filtered = df_in.dropna(subset=[col_e_name]).copy()
                
                chunks = []
                current_chunk = []
                current_rows = 0
                current_sum = 0.0
                
                for _, row in df_filtered.iterrows():
                    val_e = float(row[col_e_name])
                    if (current_rows + 1 > 7999) or (current_sum + val_e >= 500000000):
                        if current_chunk:
                            chunks.append(pd.DataFrame(current_chunk))
                        current_chunk = [row]
                        current_rows = 1
                        current_sum = val_e
                    else:
                        current_chunk.append(row)
                        current_rows += 1
                        current_sum += val_e
                        
                if current_chunk:
                    chunks.append(pd.DataFrame(current_chunk))
                    
                if not chunks:
                    st.error("❌ No valid rows found containing numbers in Column E for processing.")
                else:
                    out_buffer = io.BytesIO()
                    wb = openpyxl.Workbook()
                    wb.remove(wb.active)
                    
                    summary_data = []
                    all_contra_rows_sum = 0.0
                    
                    for idx, chunk_df in enumerate(chunks, 1):
                        sheet_title = f"{out_sheet_name} {idx}"
                        ws = wb.create_sheet(title=sheet_title)
                        
                        processed_rows = []
                        chunk_sum = 0.0
                        
                        for _, r in chunk_df.iterrows():
                            val_a = str(r.iloc[0]) if pd.notnull(r.iloc[0]) else ""
                            val_b = r.iloc[1] if pd.notnull(r.iloc[1]) else ""
                            val_e = float(r[col_e_name])
                            val_f = r.iloc[5] if pd.notnull(r.iloc[5]) else ""
                            val_g = r.iloc[6] if pd.notnull(r.iloc[6]) else ""
                            
                            letter_a = val_a[0].upper() if val_a else ""
                            rounded_e = round(val_e, 2)
                            chunk_sum += rounded_e
                            
                            ref_col_f = f"{ref_prefix}E{r['_excel_row_idx']}" if inc_pos else ref_prefix
                            
                            processed_rows.append([val_b, letter_a, rounded_e, val_f, "", ref_col_f, "", "", "", val_g])
                            
                        for row_data in processed_rows:
                            ws.append(row_data)
                            
                        final_chunk_sum = round(chunk_sum, 2)
                        all_contra_rows_sum += final_chunk_sum
                        
                        contra_row = [contra_acc, tran_type, final_chunk_sum, contra_narr, "", f"{ref_prefix}{sheet_title}", "", "", "", ""]
                        ws.append(contra_row)
                        
                        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=10):
                            row[2].number_format = '#,##0.00'
                            
                        last_row_idx = ws.max_row
                        accent_fill = PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid")
                        for col_idx in range(1, 11):
                            cell = ws.cell(row=last_row_idx, column=col_idx)
                            cell.font = Font(bold=True, color="1A365D")
                            cell.fill = accent_fill
                            
                        summary_data.append({"sheet": sheet_title, "imported_sum": final_chunk_sum, "contra_sum": final_chunk_sum, "variance": 0.0})
                    
                    # Executive Summary Page
                    ws_sum = wb.create_sheet(title="Executive Summary", index=0)
                    ws_sum.views.sheetView[0].showGridLines = True
                    
                    navy_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
                    accent_bar_fill = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
                    white_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
                    thin_border = Border(left=Side(style='thin', color='CBD5E0'), right=Side(style='thin', color='CBD5E0'), top=Side(style='thin', color='CBD5E0'), bottom=Side(style='thin', color='CBD5E0'))
                    
                    ws_sum.cell(row=2, column=2, value="DOD FastCash Engine").font = Font(name="Segoe UI", size=16, bold=True, color="1A365D")
                    ws_sum.cell(row=3, column=2, value="Executive Processing Summary Report").font = Font(name="Segoe UI", size=11, italic=True, color="4A5568")
                    
                    ws_sum.cell(row=5, column=2, value="Global Reconciliation Table").font = Font(name="Segoe UI", size=11, bold=True)
                    for c_idx, h in enumerate(["Metric Parameter Description", "Value Data"], start=2):
                        cell = ws_sum.cell(row=6, column=c_idx, value=h)
                        cell.fill = navy_fill
                        cell.font = white_font
                        cell.alignment = Alignment(horizontal="center")
                        
                    metrics = [
                        ("Target Global Contra Amount Inputted", contra_amt),
                        ("Calculated Net Sum of Generated Rows", all_contra_rows_sum),
                        ("Net Overall Variance Status", round(all_contra_rows_sum - contra_amt, 2))
                    ]
                    
                    for r_idx, (m, v) in enumerate(metrics, start=7):
                        c1 = ws_sum.cell(row=r_idx, column=2, value=m)
                        c2 = ws_sum.cell(row=r_idx, column=3, value=v)
                        c1.font = Font(name="Segoe UI", size=11)
                        c2.font = Font(name="Segoe UI", size=11, bold=True)
                        c2.number_format = '#,##0.00'
                        c1.border = thin_border
                        c2.border = thin_border
                        if "Variance" in m:
                            c2.font = Font(name="Segoe UI", bold=True, color="38A169" if v == 0 else "E53E3E")
                            
                    # Breakdown Log
                    start_r = 12
                    ws_sum.cell(row=start_r, column=2, value="Batch Data Sheet Breakdown Variance Log").font = Font(name="Segoe UI", size=11, bold=True)
                    for c_idx, h in enumerate(["Output Sheet Identifier", "Imported Rows Sum (Col C)", "Created Contra Row (Col C)", "Variance Verification"], start=2):
                        cell = ws_sum.cell(row=start_r+1, column=c_idx, value=h)
                        cell.fill = accent_bar_fill
                        cell.font = white_font
                        
                    curr_row = start_r + 2
                    for item in summary_data:
                        r_cells = [
                            ws_sum.cell(row=curr_row, column=2, value=item['sheet']),
                            ws_sum.cell(row=curr_row, column=3, value=item['imported_sum']),
                            ws_sum.cell(row=curr_row, column=4, value=item['contra_sum']),
                            ws_sum.cell(row=curr_row, column=5, value=item['variance'])
                        ]
                        for idx, c in enumerate(r_cells):
                            c.font = Font(name="Segoe UI", size=11, bold=(idx==3))
                            c.border = thin_border
                            if idx > 0:
                                c.number_format = '#,##0.00'
                        curr_row += 1
                        
                    for ws_obj in wb.worksheets:
                        for col in ws_obj.columns:
                            max_len = max(len(str(cell.value or '')) for cell in col)
                            col_letter = openpyxl.utils.get_column_letter(col[0].column)
                            ws_obj.column_dimensions[col_letter].width = max(max_len + 3, 12)
                            
                    wb.save(out_buffer)
                    out_buffer.seek(0)
                    
                    st.success("🎉 Execution Processing Completed Successfully!")
                    st.download_button(
                        label="📥 Click to Download Output File",
                        data=out_buffer.getvalue(),
                        file_name=f"{out_sheet_name}_Processed.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"❌ Processing Error: {str(e)}")