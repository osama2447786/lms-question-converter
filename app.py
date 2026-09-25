import streamlit as st
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import pandas as pd
import io

# --- Helper function to read paragraphs and tables in exact document order ---
def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Invalid parent object")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# --- Updated Parsing Logic ---
def parse_and_skip_images(file_obj, exam_id):
    doc = docx.Document(file_obj)
    blocks = list(iter_block_items(doc))
    
    parsed_data = []
    i = 0
    parsing_started = False
    
    while i < len(blocks):
        block = blocks[i]
        
        # Start parsing when Part 1 begins
        if isinstance(block, Paragraph) and "Part-1:" in block.text:
            parsing_started = True
            i += 1
            continue
            
        if not parsing_started:
            i += 1
            continue
            
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                i += 1
                continue
                
            # Check if this paragraph contains an image
            has_image = 'w:drawing' in block._element.xml or 'v:imagedata' in block._element.xml
            
            current_question = text
            options = []
            correct_option = None
            skip_question = has_image
            
            i += 1
            
            # Check if the options are inside a Table (like Question 6)
            if i < len(blocks) and isinstance(blocks[i], Table):
                tbl_block = blocks[i]
                
                # Check if the table contains an image
                if 'w:drawing' in tbl_block._element.xml or 'v:imagedata' in tbl_block._element.xml:
                    skip_question = True
                    
                for row in tbl_block.rows:
                    for cell in row.cells:
                        c_text = cell.text.strip()
                        if c_text:
                            # Split by newline in case multiple options are in one cell
                            for line in c_text.split('\n'):
                                if line.strip():
                                    options.append(line.strip())
                                    # Check for red text formatting for correct answer
                                    for p in cell.paragraphs:
                                        for r in p.runs:
                                            if r.font.color and r.font.color.rgb:
                                                if str(r.font.color.rgb) == 'FF0000':
                                                    correct_option = line.strip()
                i += 1
            else:
                # The options are in subsequent paragraphs
                lines_collected = 0
                while i < len(blocks) and lines_collected < 4:
                    opt_block = blocks[i]
                    if isinstance(opt_block, Paragraph):
                        if 'w:drawing' in opt_block._element.xml or 'v:imagedata' in opt_block._element.xml:
                            skip_question = True
                            
                        opt_text = opt_block.text.strip()
                        if opt_text:
                            for line in opt_text.split('\n'):
                                line = line.strip()
                                if line:
                                    options.append(line)
                                    lines_collected += 1
                                    
                                    for r in opt_block.runs:
                                        if r.font.color and r.font.color.rgb:
                                            if str(r.font.color.rgb) == 'FF0000':
                                                correct_option = line
                    elif isinstance(opt_block, Table):
                        # Stop if we hit a table that isn't ours
                        break 
                    i += 1

            # Only append the question if it doesn't have an image and has options
            if not skip_question and len(options) >= 2:
                parsed_data.append({
                    'q_id': f"{exam_id}_{len(parsed_data) + 1}",
                    'text': current_question,
                    'options': options,
                    'correct': correct_option
                })
        else:
            i += 1
            
    return parsed_data

# --- Excel Generation Logic ---
def generate_lms_excel(parsed_data, domain_id):
    rows = []
    for q in parsed_data:
        for opt_idx, option_text in enumerate(q['options']):
            is_correct = "Y" if option_text == q['correct'] else "N"
            row = {
                'Question ID (*required)': q['q_id'],
                'Locale ID (*required)': "English",
                'Domain ID (*required)': domain_id,
                'Active': 'Y',
                'Point Value': 1,
                'Variant Number (*required)': 1,
                'Question Type (*required)': "Multiple Choice",
                'Question Text (*required)': q['text'],
                'Answer Choice Number (*required)': opt_idx + 1,
                'Answer Choice Value': option_text,
                'Is Correct (*required)': is_correct,
            }
            rows.append(row)
            
    df_new = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_new.to_excel(writer, sheet_name='Question', index=False)
        pd.DataFrame().to_excel(writer, sheet_name='Feedback', index=False)
        pd.DataFrame().to_excel(writer, sheet_name='Objectives', index=False)
    
    output.seek(0)
    return output

# --- Streamlit UI ---
st.set_page_config(page_title="LMS Question Converter", layout="centered")

st.title("📄 Word to SAP SF LMS Converter")
st.write("Upload your formatted Word document to generate the LMS import file. Questions containing images will be automatically skipped to prevent formatting errors.")

exam_id = st.text_input("Enter Exam ID (e.g., MCT_2023)", "MCT_2023")
domain_id = st.text_input("Enter LMS Domain ID", "YOUR_DOMAIN")
uploaded_word = st.file_uploader("Upload Word Document (.docx)", type=["docx"])

if st.button("Convert Document"):
    if uploaded_word is not None and exam_id:
        with st.spinner("Parsing document and skipping image questions..."):
            try:
                questions = parse_and_skip_images(uploaded_word, exam_id)
                
                if not questions:
                    st.error("No valid text-based questions found. Please check the document format.")
                else:
                    excel_data = generate_lms_excel(questions, domain_id)
                    st.success(f"Successfully processed {len(questions)} text-based questions! (Image questions were skipped)")
                    
                    st.download_button(
                        label="📥 Download LMS Excel File",
                        data=excel_data,
                        file_name=f"LMS_Import_{exam_id}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"An error occurred: {e}")
    else:
        st.warning("Please upload a Word document and provide an Exam ID.")
