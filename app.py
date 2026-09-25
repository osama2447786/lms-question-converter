import streamlit as st
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import pandas as pd
import io

# --- Helper functions to parse Word structure ---
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

def extract_colored_lines(block):
    """Extracts text line-by-line from a block, detecting red text formatting."""
    lines = []
    current_line_text = ""
    current_line_red = False
    
    for run in block.runs:
        run_text = run.text
        is_red = False
        # Detect Red text (hex code FF0000)
        if run.font.color and run.font.color.rgb and str(run.font.color.rgb) == 'FF0000':
            is_red = True
            
        parts = run_text.split('\n')
        for idx, part in enumerate(parts):
            current_line_text += part
            if is_red and part.strip():
                current_line_red = True
                
            # If a newline occurs, save the collected line and reset
            if idx < len(parts) - 1: 
                if current_line_text.strip():
                    lines.append({'text': current_line_text.strip(), 'is_red': current_line_red})
                current_line_text = ""
                current_line_red = False
                
    # Catch any trailing text
    if current_line_text.strip():
        lines.append({'text': current_line_text.strip(), 'is_red': current_line_red})
        
    return lines

# --- Core Parsing Logic ---
def parse_with_placeholders(file_obj, exam_id):
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
            if not block.text.strip():
                i += 1
                continue
                
            has_image = 'w:drawing' in block._element.xml or 'v:imagedata' in block._element.xml
            
            # Extract all lines (this fixes the soft return / \n issue)
            para_lines = extract_colored_lines(block)
            
            # SCENARIO 1: Question and Options are bundled in the SAME paragraph block
            if len(para_lines) >= 5:
                current_question = para_lines[0]['text']
                options = [line['text'] for line in para_lines[1:]]
                
                # Find the red option, or default to the first one
                correct_option = next((line['text'] for line in para_lines[1:] if line['is_red']), options[0])
                
                if has_image:
                    parsed_data.append({
                        'q_id': f"{exam_id}_{len(parsed_data) + 1}",
                        'text': f"[IMAGE REQUIRED - UPDATE MANUALLY] {current_question}",
                        'options': ["Option A", "Option B", "Option C", "Option D"],
                        'correct': "Option A"
                    })
                else:
                    parsed_data.append({
                        'q_id': f"{exam_id}_{len(parsed_data) + 1}",
                        'text': current_question,
                        'options': options,
                        'correct': correct_option
                    })
                i += 1
                continue
                
            # SCENARIO 2: Question is one block, options are in the following blocks or tables
            current_question = para_lines[0]['text'] if para_lines else block.text.strip()
            options = []
            correct_option = None
            skip_question = has_image
            
            i += 1
            # Look ahead for options
            while i < len(blocks):
                next_block = blocks[i]
                
                if isinstance(next_block, Paragraph):
                    if not next_block.text.strip():
                        i += 1
                        continue
                        
                    # Stop looking if we already collected 4 options
                    if len(options) >= 4:
                        break
                        
                    if 'w:drawing' in next_block._element.xml or 'v:imagedata' in next_block._element.xml:
                        skip_question = True
                        
                    nxt_lines = extract_colored_lines(next_block)
                    for nl in nxt_lines:
                        options.append(nl['text'])
                        if nl['is_red']:
                            correct_option = nl['text']
                    i += 1
                    
                elif isinstance(next_block, Table):
                    if 'w:drawing' in next_block._element.xml or 'v:imagedata' in next_block._element.xml:
                        skip_question = True
                        
                    for row in next_block.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                cell_lines = extract_colored_lines(p)
                                for cl in cell_lines:
                                    options.append(cl['text'])
                                    if cl['is_red']:
                                        correct_option = cl['text']
                    i += 1
                    break # Tables usually hold all the options, so stop looking after parsing it
                    
            # Append gathered data
            if skip_question:
                parsed_data.append({
                    'q_id': f"{exam_id}_{len(parsed_data) + 1}",
                    'text': f"[IMAGE REQUIRED - UPDATE MANUALLY] {current_question}",
                    'options': ["Placeholder A", "Placeholder B", "Placeholder C", "Placeholder D"],
                    'correct': "Placeholder A"
                })
            elif len(options) >= 2:
                if not correct_option: 
                    correct_option = options[0] # Fallback if no red text is found
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
st.write("Upload your formatted Word document to generate the LMS import file. Questions containing images will be generated as placeholders to preserve numbering.")

exam_id = st.text_input("Enter Exam ID (e.g., MCT_2023)", "MCT_2023")
domain_id = st.text_input("Enter LMS Domain ID", "YOUR_DOMAIN")
uploaded_word = st.file_uploader("Upload Word Document (.docx)", type=["docx"])

if st.button("Convert Document"):
    if uploaded_word is not None and exam_id:
        with st.spinner("Parsing document..."):
            try:
                questions = parse_with_placeholders(uploaded_word, exam_id)
                
                if not questions:
                    st.error("No valid questions found. Please check the document format.")
                else:
                    excel_data = generate_lms_excel(questions, domain_id)
                    st.success(f"Successfully processed {len(questions)} questions!")
                    
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
