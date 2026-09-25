import streamlit as st
import docx
import pandas as pd
import io

# 1. Parsing Logic
def parse_colored_word_doc(file_obj, exam_id):
    doc = docx.Document(file_obj)
    parsed_data = []
    current_question = None
    current_options = []
    correct_option_value = None
    parsing_started = False 

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
        if "Part-1:" in text:
            parsing_started = True
            continue
        if not parsing_started:
            continue

        # Detect red text for correct answer
        is_red = False
        for run in para.runs:
            if run.font.color and run.font.color.rgb:
                if str(run.font.color.rgb) == 'FF0000': 
                    is_red = True
                    break
        
        # Group into 4 options per question
        if len(current_options) == 4 or (current_question is None):
            if current_question and current_options:
                parsed_data.append({
                    'q_id': f"{exam_id}_{len(parsed_data) + 1}",
                    'text': current_question,
                    'options': current_options,
                    'correct': correct_option_value
                })
            current_question = text
            current_options = []
            correct_option_value = None
        else:
            current_options.append(text)
            if is_red:
                correct_option_value = text 

    if current_question and current_options:
        parsed_data.append({
            'q_id': f"{exam_id}_{len(parsed_data) + 1}",
            'text': current_question,
            'options': current_options,
            'correct': correct_option_value
        })

    return parsed_data

# 2. Excel Generation Logic
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
    
    # Write to an in-memory buffer instead of a physical file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_new.to_excel(writer, sheet_name='Question', index=False)
        pd.DataFrame().to_excel(writer, sheet_name='Feedback', index=False)
        pd.DataFrame().to_excel(writer, sheet_name='Objectives', index=False)
    
    output.seek(0)
    return output

# 3. Web UI (Streamlit)
st.set_page_config(page_title="LMS Question Converter", layout="centered")

st.title("📄 Word to SAP SF LMS Converter")
st.write("Upload your formatted Word document to generate the LMS import file.")

# Input fields
exam_id = st.text_input("Enter Exam ID (e.g., MCT_2023)", "MCT_2023")
domain_id = st.text_input("Enter LMS Domain ID", "YOUR_DOMAIN")
uploaded_word = st.file_uploader("Upload Word Document (.docx)", type=["docx"])

# Processing trigger
if st.button("Convert Document"):
    if uploaded_word is not None and exam_id:
        with st.spinner("Parsing document and generating Excel..."):
            try:
                # Process the file
                questions = parse_colored_word_doc(uploaded_word, exam_id)
                
                if not questions:
                    st.error("No questions found. Please check the document format.")
                else:
                    # Generate Excel buffer
                    excel_data = generate_lms_excel(questions, domain_id)
                    
                    st.success(f"Successfully processed {len(questions)} questions!")
                    
                    # Provide download button
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
