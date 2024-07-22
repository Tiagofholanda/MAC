import streamlit as st
import pandas as pd
import os
from datetime import datetime
from fpdf import FPDF
from spellchecker import SpellChecker # type: ignore

# Analysis functions for Fill, Types, Spelling, Names, Ranges, Conditionals
def analyze_fill(df_bd, df_matrix):
    result = []
    for column in range(1, len(df_matrix.columns)):
        text = df_matrix.columns[column]
        rule = df_matrix.iloc[4, column]
        if text in df_bd.columns:
            for row in range(len(df_bd)):
                cell_value = df_bd.loc[row, text]
                if (rule == "REQUIRED" and pd.isna(cell_value)) or (rule == "EMPTY" and not pd.isna(cell_value)):
                    reason = "Required fill" if rule == "REQUIRED" else "No fill"
                    result.append({
                        'Date': pd.Timestamp.now(),
                        'Row': row+1,
                        'Column': text,
                        'Value': cell_value,
                        'Analysis': 'Fill'
                    })
    return pd.DataFrame(result)

def analyze_types(df_bd, df_matrix):
    results = []
    for column in range(2, len(df_matrix.columns)):
        text = df_matrix.columns[column] 
        parameter = df_matrix.iloc[5, column]
        if text in df_bd.columns:
            if parameter == "TEXT ABC":
                alpha_filter = df_bd[text].apply(lambda x: str(x).replace(" ", "").isalpha())
                inconsistencies = df_bd[~alpha_filter]
            elif parameter == "DATE":
                date_filter = pd.to_datetime(df_bd[text], errors='coerce').isna()
                inconsistencies = df_bd[date_filter]
            elif parameter == "NUMBER":
                numeric_filter = pd.to_numeric(df_bd[text], errors='coerce').isna()
                inconsistencies = df_bd[numeric_filter]
            else:
                inconsistencies = pd.DataFrame()
            for index, row in inconsistencies.iterrows():
                results.append({
                    'Date': pd.Timestamp.now(),
                    'Row': index+1,
                    'Column': text,
                    'Value': row[text],
                    'Analysis': 'Type'
                })
    return pd.DataFrame(results)

def main_check_spelling(df_matrix, df_bd):
    spelling_errors = []
    for column in df_matrix.iloc[6].dropna().index:
        if df_matrix.iloc[6, df_matrix.columns.get_loc(column)] == "SPELLING DICTIONARY":
            spelling_errors += check_spelling(column, df_bd)
    return pd.DataFrame(spelling_errors)

def check_spelling(column_name, df_bd):
    spell = SpellChecker(language='en')
    spelling_errors = []

    for index, value in df_bd[column_name].items():
        if pd.notnull(value):
            words = str(value).split()
            for word in words:
                if spell.correction(word) != word:
                    spelling_errors.append({
                        'Date': datetime.now(),
                        'Row': index+1,
                        'Column': column_name,
                        'Value': value,
                        'Analysis': 'Spelling'
                    })
    return spelling_errors

def check_names(df_matrix, df_bd, df_names):
    found_errors = []
    for column in df_matrix.iloc[7].dropna().index:
        if df_matrix.iloc[7, df_matrix.columns.get_loc(column)] == "NAMES DATABASE":
            for index, value in enumerate(df_bd[column]):
                if pd.notnull(value):
                    name_parts = value.split()
                    for name_part in name_parts:
                        if not is_valid_name(name_part, df_names):
                            found_errors.append({
                                'Date': datetime.now(),
                                'Row': index + 1,
                                'Column': column,
                                'Value': value,
                                'Analysis': 'Names'
                            })
    return pd.DataFrame(found_errors)

def is_valid_name(name, df_names):
    return name in df_names.iloc[:, 0].values

def check_ranges(df_matrix, df_bd, df_ranges):
    df_bd1 = df_bd.fillna(value='')
    inconsistencies = []
    row = df_matrix.iloc[8]
    for col in df_matrix.columns:
        if 'range' in str(row[col]):
            range_str = str(row[col])
            if col in df_bd1.columns:
                for idx, value_bd in enumerate(df_bd1[col]):
                    try:
                        float(value_bd)
                        if value_bd not in df_ranges[range_str]:
                            analysis_time = datetime.now()
                            inconsistencies.append({
                                'Date': analysis_time,
                                'Row': idx + 1,
                                'Column': col,
                                'Value': value_bd,
                                'Analysis': 'Ranges'
                            })
                    except:
                        if str(value_bd) not in str(df_ranges[range_str]):
                            analysis_time = datetime.now()
                            inconsistencies.append({
                                'Date': analysis_time,
                                'Row': idx + 1,
                                'Column': col,
                                'Value': value_bd,
                                'Analysis': 'Ranges'
                            })
    return pd.DataFrame(inconsistencies)

def check_conditionals(df_matrix, df_bd):
    inconsistencies = []
    df_matrix1 = df_matrix.fillna(value='')
    for row in range(15, 92, 4):
        matrix_row = df_matrix1.iloc[row]
        for col in range(3, len(matrix_row)):
            value = matrix_row.iloc[col]
            if value != "":
                condition_column = value
                conditional_column = df_matrix.columns[col]
                conditional_value = df_matrix1.iloc[row + 1, col]
                result = df_matrix1.iloc[row + 2, col]
                if condition_column in df_bd.columns:
                    for index, value_bd in enumerate(df_bd[condition_column]):
                        if value_bd == conditional_value:
                            if conditional_column in df_bd.columns:
                                if df_bd[conditional_column][index] != result:
                                    inconsistencies.append({
                                        'Date': datetime.now(),
                                        'Row': index + 1,
                                        'Column': conditional_column,
                                        'Value': value_bd,
                                        'Analysis': 'Conditionals'
                                    })
                else:
                    st.write(f'Conditional column "{condition_column}" not found in df_bd.')
    return pd.DataFrame(inconsistencies)

# Function to select database file
def select_database_file():
    database_file_path = st.file_uploader("Select the Excel database file", type=["xlsx"])
    if database_file_path:
        try:
            global df_bd
            df_bd = pd.read_excel(database_file_path)
            st.success("Database file loaded successfully.")
        except Exception as e:
            st.error(f"An error occurred while loading the database file: {str(e)}")

# Function to select matrix file
def select_matrix_file():
    matrix_file_path = st.file_uploader("Select the Excel matrix file", type=["xlsx"])
    if matrix_file_path:
        try:
            global df_matrix
            df_matrix = pd.read_excel(matrix_file_path, header=1)
            st.success("Matrix file loaded successfully.")
        except Exception as e:
            st.error(f"An error occurred while loading the matrix file: {str(e)}")

# Function to select names file
def select_names_file():
    names_file_path = st.file_uploader("Select the Excel names file", type=["xlsx"])
    if names_file_path:
        try:
            global df_names
            df_names = pd.read_excel(names_file_path)
            st.success("Names file loaded successfully.")
        except Exception as e:
            st.error(f"An error occurred while loading the names file: {str(e)}")

# Function to select ranges file
def select_ranges_file():
    ranges_file_path = st.file_uploader("Select the Excel ranges file", type=["xlsx"])
    if ranges_file_path:
        try:
            global df_ranges
            df_ranges = pd.read_excel(ranges_file_path)
            st.success("Ranges file loaded successfully.")
        except Exception as e:
            st.error(f"An error occurred while loading the ranges file: {str(e)}")

# Function to start fill analysis
def start_fill_analysis():
    if 'df_bd' in globals() and 'df_matrix' in globals():
        result = analyze_fill(df_bd, df_matrix)
        st.write(result)
    else:
        st.warning('Select the database and matrix files before starting the analysis.')

# Function to start type analysis
def start_type_analysis():
    if 'df_bd' in globals() and 'df_matrix' in globals():
        result = analyze_types(df_bd, df_matrix)
        st.write(result)
    else:
        st.warning('Select the database and matrix files before starting the analysis.')

# Function to start spelling check
def start_spelling_check():
    if 'df_bd' in globals() and 'df_matrix' in globals():
        result = main_check_spelling(df_matrix, df_bd)
        st.write(result)
    else:
        st.warning('Select the database and matrix files before starting the check.')

# Function to start names check
def start_names_check():
    if 'df_bd' in globals() and 'df_matrix' in globals() and 'df_names' in globals():
        result = check_names(df_matrix, df_bd, df_names)
        st.write(result)
    else:
        st.warning('Select the database, matrix, and names files before starting the check.')

# Function to start ranges check
def start_ranges_check():
    if 'df_bd' in globals() and 'df_matrix' in globals() and 'df_ranges' in globals():
        result = check_ranges(df_matrix, df_bd, df_ranges)
        st.write(result)
    else:
        st.warning('Select the database, matrix, and ranges files before starting the check.')

# Function to start conditionals check
def start_conditionals_check():
    if 'df_bd' in globals() and 'df_matrix' in globals():
        result = check_conditionals(df_matrix, df_bd)
        st.write(result)
    else:
        st.warning('Select the database and matrix files before starting the check.')

# Function to concatenate all results into a report
def concatenate_results():
    global concatenated_result
    if 'df_bd' in globals() and 'df_matrix' in globals():
        result1 = analyze_fill(df_bd, df_matrix)
        result2 = analyze_types(df_bd, df_matrix)
        result3 = main_check_spelling(df_matrix, df_bd)
        result4 = check_names(df_matrix, df_bd, df_names) if 'df_names' in globals() else pd.DataFrame()
        result5 = check_ranges(df_matrix, df_bd, df_ranges) if 'df_ranges' in globals() else pd.DataFrame()
        result6 = check_conditionals(df_matrix, df_bd)
        concatenated_result = pd.concat([result1, result2, result3, result4, result5, result6], ignore_index=True)
        st.write(concatenated_result.head(20))
    else:
        st.warning('Select all necessary files before running the checks.')

# Function to export results to PDF
def export_results_to_pdf():
    global concatenated_result

    if concatenated_result is not None and not concatenated_result.empty:
        try:
            # Class to generate the PDF
            class PDF(FPDF):
                def header(self):
                    if hasattr(self, 'logo_path'):
                        self.image(self.logo_path, 10, 8, 33)
                    self.set_font('Arial', 'B', 12)
                    self.cell(0, 10, 'Inconsistencies Report', 0, 1, 'C')
                    self.ln(10)

                def footer(self):
                    self.set_y(-30)
                    self.set_font('Arial', 'I', 8)
                    self.cell(0, 10, 'Source: GIS NMC Team: Pedro Reis, Carolina Peres, Melquisedeque Nunes, Tiago Holanda', 0, 1, 'C')
                    self.set_y(-15)
                    self.cell(0, 10, 'Page %s' % self.page_no(), 0, 0, 'C')

                def chapter_title(self, title):
                    self.set_font('Arial', 'B', 16)
                    self.cell(0, 10, title, 0, 1, 'C')
                    self.ln(10)

                def chapter_body(self, body):
                    self.set_font('Arial', '', 12)
                    self.multi_cell(0, 10, body)
                    self.ln()

                def table_row(self, row):
                    self.set_font('Arial', '', 12)
                    for item in row:
                        self.cell(0, 10, f'{item}', border=1, ln=0)
                    self.ln(10)

            logo_path = "path_to_logo"  # Update with the path to the desired logo

            # Generating a PDF for each row
            for index, row in concatenated_result.iterrows():
                pdf = PDF()
                pdf.logo_path = logo_path
                pdf.add_page()
                pdf.chapter_title(f'Results for Row {index + 1}')
                
                for column, value in row.items():
                    pdf.set_font('Arial', 'B', 12)
                    column_width = pdf.get_string_width(f'{column}:') + 10
                    pdf.cell(column_width, 10, f'{column}:', border=1)
                    pdf.set_font('Arial', '', 12)
                    pdf.multi_cell(0, 10, f'{value}', border=1)
                    pdf.ln(5)

                output_dir = "PDFs"
                os.makedirs(output_dir, exist_ok=True)
                output_path = os.path.join(output_dir, f'Row_{index + 1}.pdf')
                pdf.output(output_path)

            # Read the PDF file and allow download
            with open(output_path, "rb") as pdf_file:
                PDFbyte = pdf_file.read()

            st.download_button(label="Download PDF",
                               data=PDFbyte,
                               file_name="inconsistencies_report.pdf",
                               mime='application/octet-stream')

            st.success('PDFs generated successfully!')
        except Exception as e:
            st.error(f'An error occurred while exporting the data:\n{str(e)}')
    else:
        st.warning('No data to export.')

# Streamlit interface
st.title('Inconsistencies Checker - NMC')

# Load files
st.sidebar.header('Import Files')
select_database_file()
select_matrix_file()
select_names_file()
select_ranges_file()

# Run analyses
st.sidebar.header('Run Analyses')
if st.sidebar.button('Fill Analysis'):
    start_fill_analysis()
if st.sidebar.button('Type Analysis'):
    start_type_analysis()
if st.sidebar.button('Spelling Check'):
    start_spelling_check()
if st.sidebar.button('Names Check'):
    start_names_check()
if st.sidebar.button('Ranges Check'):
    start_ranges_check()
if st.sidebar.button('Conditionals Check'):
    start_conditionals_check()
if st.sidebar.button('Run Checks'):
    concatenate_results()

# Button to export results to PDF
if st.sidebar.button('Export Results to PDF'):
    export_results_to_pdf()
