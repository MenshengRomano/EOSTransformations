from flask import Flask, render_template, request, send_file
import os
import shutil
import pandas as pd
from modules.forms import render_form
from modules.template_reader import extract_tables, load_template, load_mechanical_template
from modules.data_processor import process_data
from modules.mechanical_processor import process_bid_summary, process_kpis, process_bid_info

app = Flask(__name__)


@app.route('/')
def index():
    return render_form()


@app.route('/process', methods=['POST'])
def process():
    source_file = request.files['source_file']
    template_file = request.files['template_file']
    estimate_type = request.form["estimate_type"]

    if not source_file or not template_file:
        return "Please upload both source and template files."

    # Save the uploaded files to a temporary location
    source_file_path = os.path.join('uploads', source_file.filename)
    template_file_path = os.path.join('uploads', template_file.filename)
    source_file.save(source_file_path)
    template_file.save(template_file_path)

    # Copy the template file to preserve the original
    output_file_path = os.path.join('outputs', f"{source_file.filename.split('.')[0]} (Cortex import).xlsx")
    shutil.copy(template_file_path, output_file_path)

    # Check the estimate type
    if estimate_type == "electrical":
        try:
            source_df = pd.read_excel(source_file_path, sheet_name="Extension")
            wb, mapping_ws, item_ws = load_template(output_file_path)
            tables = extract_tables(mapping_ws)

            item_df, mapping_df = process_data(source_df, tables, item_ws)

            # Save the updated workbook
            wb.save(output_file_path)
            print(f"Saved updated workbook to {output_file_path}")

            item_df_html = item_df.to_html(classes='table table-striped')
            mapping_df_html = mapping_df.to_html(classes='table table-striped')

            return render_template('index.html', tables=[item_df_html, mapping_df_html], output_file_path=output_file_path)

        except Exception as e:
            return f"An error occurred: {e}"
        
    elif estimate_type == "mechanical":
        try:
            bid_summary_df = pd.read_excel(source_file_path, sheet_name="1-Bid Summary")
            bid_summary_df.columns = [chr(col + 97) for col in range(len(bid_summary_df.columns))]
            wb, project_ws, item_ws = load_mechanical_template(output_file_path)
            process_bid_summary(bid_summary_df, item_ws)
            
            bid_analysis_kpi_df = pd.read_excel(source_file_path, sheet_name="Bid Analysis KPI")
            bid_analysis_kpi_df.columns = [chr(col + 97) for col in range(len(bid_analysis_kpi_df.columns))]
            process_kpis(project_ws, bid_analysis_kpi_df)
            
            bid_info_df = pd.read_excel(source_file_path, sheet_name="0-Bid Info")
            bid_info_df.columns = [chr(col + 97) for col in range(len(bid_info_df.columns))]
            process_bid_info(project_ws, bid_info_df)
            source_file = source_file_path.split('\\')[-1]
            project_ws.cell(row=42, column=4, value=source_file)

            wb.save(output_file_path)
            print(f"Saved updated workbook to {output_file_path}")
            return "Mechanical processing is not implemented yet."
        except Exception as e:
            return f"An error occurred: {e}"
    # no else because a selection is required

@app.route('/download/<path:filename>', methods=['GET'])
def download(filename):
    return send_file(filename, as_attachment=True)


if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    if not os.path.exists('outputs'):
        os.makedirs('outputs')
    app.run(debug=True)
