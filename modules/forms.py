from flask import render_template

def render_form():
    return render_template('index.html')

def render_table(item_df, output_file_path):
    item_df_html = item_df.to_html(classes='table table-striped')
    return render_template('index.html', tables=[item_df_html], output_file_path=output_file_path)
