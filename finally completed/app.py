from flask import Flask, request, send_file, render_template_string
import pandas as pd
import os

app = Flask(__name__)

# Directories to store uploaded and processed files
UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Excel Comparison Tool</title>
    </head>
    <body>
        <h1>Upload Two Excel Files for Comparison</h1>
        <form action="/compare" method="post" enctype="multipart/form-data">
            <label for="file1">File 1:</label>
            <input type="file" name="file1" required><br><br>
            <label for="file2">File 2:</label>
            <input type="file" name="file2" required><br><br>
            <button type="submit">Compare Files</button>
        </form>
    </body>
    </html>
    '''

def compare_excel_files(file1_path, file2_path):
    """
    Compares two Excel files based on 'AWB number' and 'Weight' columns,
    and returns the rows where both columns match or mismatch.
    """
    try:
        # Read the Excel files into pandas DataFrames
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)

        # Normalize column names (strip spaces and convert to lowercase for consistency)
        df1.columns = df1.columns.str.strip().str.lower()
        df2.columns = df2.columns.str.strip().str.lower()

        # Check if the necessary columns ('awb number' and 'weight') exist in both files
        if 'awb number' not in df1.columns or 'weight' not in df1.columns:
            raise ValueError("File 1 is missing required columns: 'AWB number' and/or 'Weight'.")
        if 'awb number' not in df2.columns or 'weight' not in df2.columns:
            raise ValueError("File 2 is missing required columns: 'AWB number' and/or 'Weight'.")

        # Filter relevant columns ('AWB number' and 'Weight') from both files
        df1_processed = df1[['awb number', 'weight']]
        df2_processed = df2[['awb number', 'weight']]

        # Merge both DataFrames on the 'AWB number' column
        merged_df = pd.merge(df1_processed, df2_processed, on='awb number', how='outer', indicator=True, suffixes=('_file1', '_file2'))

        # Calculate the absolute difference in weights between the two files
        merged_df['weight_diff'] = (merged_df['weight_file1'] - merged_df['weight_file2']).abs()

        # Separate matching and mismatching rows
        matching_rows = merged_df[merged_df['_merge'] == 'both'].drop(columns=['_merge'])
        mismatching_rows = merged_df[merged_df['_merge'] != 'both'].drop(columns=['_merge'])

        # Save the results to new Excel files
        matching_rows_path = os.path.join(PROCESSED_FOLDER, 'matching_awb_numbers.xlsx')
        mismatching_rows_path = os.path.join(PROCESSED_FOLDER, 'mismatching_awb_numbers.xlsx')
        matching_rows.to_excel(matching_rows_path, index=False)
        mismatching_rows.to_excel(mismatching_rows_path, index=False)

        return matching_rows, mismatching_rows, matching_rows_path, mismatching_rows_path

    except Exception as e:
        print(f"Error comparing the Excel files: {e}")
        return None, None, None, None


@app.route('/compare', methods=['POST'])
def compare_files():
    if 'file1' not in request.files or 'file2' not in request.files:
        return "Please upload both files.", 400

    # Get the uploaded files
    file1 = request.files['file1']
    file2 = request.files['file2']

    # Save the files temporarily
    file1_path = os.path.join(UPLOAD_FOLDER, file1.filename)
    file2_path = os.path.join(UPLOAD_FOLDER, file2.filename)
    file1.save(file1_path)
    file2.save(file2_path)

    # Compare the files
    matching_data, mismatching_data, matching_file, mismatching_file = compare_excel_files(file1_path, file2_path)

    # If files were processed successfully, display matching and mismatching data in tables
    if matching_data is not None and mismatching_data is not None:
        # Convert matching and mismatching rows to HTML tables
        matching_html = matching_data[['awb number', 'weight_file1', 'weight_file2', 'weight_diff']].to_html(classes='table table-bordered', index=False)
        mismatching_html = mismatching_data[['awb number', 'weight_file1', 'weight_file2', 'weight_diff']].to_html(classes='table table-bordered', index=False)

        return render_template_string('''
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>File Comparison Results</title>
        </head>
        <body>
            <h1>File Comparison Results</h1>

            <h2>Matching Rows</h2>
            <p>These rows have the same AWB number and weight in both files.</p>
            {{ matching_html | safe }}

            <h2>Mismatching Rows</h2>
            <p>These rows have AWB numbers that are either missing or different between the files.</p>
            {{ mismatching_html | safe }}

            <h2>Download Links</h2>
            <ul>
                <li><a href="/download/matching_awb_numbers.xlsx" download>Download Matching AWB Numbers</a></li>
                <li><a href="/download/mismatching_awb_numbers.xlsx" download>Download Mismatching AWB Numbers</a></li>
            </ul>
        </body>
        </html>
        ''', matching_html=matching_html, mismatching_html=mismatching_html)
    else:
        return "There was an error processing the files.", 500


@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(PROCESSED_FOLDER, filename)
    return send_file(file_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
