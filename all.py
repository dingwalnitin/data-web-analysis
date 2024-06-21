import os
from flask import Flask, request, render_template_string, jsonify
import pandas as pd
import uuid
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend
import matplotlib.pyplot as plt
from io import BytesIO, StringIO
import base64
from werkzeug.serving import run_simple
from werkzeug.middleware.shared_data import SharedDataMiddleware

app = Flask(__name__)

# Ensure the upload folder exists
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# HTML Templates as strings
UPLOAD_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Upload Excel File</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f4f4f4;
        }
        h1 {
            color: #2c3e50;
            text-align: center;
            margin-bottom: 30px;
        }
        form {
            background-color: #fff;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        input[type="file"] {
            width: 100%;
            padding: 10px;
            margin-bottom: 20px;
        }
        input[type="submit"] {
            display: block;
            width: 100%;
            padding: 10px;
            background-color: #3498db;
            color: #fff;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        #progress-bar {
            width: 100%;
            background-color: #f0f0f0;
            padding: 3px;
            border-radius: 3px;
            box-shadow: inset 0 1px 3px rgba(0, 0, 0, .2);
            display: none;
        }
        #progress-bar-fill {
            height: 20px;
            background-color: #3498db;
            border-radius: 3px;
            transition: width 500ms ease-in-out;
        }
        #processing {
            text-align: center;
            display: none;
        }
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #3498db;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
    </style>
</head>
<body>
    <h1>Upload Excel File</h1>
    <form id="upload-form" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xls,.xlsx" required>
        <input type="submit" value="Upload">
    </form>
    <div id="progress-bar">
        <div id="progress-bar-fill"></div>
    </div>
    <div id="processing">
        <p>Processing file...</p>
        <div class="spinner"></div>
    </div>

    <script>
        $(document).ready(function() {
            $('#upload-form').on('submit', function(e) {
                e.preventDefault();
                var formData = new FormData(this);

                $.ajax({
                    xhr: function() {
                        var xhr = new window.XMLHttpRequest();
                        xhr.upload.addEventListener("progress", function(evt) {
                            if (evt.lengthComputable) {
                                var percentComplete = evt.loaded / evt.total;
                                percentComplete = parseInt(percentComplete * 100);
                                $('#progress-bar-fill').css('width', percentComplete + '%');
                                if (percentComplete === 100) {
                                    $('#progress-bar').hide();
                                    $('#processing').show();
                                }
                            }
                        }, false);
                        return xhr;
                    },
                    url: '/',
                    type: 'POST',
                    data: formData,
                    processData: false,
                    contentType: false,
                    beforeSend: function() {
                        $('#progress-bar').show();
                    },
                    success: function(response) {
                        $('#processing').hide();
                        $('body').html(response);
                    },
                    error: function(error) {
                        $('#processing').hide();
                        alert('An error occurred: ' + error.responseJSON.error);
                    }
                });
            });
        });
    </script>
</body>
</html>
'''

RESULTS_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Analysis Results</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <style>
        body {
            font-family: 'Arial', sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        h1, h2 {
            color: #2c3e50;
            text-align: center;
        }
        .plot-container {
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            gap: 20px;
            margin-bottom: 30px;
        }
        .plot-container img {
            max-width: 100%;
            height: auto;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        .data-form {
            background-color: #fff;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            max-width: 400px;
            margin: 0 auto;
        }
        input[type="number"] {
            width: 100%;
            padding: 10px;
            margin-bottom: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        button {
            display: block;
            width: 100%;
            padding: 10px;
            background-color: #3498db;
            color: #fff;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            transition: background-color 0.3s;
        }
        button:hover {
            background-color: #2980b9;
        }
        #result {
            margin-top: 20px;
            background-color: #fff;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        #result p {
            margin: 5px 0;
        }
    </style>
</head>
<body>
    <h1>Analysis Results</h1>
    
    <div class="plot-container">
        {% for plot in plots %}
            <img src="data:image/png;base64,{{ plot }}" alt="Analysis Plot">
        {% endfor %}
    </div>

    <h2>Get Data for Specific Second</h2>
    <div class="data-form">
        <input type="number" id="seconds" placeholder="Enter seconds">
        <button onclick="getData()">Get Data</button>
        <div id="result"></div>
    </div>

    <script>
    function getData() {
        var seconds = document.getElementById('seconds').value;
        $.ajax({
            url: '/get_data',
            method: 'POST',
            data: {
                seconds: seconds,
                df: '{{ df | safe }}'
            },
            success: function(response) {
                if (response.error) {
                    $('#result').html('<p style="color: red;">' + response.error + '</p>');
                } else {
                    $('#result').html(
                        '<p><strong>COUN Count:</strong> ' + response.COUN_Count + '</p>' +
                        '<p><strong>COUN Cumulative:</strong> ' + response.COUN_Cumulative + '</p>' +
                        '<p><strong>ENER Count:</strong> ' + response.ENER_Count + '</p>' +
                        '<p><strong>ENER Cumulative:</strong> ' + response.ENER_Cumulative + '</p>'
                    );
                }
            }
        });
    }
    </script>
</body>
</html>
'''

def generate_unique_filename(filename):
    """Generate a unique filename using UUID."""
    ext = os.path.splitext(filename)[1]
    return f"{uuid.uuid4().hex}{ext}"

def process_file(file_path):
    # Read the specific columns from the Excel file
    df = pd.read_excel(file_path, usecols=["HH:MM:SS.mmmuuun", "COUN", "ENER"])
    
    df = df.rename(columns={'COUN': 'COUN_Count', 'ENER': 'ENER_Count'})
    df["Seconds"] = df["HH:MM:SS.mmmuuun"].apply(lambda x: x.hour * 3600 + x.minute * 60 + x.second)

    # Group by the total seconds and sum the numeric values
    grouped_df = df.groupby("Seconds")[["COUN_Count", "ENER_Count"]].sum()

    # Add cumulative sum columns
    grouped_df['COUN_Cumulative'] = grouped_df['COUN_Count'].cumsum()
    grouped_df['ENER_Cumulative'] = grouped_df['ENER_Count'].cumsum()

    # Create a complete range of seconds
    all_seconds = pd.Series(range(grouped_df.index.min(), grouped_df.index.max() + 1))

    # Reindex the DataFrame to include all seconds
    filled_df = grouped_df.reindex(all_seconds)

    # Fill the new rows with 0 for COUN and ENER
    filled_df[['COUN_Count', 'ENER_Count']] = filled_df[['COUN_Count', 'ENER_Count']].fillna(0)

    # Forward fill the cumulative columns
    filled_df['COUN_Cumulative'] = filled_df['COUN_Cumulative'].ffill()
    filled_df['ENER_Cumulative'] = filled_df['ENER_Cumulative'].ffill()

    # Reset the index to make 'Seconds' a column
    filled_df = filled_df.reset_index().rename(columns={'index': 'Seconds'})

    return filled_df

def create_plot(df, x, y, title):
    plt.figure(figsize=(12, 6))
    plt.plot(df[x], df[y])
    plt.xlabel(x)
    plt.ylabel(y)
    plt.title(title)
    plt.grid(True)
    
    img = BytesIO()
    plt.savefig(img, format='png')
    img.seek(0)
    plt.close()
    return base64.b64encode(img.getvalue()).decode()

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400
        if file:
            filename = generate_unique_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            try:
                df = process_file(file_path)
            except Exception as e:
                return jsonify({'error': f'Error processing file: {str(e)}'}), 500

            plots = [
                create_plot(df, 'Seconds', 'ENER_Count', 'ENER vs Seconds'),
                create_plot(df, 'Seconds', 'COUN_Count', 'COUN vs Seconds'),
                create_plot(df, 'Seconds', 'ENER_Cumulative', 'Cumulative ENER vs Seconds'),
                create_plot(df, 'Seconds', 'COUN_Cumulative', 'Cumulative COUN vs Seconds')
            ]

            return render_template_string(RESULTS_TEMPLATE, plots=plots, df=df.to_json())
    return render_template_string(UPLOAD_TEMPLATE)

@app.route('/get_data', methods=['POST'])
def get_data():
    seconds = int(request.form['seconds'])
    df = pd.read_json(StringIO(request.form['df']))
    row = df[df['Seconds'] == seconds]
    if not row.empty:
        return jsonify({
            'COUN_Count': int(row['COUN_Count'].values[0]),
            'COUN_Cumulative': float(row['COUN_Cumulative'].values[0]),
            'ENER_Count': int(row['ENER_Count'].values[0]),
            'ENER_Cumulative': float(row['ENER_Cumulative'].values[0])
        })
    else:
        return jsonify({'error': 'No data for the given second'})

if __name__ == '__main__':
    # Use Werkzeug's run_simple for better performance
    app.wsgi_app = SharedDataMiddleware(app.wsgi_app, {
        '/uploads': app.config['UPLOAD_FOLDER']
    })
    run_simple('127.0.0.1', 5231, app, use_reloader=True, threaded=True)
