from flask import Flask, render_template, request,  jsonify, send_from_directory
from flask_cors import CORS  # Import CORS from Flask-CORS
import os
import pandas as pd
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

@app.route('/')
def index():
    html_content = generate_html()  # Call generate_html() here
    return html_content

@app.route('/output/<path:filename>', methods=['GET'])
def download_file(filename):
    directory = 'output'
    return send_from_directory(directory, filename, as_attachment=True)

def extract_data_from_excel(file_path, isins):
    try:
        # Extract date from the file name
        file_name = os.path.basename(file_path)
        date = file_name.split('TRADING REPORT FOR GFIM-')[1][:8]

        # Load the Excel file
        df = pd.read_excel(file_path, sheet_name='NEW GOG NOTES AND BONDS ', header=3)

        # Filter the DataFrame based on column numbers
        column_d = 3  # Column D (ISIN)
        column_f = 5  # Column F (CLOSING YIELD)
        column_g = 6  # Column G (END OF DAY CLOSING PRICE)
        column_h = 7  # Column H (VOLUME)
        column_i = 8  # Column I (NUMBER TRADED)

        # Filter the DataFrame
        filtered_df = df.iloc[:, [column_d, column_f, column_g,column_h,column_i]]

        # Rename columns
        filtered_df.columns = ['ISIN','CLOSING YIELD','END OF DAY CLOSING PRICE','VOLUME','NUMBER TRADED']

        # Filter rows based on ISIN values from user input
        filtered_df = filtered_df[filtered_df['ISIN'].isin(isins)]

        # Convert date column format from "ddmmyyyy" to "mm/dd/yyyy"
        date_formatted = datetime.strptime(date, '%d%m%Y').strftime('%m/%d/%Y')
        filtered_df['Date'] = date_formatted

        return filtered_df
    except ValueError:
        print(f"Worksheet named 'NEW GOG NOTES AND BONDS ' not found in file: {file_path}")
        return None


@app.route('/process_data', methods=['POST'])
def process_data():
    if request.method == 'POST':
        # Retrieve form data
        data = request.json
        start_date = data['startDate']
        end_date = data['endDate']
        isins = data['isins']
        
        # Define folder containing Excel files
        folder_path = "DATA"

        # Initialize an empty DataFrame to store all data
        combined_df = pd.DataFrame()

        # Iterate through dates from start_date to end_date
        current_date = datetime.strptime(start_date, '%Y-%m-%d')
        end_date = datetime.strptime(end_date, '%Y-%m-%d')
        while current_date <= end_date:
            # Generate file name based on the current date
            file_name = 'TRADING REPORT FOR GFIM-' + current_date.strftime('%d%m%Y') + '.xlsx'
            file_path = os.path.join(folder_path, file_name)

            # Check if the file exists
            if os.path.isfile(file_path):
                # Extract data from the file
                df = extract_data_from_excel(file_path, isins)
                if df is not None:
                    combined_df = pd.concat([combined_df, df], ignore_index=True)
            else:
                # Check if the current day is a weekend
                if current_date.weekday() == 5 or current_date.weekday() == 6:  # Saturday or Sunday
                    print(f"File not found: {file_path}. It's a Weekend.")
                else:
                    print(f"File not found: {file_path}")
            
            # Move to the next date
            current_date += timedelta(days=1)

        # Export the combined DataFrame to Excel
        output_folder = "output"
        os.makedirs(output_folder, exist_ok=True)
        output_file_path = os.path.join(output_folder, 'historical_gfim.xlsx')
        combined_df.to_excel(output_file_path, index=False)
        
        return "Data processing completed successfully!"
    
def generate_html():
    html = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GFIM Historical Price Downloader</title>
<style>
    body {
        font-family: Montserrat, sans-serif;
        background-image: url('https://onlineservice.databankgroup.com/redemptions/Images/bg5.jpg');
        background-size: cover;
        background-repeat: no-repeat;
        margin: 0;
       
    }

    .container {
        max-width: 600px;
        background-color: rgba(241, 241, 241, 0.7); 
        margin: 50px auto;
        padding: 60px;
        border: 1px solid #ccc;
        border-radius: 5px;
    }

    .form-group {
        margin-bottom: 15px;
    }

    label {
        display: block;
        font-weight: bold;
    }

    input[type="date"],
    input[type="text"] {
        width: 100%;
        padding: 10px;
        border: 1px solid #ccc;
        border-radius: 5px;
    }

    button {
        background-color: #4CAF50;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        margin-top: 10px;
    }

    #submitBtn {
        display: inline-block; /* Align buttons horizontally */
    }

    button:hover {
        background-color: #45a049;
    }

    #status {
        margin-top: 10px;
        font-weight: bold;
    }

    p {
        text-align: justify;
        animation: mymove 5s infinite;
    }

    @keyframes mymove {
        40% {font-weight: bold;}
    }

    h2{
        text-align: center;
       
    }
</style>
</head>
<body>
<div class="container">
  <h2>GFIM Historical Price Downloader</h2>
  <form id="dataForm">
    <div class="form-group">
      <label for="startDate">Start Date:</label>
      <input type="date" id="startDate" name="startDate" required>
    </div>
    <div class="form-group">
      <label for="endDate">End Date:</label>
      <input type="date" id="endDate" name="endDate" required>
    </div>
    <div class="form-group">
      <label for="isin">ISINs (Separated by commas):</label>
      <input type="text" id="isin" name="isin" placeholder="Enter ISINs" required>
    </div>
    <div>
        <button id="submitBtn" type="button">Fetch Data</button>
        <span style="margin: 0 10px;"></span> <!-- Space between buttons -->
        <button id="downloadFile">Download File</button>
    </div>
    <div id="status"></div> <!-- Status div -->
  </form>
  <span style="margin: 0 30px;"></span> <!-- Space -->
  <p>This program enables users to download historical daily prices and yields for GFIM securities. Users can input the start and end dates, along with the ISINs of the securities they are interested in. Fetch the data to create a downloadable excel file </p>
</div>
<script>
    document.getElementById("submitBtn").addEventListener("click", function(event) {
        event.preventDefault();
        var startDate = document.getElementById("startDate").value;
        var endDate = document.getElementById("endDate").value;
        var isins = document.getElementById("isin").value.split(",").map(item => item.trim());

        // Create a data object with the form data
        var data = {
            startDate: startDate,
            endDate: endDate,
            isins: isins
        };

        // Display status
        document.getElementById('status').innerText = "System fetching data...";

        // Send a POST request to the Flask server
        fetch('/process_data', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        })
        .then(response => {
            if (response.ok) {
                console.log("Data sent successfully!");
                // Update status
                document.getElementById('status').innerText = "Data successfully fetched, please download it";
                // Display the download file button
                document.getElementById('downloadFile').style.display = 'block';
            } else {
                console.error("Error:", response.statusText);
                // Update status in case of error
                document.getElementById('status').innerText = "Error fetching data";
            }
        })
        .catch(error => {
            console.error("Error:", error);
            // Update status in case of error
            document.getElementById('status').innerText = "Error fetching data";
        });
    });

  
    document.getElementById('downloadFile').addEventListener('click', function() {
          
        // Construct the URL for the file download
        var downloadUrl = `/output/historical_gfim.xlsx`;

        // Create a temporary anchor element
        var downloadLink = document.createElement("a");
        downloadLink.href = downloadUrl;
        downloadLink.download = `historical_gfim.xlsx`;
        downloadLink.style.display = "none";
        document.body.appendChild(downloadLink);

        // Trigger the download by programmatically clicking the link
        downloadLink.click();

        // Remove the temporary link after triggering the download
        document.body.removeChild(downloadLink);
    });
</script>
</body>
</html>

"""
    return html

if __name__ == '__main__':
    html_content = generate_html()  # Call generate_html() here before running the app
    app.run(host="0.0.0.0", port=1000, debug=True)

