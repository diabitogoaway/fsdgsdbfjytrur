import os
import tempfile
import requests
import pandas as pd
from bs4 import BeautifulSoup
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

# Function to scrape and save data as Excel
def scrape_and_save(url):
    try:
        # Send an HTTP request to the URL
        response = requests.get(url)
    except requests.RequestException as e:
        print(f"RequestException: {e}")
        return None, None

    # Check if the HTTP response status code is 200 (successful)
    if response.status_code == 200:
        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')
        # Find and extract the collection title
        title = soup.find('h1', class_='collection__title').text.strip()
        # Find all product elements
        products = soup.find_all('p', class_='grid__title grid__title--product mb0')
        # Extract product names from product elements
        product_names = [product.find('span').text.strip() for product in products]
        # Create a dictionary containing the data to be saved as Excel
        data = {'Title': [title] + [''] * (len(product_names) - 1), 'Product Number': list(range(1, len(product_names) + 1)), 'Product Name': product_names}
        # Convert the dictionary to a pandas DataFrame
        df = pd.DataFrame(data)

        # Create a temporary file to save the DataFrame as Excel
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            # Save the DataFrame as Excel
            df.to_excel(temp_file, index=False)
            # Flush the temp file to ensure all data is written
            temp_file.flush()
            # Return the collection title and the temporary file path
            return title, temp_file.name
    else:
        print(f"Failed to fetch the webpage: {response.status_code}")
        # If the HTTP response status code is not 200, return None for both title and file path
        return None, None

# Define the route for the main page
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        url = request.form["url"]
        # Call the scrape_and_save function and get the collection title and temporary file path
        title, temp_file_path = scrape_and_save(url)
        # Check if the title and temp_file_path are not None (successful scraping)
        if title is not None and temp_file_path is not None:
            # Send the Excel file as a download
            return send_file(temp_file_path, as_attachment=True, download_name=f'{title}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            # Display an error message if the scraping failed
            return "Failed to fetch the webpage."

    # Render the main page template if it's a GET request
    return render_template("index.html")

# Run the Flask app
if __name__ == "__main__":
    app.run()
