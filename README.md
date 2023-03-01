# Flask API for Editing Excel File and Sending it back to the Front

This is a Flask API that allows you to edit an Excel file and send it back to the front. This API was built using Flask and openpyxl libraries for Python.

## Prerequisites

Make sure you have the following installed on your system:

    Python 3.9+
    Flask (pip install flask)
    openpyxl (pip install openpyxl)

## How to Run

1. Clone the repository to your local machine.

2. Open the command prompt or terminal and navigate to the project directory.

3. Run the following command to start the Flask server:

```
python app.py
```

4. Once the server is running, you can access the API endpoints by visiting http://localhost:5000 in your web browser.

## API Endpoints

There are two API endpoints available:

### GET `/get_excel`

This endpoint returns a welcome message and a list of available API endpoints.

### POST `/gen_excel`

This endpoint allows you to edit an Excel file and send it back to the front.
Request

The request must contain the following:

- file: the Excel file to edit.
- sheet_name: the name of the sheet to edit.
- cell: the cell to edit.
- value: the new value for the cell.

Response

The response will contain the updated Excel file as a byte string. You can use this byte string to download the updated file on the front-end.
