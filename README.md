# Inventory System using Google Apps Script & Tailwind CSS

This is a simple inventory system that uses Google Sheets as a database and Google Apps Script for the backend. The frontend is built with HTML and Tailwind CSS.

## Files

- `index.html`: The frontend user interface.
- `Code.js`: The backend logic to be placed in Google Apps Script.

## Setup Instructions

1.  **Create a Google Sheet**
    - Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet.
    - Name it "Inventory System" (or anything you like).

2.  **Open Apps Script**
    - In your Google Sheet, go to the menu: `Extensions` > `Apps Script`.

3.  **Setup Backend Code**
    - Clear the default code in `Code.gs`.
    - Copy the content of the local `Code.js` file and paste it into `Code.gs` in the browser.
    - Save the project (Ctrl+S).

4.  **Setup Frontend Code**
    - In the Apps Script editor, click the `+` icon next to **Files** and select **HTML**.
    - Name the file `index`. (It will become `index.html`).
    - Copy the content of the local `index.html` file and paste it into this new file in the browser.

5.  **Deploy**
    - Click the blue **Deploy** button > **New deployment**.
    - Click the gear icon (Select type) > **Web app**.
    - Description: "Initial deploy".
    - **Execute as**: "Me" (your email).
    - **Who has access**: "Anyone" (easiest for testing) or "Anyone with Google account".
    - Click **Deploy**.
    - Authorize access if prompted (Click Review permissions > Choose account > Advanced > Go to (Project Name) (unsafe) > Allow).

6.  **Run**
    - You will get a **Web App URL**. Click it to open your Inventory System.
    - Try adding an item. It should appear in your Google Sheet and in the list below the form!
