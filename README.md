# Lab Inventory Website

This project reads the lab inventory workbook and renders a simple web interface to:

- browse available parts
- see where each item is stored in the lab
- submit a borrow request
- approve or reject requests from an admin page
- show approved borrowed items in a borrowed-parts list

## Run locally

1. Install dependencies:
   pip install -r requirements.txt

2. Start the app:
   python app.py

3. Open the site in a browser:
   http://127.0.0.1:5000

4. Open the admin page:
   http://127.0.0.1:5000/admin

## Deploy to Render

This project is set up for a standard Render Python web service.

1. Push the repository to GitHub.
2. In Render, create a new Web Service and connect the repo.
3. Use the default Python environment, or keep the included render.yaml file to configure the service.
4. Render will run the app with:
   gunicorn app:app --bind 0.0.0.0:$PORT

The service automatically uses the generated secret key and secure session cookie settings for HTTPS deployments.
