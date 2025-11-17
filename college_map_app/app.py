from flask import Flask, render_template, request, jsonify, send_file, redirect, session, url_for
import pandas as pd
import os
import shutil
import tempfile

app = Flask(__name__)
app.secret_key = "sudha_super_secret_key"

feedbacks = []

@app.route('/')
def home():
    if 'user' in session:
        return redirect('/map')
    return render_template('login.html')

@app.route('/map')
def map_page():
    if 'user' not in session:
        return redirect('/')
    return render_template('index.html')

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect('/')

@app.route('/submit_login', methods=['POST'])
def submit_login():
    name = request.form.get("name")
    mobile = request.form.get("mobile")
    email = request.form.get("email")
    purpose = request.form.get("purpose")

    if not (name and mobile and purpose):
        return "Missing required fields", 400

    session['user'] = name

    login_data = {
        "Name": name,
        "Mobile": mobile,
        "Email": email,
        "Purpose": purpose
    }

    file = "logins.xlsx"

    try:
        if os.path.exists(file):
            df = pd.read_excel(file, engine="openpyxl")
            df = pd.concat([df, pd.DataFrame([login_data])], ignore_index=True)
        else:
            df = pd.DataFrame([login_data])


        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            temp_path = tmp.name
        df.to_excel(temp_path, index=False, engine="openpyxl")


        shutil.move(temp_path, file)

    except PermissionError:
        return "The logins.xlsx file is currently in use. Please close it and try again.", 500

    return redirect("/map")

@app.route('/submit_feedback', methods=['POST'])
def submit_feedback():
    data = request.get_json()
    name = data.get("name") or "Anonymous"
    location = data.get("location")
    comment = data.get("comment")
    category = data.get("category", "General")
    rating = data.get("rating", "Not Rated")

    if not location or not comment:
        return jsonify({"error": "Missing data"}), 400

    entry = {
        "Name": name,
        "Location": location,
        "Category": category,
        "Rating": rating,
        "Comment": comment
    }

    feedbacks.append(entry)

    try:
        df = pd.DataFrame(feedbacks)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            temp_path = tmp.name
        df.to_excel(temp_path, index=False, engine="openpyxl")
        shutil.move(temp_path, "feedbacks.xlsx")
    except PermissionError:
        return jsonify({"error": "Feedback file is currently in use. Please close it."}), 500

    return jsonify({"message": "Feedback saved", "feedbacks": feedbacks})

@app.route('/get_feedbacks')
def get_feedbacks():
    return jsonify(feedbacks)

@app.route('/download_excel')
def download_excel():
    file_path = "feedbacks.xlsx"
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "No feedback file found", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

