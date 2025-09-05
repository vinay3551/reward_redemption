from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
from flask import send_file

app = Flask(__name__)

STUDENTS_FILE = "students.xlsx"
REDEMPTIONS_FILE = "redemptions.xlsx"

PRIZES = {
    "Field Visit": 100,
    "Novel": 70,
    "Food from teacher": 40,
    "Free PE": 50,
    "Bookmark": 10
}

def ensure_files():
    if not os.path.exists(STUDENTS_FILE):
        df = pd.DataFrame({
            "Name": ["Amit Kumar", "Riya Sharma", "Sam Patel", "Amit Sharma"],
            "Coins": [120, 85, 35, 90]
        })
        df.to_excel(STUDENTS_FILE, index=False)

    if not os.path.exists(REDEMPTIONS_FILE):
        df = pd.DataFrame(columns=["Name", "Prizes", "Remaining Coins", "Timestamp"])
        df.to_excel(REDEMPTIONS_FILE, index=False)

def load_students():
    df = pd.read_excel(STUDENTS_FILE)
    df["Name"] = df["Name"].astype(str).str.strip()
    return df

def save_students(df):
    df.to_excel(STUDENTS_FILE, index=False)

def log_redemption(name, prizes, remaining):
    if os.path.exists(REDEMPTIONS_FILE):
        df = pd.read_excel(REDEMPTIONS_FILE)
    else:
        df = pd.DataFrame(columns=["Name", "Prizes", "Remaining Coins", "Timestamp"])

    new_row = {
        "Name": name,
        "Prizes": ", ".join(prizes),
        "Remaining Coins": remaining,
        "Timestamp": pd.Timestamp.now()
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(REDEMPTIONS_FILE, index=False)

@app.route("/")
def index():
    ensure_files()
    return render_template("index.html", prizes=PRIZES)

@app.route("/search", methods=["POST"])
def search():
    data = request.get_json() or {}
    query = (data.get("name") or "").strip().lower()
    if not query:
        return jsonify({"error": "Provide a student name"}), 400

    df = load_students()
    matches = df[df["Name"].str.lower().str.contains(query)]

    if matches.empty:
        return jsonify({"error": "No students found"}), 404

    if len(matches) == 1:
        student = matches.iloc[0]
        return jsonify({"name": student["Name"], "coins": int(student["Coins"])})
    else:
        return jsonify({"matches": matches["Name"].tolist()})

@app.route("/redeem", methods=["POST"])
def redeem():
    data = request.get_json() or {}
    name = (data.get("name") or "").strip()
    prizes = data.get("prizes", {})  # dict: {"Novel":2, "Bookmark":1}

    df = load_students()
    mask = df["Name"].str.lower() == name.lower()
    if not mask.any():
        return jsonify({"error": "Student not found"}), 404

    idx = df[mask].index[0]
    balance = int(df.loc[idx, "Coins"])

    total_cost = sum(PRIZES.get(p,0)*qty for p,qty in prizes.items())

    if total_cost > balance:
        return jsonify({"error": f"Not enough coins. Total cost: {total_cost}, Available: {balance}"}), 400

    remaining = balance - total_cost
    df.loc[idx, "Coins"] = remaining
    save_students(df)
    log_redemption(df.loc[idx, "Name"], [f"{p} x{qty}" for p, qty in prizes.items()], remaining)

    return jsonify({"success": True, "remaining": remaining})

@app.route("/download/redemptions")
def download_redemptions():
    """
    Send the redemptions.xlsx file to the browser as a download.
    """
    try:
        return send_file("redemptions.xlsx", as_attachment=True)
    except Exception as e:
        return f"Error downloading file: {e}"

if __name__ == "__main__":
    app.run(debug=True)
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
    
