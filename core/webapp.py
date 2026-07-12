from flask import Flask, render_template, redirect, url_for
import subprocess
import threading
import webbrowser

app = Flask(__name__)

# Route for the front page
@app.route("/")
def index():
    return render_template("index.html")  # Your front page HTML in templates/

# Route to start Streamlit resume analyzer
@app.route("/streamlit")
def launch_streamlit():
    # Run Streamlit in a separate thread so Flask doesn't block
    def run_streamlit():
        subprocess.run(["streamlit", "run", "app.py"])  # Path to your existing Streamlit app

    threading.Thread(target=run_streamlit).start()
    # Open Streamlit in a new tab (optional)
    webbrowser.open_new_tab("http://localhost:8501")
    return redirect("http://localhost:8501")  # Redirect user to Streamlit app

if __name__ == "__main__":
    app.run(debug=True)
