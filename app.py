from flask import Flask, render_template, request
from comparison_logic import compare_files  # Import comparison function

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/compare', methods=['POST'])
def compare():
    # Get uploaded files
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    # Call comparison function (to be implemented by other developers)
    result = compare_files(
        file1.stream.read().decode('utf-8'),
        file2.stream.read().decode('utf-8')
    )
    
    return render_template('index.html', result=result)

if __name__ == '__main__':
    app.run(debug=True)