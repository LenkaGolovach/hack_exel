from flask import Flask, request, jsonify, send_from_directory
import subprocess

app = Flask(__name__, static_folder='../build')

# In-memory storage for demonstration purposes
spreadsheet_data = {}

@app.route('/')
def serve():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/<path:path>')
def static_proxy(path):
    return send_from_directory(app.static_folder, path)

@app.route('/save', methods=['POST'])
def save_data():
    data = request.json
    spreadsheet_data[data['id']] = data['content']
    return jsonify({'status': 'success', 'message': 'Data saved successfully'})

@app.route('/load/<sheet_id>', methods=['GET'])
def load_data(sheet_id):
    data = spreadsheet_data.get(sheet_id, None)
    if data:
        return jsonify({'status': 'success', 'content': data})
    else:
        return jsonify({'status': 'error', 'message': 'Data not found'}), 404

@app.route('/start-react', methods=['GET'])
def start_react():
    try:
        subprocess.Popen(['npm', 'start'], cwd='../')
        return jsonify({'status': 'success', 'message': 'React app started'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
