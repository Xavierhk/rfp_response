import os
from flask import Flask, request, render_template, jsonify
import pandas as pd
import openai
from dotenv import load_dotenv
import csv

# Load environment variables
load_dotenv()
openai.api_key = os.getenv('OPENAI_API_KEY')

app = Flask(__name__)

def get_chatgpt_response(question):
    try:
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are my assistant to answer an RFP. You will use Juniper Network's SRX4700 as the solution for the RFP and answer all the questions."},
                {"role": "user", "content": question}
            ]
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {str(e)}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and file.filename.endswith('.csv'):
        try:
            # Read CSV file
            df = pd.read_csv(file)
            
            if 'questions' not in df.columns:
                return jsonify({'error': 'CSV file must contain a "questions" column'}), 400
            
            questions = df['questions'].tolist()
            responses = []
            
            # Get responses from ChatGPT for each question
            for question in questions:
                response = get_chatgpt_response(question)
                responses.append(response)
            
            # Create table data
            table_data = []
            for q, r in zip(questions, responses):
                table_data.append({'question': q, 'response': r})
            
            return render_template('results.html', table_data=table_data)
            
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    
    return jsonify({'error': 'Invalid file format. Please upload a CSV file'}), 400

if __name__ == '__main__':
    app.run(debug=True)