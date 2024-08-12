from dotenv import load_dotenv
import os
import google.generativeai as genai
from flask import Flask, request, jsonify
from flask_cors import CORS
from flask_sslify import SSLify

app = Flask(__name__)
CORS(app)
SSLify(app)

def summarize_email(email_content):
    try:
        load_dotenv()
        genai.configure(api_key=os.getenv("api_key"))
        generation_config = {
            "temperature": 1,
            "top_p": 0.95,
            "top_k": 64,
            "max_output_tokens": 8192,
            "response_mime_type": "text/plain",
        }
        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            system_instruction="You are a mail summarizer.",
            generation_config=generation_config,
        )
        chat_session = model.start_chat()
        return chat_session.send_message(f"summarize this mail: \n{email_content}").text
    except Exception as e:
        return f"BaridAI faced problems: {e}"

@app.route('/summarize', methods=['POST'])
def summarize():
    email_content = request.json.get('email_content', '')
    if not email_content:
        return jsonify({'error': 'Email content is required'}), 400
    
    summary = summarize_email(email_content)
    return jsonify({'summary': summary}), 200

if __name__ == '__main__':
    app.run(debug=True)