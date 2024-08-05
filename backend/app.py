from dotenv import load_dotenv
import os
import google.generativeai as genai
import markdown2
from flask import Flask, jsonify, request
from flask_cors import CORS
from flask_sslify import SSLify

app = Flask(__name__)
CORS(app)
SSLify(app)

# Function to generate email content
def generate_mail(user_input):
    load_dotenv()
    genai.configure(api_key=os.getenv("api_key"))
    generation_config = {
        "temperature": 1,
        "top_p": 0.95,
        "top_k": 64,
        "max_output_tokens": 8192,
        "response_mime_type": "text/plain",
    }
    model1 = genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        system_instruction="You are a mail generator destined to generate only the body of a professional mail. The body must not contain [Your Name] or [Reason] or [Date]. Do not include the signature.",
        generation_config=generation_config,
    )
    model2 = genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        system_instruction="You are a mail generator destined to generate only the subject of a professional mail.",
        generation_config=generation_config,
    )
    chat_session1 = model1.start_chat()
    chat_session2 = model2.start_chat()
    return markdown2.markdown(chat_session1.send_message(f"{user_input}").text), chat_session2.send_message(f"{user_input}").text

@app.route('/generate-email', methods=['POST'])
def generate_email():
    data = request.json
    user_input = data.get('user_input')
    if not user_input:
        return jsonify({"error": "No input provided"}), 400
    email_content, subject_content = generate_mail(user_input)
    return jsonify({"email_content": email_content, "subject_content": subject_content})

if __name__ == '__main__':
    app.run(debug=True)
