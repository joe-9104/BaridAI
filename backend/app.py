from dotenv import load_dotenv
import os
import google.generativeai as genai
import re
from deep_translator import GoogleTranslator
from flask import Flask, jsonify, request
from flask_cors import CORS
from flask_sslify import SSLify

app = Flask(__name__)
CORS(app)
SSLify(app)

# Function to convert text to Markdown format
def to_markdown(text):
    text = text.replace('•', '  *')
    text = re.sub(r'\*\*(.*?)\*\*', r'**\1**', text)
    text = re.sub(r'__(.*?)__', r'**\1**', text)
    text = re.sub(r'\*(.*?)\*', r'*\1*', text)
    text = re.sub(r'_(.*?)_', r'*\1*', text)
    return text

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
    model = genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        generation_config=generation_config,
    )
    chat_session = model.start_chat(
        history=[]
    )
    complementary_input = 'generate only the body of the mail'
    final_input = user_input + ' \n' + complementary_input
    return to_markdown(chat_session.send_message(f"{final_input}").text)

# Function to translate text
def translate_text(text, target_language='fr'):
    try:
        translator = GoogleTranslator(source='auto', target=target_language)
        return translator.translate(text)
    except Exception as e:
        print(f"Error in translation: {e}")
        return text

@app.route('/generate-email', methods=['POST'])
def generate_email():
    data = request.json
    user_input = data.get('user_input')
    if not user_input:
        return jsonify({"error": "No input provided"}), 400
    email_content = generate_mail(user_input)
    return jsonify({"email_content": email_content})

@app.route('/translate-text', methods=['POST'])
def translate():
    data = request.json
    text = data.get('text')
    target_language = data.get('target_language', 'fr')
    if not text:
        return jsonify({"error": "No text provided"}), 400
    translated_text = translate_text(text, target_language)
    return jsonify({"translated_text": translated_text})

if __name__ == '__main__':
    app.run(debug=True)
