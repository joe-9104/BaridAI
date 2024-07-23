from dotenv import load_dotenv
import os
import google.generativeai as genai
import re
from deep_translator import GoogleTranslator
#Function to generate the mail content
def to_markdown(text):
    text = text.replace('•', '  *')
    text = re.sub(r'\*\*(.*?)\*\*', r'**\1**', text)
    text = re.sub(r'__(.*?)__', r'**\1**', text)
    text = re.sub(r'\*(.*?)\*', r'*\1*', text)
    text = re.sub(r'_(.*?)_', r'*\1*', text)
    return text

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
        # Start a chat session
    chat_session = model.start_chat(
        history=[]
    )
    complementary_input = 'generate only the body of the mail'
    final_input = user_input+' \n'+complementary_input
    return to_markdown(chat_session.send_message(f"{final_input}").text)
#Function to fetch the history of response mails

#Function to translate text
def translate_text(text, target_language='fr'):
    try:
        translator = GoogleTranslator(source='auto', target=target_language)
        return translator.translate(text)
    except Exception as e:
        print(f"Error in translation: {e}")
        return text