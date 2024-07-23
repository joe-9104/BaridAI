Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      // Initialization code
    }
  });
  
  async function generateEmail() {
    const userInput = document.getElementById('userInput').value;
    const response = await fetch('http://localhost:5000/generate-email', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ user_input: userInput })
    });
    const data = await response.json();
    document.getElementById('output').innerText = data.email_content;
    // Code to insert the generated content into a new email
  }
  
  async function translateEmail() {
    const emailText = document.getElementById('emailText').value;
    const targetLanguage = document.getElementById('targetLanguage').value;
    const response = await fetch('http://localhost:5000/translate-text', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ text: emailText, target_language: targetLanguage })
    });
    const data = await response.json();
    document.getElementById('output').innerText = data.translated_text;
    // Code to replace the email content with the translated content
  }
  