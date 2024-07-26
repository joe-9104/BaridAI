Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
      // Office is ready
  }
});

async function getAccessToken() {
  return new Promise((resolve, reject) => {
      Office.auth.getAccessTokenAsync({ allowSignInPrompt: true }, result => {
          if (result.status === "succeeded") {
              resolve(result.value);
          } else {
              reject(result.error);
          }
      });
  });
}

async function generateEmail() {
  try {
      const token = await getAccessToken();
      const emailInput = document.getElementById("emailInput").value;

      // Fetch the history of emails based on the subject (mockup for example)
      const emailHistory = []; // Fetch from your backend or use dummy data

      const response = await fetch('https://your-backend-url/generate-email', {
          method: 'POST',
          headers: {
              'Authorization': `Bearer ${token}`,
              'Content-Type': 'application/json'
          },
          body: JSON.stringify({ user_input: emailInput, email_history: emailHistory })
      });
      const data = await response.json();
      document.getElementById("result").innerText = data.email_content;
  } catch (error) {
      console.error("Error generating email:", error);
  }
}

async function translateEmail() {
  try {
    Office.context.mailbox.item.body.getAsync("text", async result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const emailBody = result.value;
        console.log('email body received');
        console.log(emailBody);

        try {
          const response = await fetch('http://localhost:5000/translate-text', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({ text: emailBody, target_language: 'fr' })
          });

          if (!response.ok) {
            throw new Error('Network response was not ok.');
          }

          const data = await response.json();  // Ensure response is valid JSON
          console.log('response created');
          document.getElementById("result").innerText = data.translated_text;
        } catch (fetchError) {
          console.error('Fetch error:', fetchError);
        }

      } else {
        console.error("Error getting email body:", result.error);
      }
    });
    
  } catch (error) {
    console.error("Error translating email:", error);
  }
}

