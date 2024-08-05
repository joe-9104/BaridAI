Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
      // Office is ready
  }
});

async function generateEmail() {
  try {
    // Get the user input from the textarea element
    const userInput = document.getElementById('emailInput').value;
    console.log('User input:', userInput);

    // Make a POST request to the backend with the user input
    const response = await fetch('http://localhost:5000/generate-email', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ user_input: userInput })
    });

    if (!response.ok) {
      throw new Error('Failed to generate email content');
    }

    const data = await response.json();
    const generatedContent = data.email_content;
    const generatedSubject = data.subject_content;
    console.log('Generated email content:', generatedContent, generatedSubject);

    // Update the email subject with the generated content
    Office.context.mailbox.item.subject.setAsync(generatedSubject, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Email subject updated successfully");
      } else {
        console.error("Failed to update email subject:", result.error);
      }
    });
    // Update the email body with the generated content
    Office.context.mailbox.item.body.setAsync(generatedContent, { coercionType: Office.CoercionType.Html }, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Email body updated successfully");
      } else {
        console.error("Failed to update email body:", result.error);
      }
    });
  } catch (error) {
    console.error("Error generating email:", error);
  }
}


