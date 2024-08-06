//Function to set the language of the add-in
function setLanguage(lang){
  const headerText = document.getElementById('headerText');
  const descriptionText = document.getElementById('descriptionText');
  const emailInput = document.getElementById('emailInput');

  if (lang == 'fr'){
    headerText.textContent= "Extension d'email";
    descriptionText.textContent= "Cette extension exploite une IA avancée pour vous aider à rédiger et à générer des e-mails sans effort. Fournissez simplement les points clés et notre IA créera pour vous un brouillon d'e-mail soigné, ce qui vous fera gagner du temps et améliorera votre productivité.";
    emailInput.placeholderMessage= "Entrez les principaux points de l'email";
  } else if (lang === 'en') {
    headerText.textContent = "Email Add-in";
    descriptionText.textContent = "This add-in leverages advanced AI to assist you in composing and generating emails effortlessly. Simply provide the key points, and our AI will craft a polished email draft for you, saving time and enhancing productivity.";
    emailInput.placeholderMessage = "Enter the main points of the email";
  }
}

//Main function: generating the email content
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
      // Office is ready
  }
});

async function generateEmail() {
  try {

    //Update the email body with a placeholder message
    const placeholderMessage = "<span style ='color: gray;'>Working on it...</span>";
    Office.context.mailbox.item.body.setAsync(placeholderMessage, {coercionType: Office.CoercionType.Html})

    //Define an error message
    const errorMessage = "<span style= 'color: red;'>Sorry :(, an error occured. Check the console for more info.</span>"

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
        Office.context.mailbox.item.body.setAsync(errorMessage, {coercionType: Office.CoercionType.Html})
      }
    });
    // Update the email body with the generated content
    Office.context.mailbox.item.body.setAsync(generatedContent, { coercionType: Office.CoercionType.Html }, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Email body updated successfully");
        // Clear the textarea after setting the email body
        document.getElementById('emailInput').value = '';
      } else {
        console.error("Failed to update email body:", result.error);
        Office.context.mailbox.item.body.setAsync(errorMessage, {coercionType: Office.CoercionType.Html})
      }
    });
  } catch (error) {
    console.error("Error generating email:", error);
    Office.context.mailbox.item.body.setAsync(errorMessage, {coercionType: Office.CoercionType.Html})
  }
}


