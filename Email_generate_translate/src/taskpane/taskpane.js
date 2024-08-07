// Define a placeholder message
let placeholderMessage = "<span style ='color: gray;'>Working on it...</span>";
let placeholderMessageSubject = "Working on it..."

// Define an error message
let errorMessage = "<span style= 'color: red;'>Sorry :(, an error occured. Check the console for more info.</span>"

// Define long and short strings
let longDescriptionTextFr = "BaridAI exploite une IA avancée pour vous aider à rédiger et générer des e-mails sans effort. Fournissez simplement les points clés et BaridAI rédigera pour vous un brouillon d'e-mail soigné, ce qui vous fera gagner du temps et améliorera la productivité."
let longDescriptionTextEn = "BaridAI leverages advanced AI to assist you in composing and generating emails effortlessly. Simply provide the key points, and BaridAI will craft a polished email draft for you, saving time and enhancing productivity."
let shortDescriptionTextFr = "BaridAI utilise une IA avancée pour vous aider à rédiger des e-mails sans effort. Gagnez du temps et augmentez la productivité."
let shortDescriptionTextEn = "BaridAI uses advanced AI to help you compose emails effortlessly. Save time and boost productivity."
let longATextFr = "Comment utiliser BaridAI?"
let longATextEn = "How to use BaridAI?"
let shortATextFr = "Voir Moins"
let shortAtextEn = "Show Less"

//Function to show more & less info
function showMoreLess(){
  const descriptionText = document.getElementById('descriptionText');
  const aText = document.getElementById('atext');
  if (aText.innerHTML == longATextEn){
    aText.innerHTML = shortAtextEn;
    descriptionText.textContent = longDescriptionTextEn;
  }else if (aText.innerHTML == shortAtextEn){
    aText.innerHTML = longATextEn;
    descriptionText.textContent = shortDescriptionTextEn;
  }else if (aText.innerHTML == shortATextFr){
    aText.innerHTML = longATextFr;
    descriptionText.textContent = shortDescriptionTextFr;
  }else if (aText.innerHTML == longATextFr){
    aText.innerHTML = shortATextFr;
    descriptionText.textContent = longDescriptionTextFr;
  }
}

//Function to set the language of the add-in
function setLanguage(lang){
  const fr_flag = document.getElementById('fr_flag');
  const uk_flag = document.getElementById('uk_flag');
  const headerText = document.getElementById('headerText');
  const descriptionText = document.getElementById('descriptionText');
  const aText = document.getElementById('atext');
  const emailInput = document.getElementById('emailInput');
  const buttonClick = document.getElementById('buttonClick');

  if (lang == 'fr'){
    fr_flag.title = "Français";
    uk_flag.title = "Basculer vers l'anglais";
    headerText.textContent = "Extension d'email";
    descriptionText.textContent = shortDescriptionTextFr;
    aText.innerHTML = longATextFr;
    emailInput.placeholder = "Entrez les principaux points de l'email";
    buttonClick.innerText = "Générer l'email"
    errorMessage = "<span style= 'color: red;'>Désolé :(, une erreur s'est produite. Consultez la console pour plus d'informations.";
    placeholderMessage = "<span style = 'color: gray;'>Travail en cours...</span>";
    placeholderMessageSubject = "Travail en cours...";
  } else if (lang === 'en') {
    fr_flag.title = "Switch to french";
    uk_flag.title = "English";
    headerText.textContent = "Email Add-in";
    descriptionText.textContent = shortDescriptionTextEn;
    aText.innerHTML = longATextEn;
    emailInput.placeholder = "Enter the main points of the email";
    buttonClick.innerText = "Generate email"
    errorMessage = "<span style= 'color: red;'>Sorry :(, an error occured. Check the console for more info.</span>";
    placeholderMessage = "<span style ='color: gray;'>Working on it...</span>";
    placeholderMessageSubject = "Working on it..."
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

    // Get the current subject and check if it starts with 'RE: '
    const currentSubject = await new Promise((resolve, reject) => {
      Office.context.mailbox.item.subject.getAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          console.error("Failed to get email subject:", result.error);
          reject(result.error);
          Office.context.mailbox.item.body.setAsync(errorMessage, {coercionType: Office.CoercionType.Html})
        }
      });
    });

    const isReply = currentSubject.startsWith('RE: ');
    console.log(isReply)

    // Get the current body
    let currentBody = '';
    currentBody = await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync("html", function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
          console.log(result.value)
        } else {
          console.error("Failed to get email body:", result.error);
          reject(result.error);
          Office.context.mailbox.item.body.setAsync(errorMessage, {coercionType: Office.CoercionType.Html})
        }
      });
    });

    // Update the email body with a placeholder message while processing
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.setAsync(placeholderMessage, { coercionType: Office.CoercionType.Html }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          console.error("Failed to set placeholder message:", result.error);
          reject(result.error);
          Office.context.mailbox.item.body.setAsync(errorMessage, {coercionType: Office.CoercionType.Html})
        }
      });
    });

    // If it's not a reply, update the subject with a placeholder message while processing
    if(!isReply){
      Office.context.mailbox.item.subject.setAsync(placeholderMessageSubject)
    }

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
    let generatedContent = data.email_content;
    const generatedSubject = data.subject_content;
    console.log('Generated email content:', generatedSubject, generatedContent);

    // Prepend the generated content to the current body
    generatedContent = generatedContent + '<br><br>' + currentBody;

    // Update the email body with the generated content
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.setAsync(generatedContent, { coercionType: Office.CoercionType.Html }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Email body updated successfully.");
          resolve();
        } else {
          console.error("Failed to update email body:", result.error);
          reject(result.error);
          Office.context.mailbox.item.body.setAsync(errorMessage, {coercionType: Office.CoercionType.Html})
        }
      });
    });

    if(!isReply){
      // Update the email subject with the generated content
      await new Promise((resolve, reject) => {
        Office.context.mailbox.item.subject.setAsync(generatedSubject, function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Email subject updated successfully.");
            resolve();
          } else {
            console.error("Failed to update email subject:", result.error);
            reject(result.error);
            Office.context.mailbox.item.body.setAsync(errorMessage, {coercionType: Office.CoercionType.Html})
          }
        });
      });
    }

    // Clear the textarea after setting the email body
    document.getElementById('emailInput').value = '';

  } catch (error) {
    console.error("Error generating email:", error);
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.setAsync(errorMessage, { coercionType: Office.CoercionType.Html }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Error message set successfully.");
          resolve();
        } else {
          console.error("Failed to update email body with error message:", result.error);
          reject(result.error);
        }
      });
    });
  }
}





