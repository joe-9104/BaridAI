// Define the placeholder message
let placeholderMessage = document.getElementById('placeholdermsg');

// Define an error message
let errorMessage = "<span style= 'color: red;'>Sorry :(, an error occured. Check the console for more info.</span>"

// Define long and short strings
let longDescriptionTextFr = "BaridAI exploite une IA avancée pour vous aider à rédiger et générer des e-mails sans effort. Fournissez simplement les points clés dans la langue de votre choix, et BaridAI rédigera pour vous un brouillon d'e-mail soigné, ce qui vous fera gagner du temps et améliorera la productivité."
let longDescriptionTextEn = "BaridAI leverages advanced AI to assist you in composing and generating emails effortlessly. Simply provide the key points in any language, and BaridAI will craft a polished email draft for you, saving time and enhancing productivity."
let shortDescriptionTextFr = "BaridAI utilise une IA avancée pour vous aider à rédiger des e-mails sans effort. Gagnez du temps et augmentez la productivité!"
let shortDescriptionTextEn = "BaridAI uses advanced AI to help you compose emails effortlessly. Save time and boost productivity!"
let longATextFr = "Comment utiliser BaridAI?"
let longATextEn = "How to use BaridAI?"
let shortATextFr = "Voir moins"
let shortAtextEn = "Show less"

//Function to show more & less info
function showMoreLess(){
  const descriptionText = document.getElementById('descriptionText');
  const aText = document.getElementById('atext');
  if (aText.innerHTML == longATextEn){
    aText.innerHTML = shortAtextEn;
    descriptionText.textContent = longDescriptionTextEn;
    descriptionText.classList.remove('collapsed');
    descriptionText.classList.add('expanded');
  }else if (aText.innerHTML == shortAtextEn){
    aText.innerHTML = longATextEn;
    descriptionText.textContent = shortDescriptionTextEn;
    descriptionText.classList.remove('expanded');
    descriptionText.classList.add('collapsed');
  }else if (aText.innerHTML == shortATextFr){
    aText.innerHTML = longATextFr;
    descriptionText.textContent = shortDescriptionTextFr;
    descriptionText.classList.remove('expanded');
    descriptionText.classList.add('collapsed');
  }else if (aText.innerHTML == longATextFr){
    aText.innerHTML = shortATextFr;
    descriptionText.textContent = longDescriptionTextFr;
    descriptionText.classList.remove('collapsed');
    descriptionText.classList.add('expanded');
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
    headerText.textContent = "BaridAI - Votre Assistant de Messagerie";
    descriptionText.textContent = shortDescriptionTextFr;
    aText.innerHTML = longATextFr;
    emailInput.placeholder = "Entrez les principaux points de l'email";
    buttonClick.innerText = "Générer l'email"
    errorMessage = "<span style= 'color: red;'>Désolé :(, une erreur s'est produite. Consultez la console pour plus d'informations.";
    placeholderMessage = "Travail en cours...";
  } else if (lang === 'en') {
    fr_flag.title = "Switch to french";
    uk_flag.title = "English";
    headerText.textContent = "BaridAI - Your Email Assistant";
    descriptionText.textContent = shortDescriptionTextEn;
    aText.innerHTML = longATextEn;
    emailInput.placeholder = "Enter the main points of the email";
    buttonClick.innerText = "Generate email"
    errorMessage = "<span style= 'color: red;'>Sorry :(, an error occured. Check the console for more info.</span>";
    placeholderMessage = "Working on it...";
  }
}

// Function to enable or disable the button based on textarea content
function updateButtonState() {
  const textarea = document.getElementById('emailInput');
  const button = document.getElementById('buttonClick');
  button.disabled = textarea.value.trim() === '';
}

// Event listener to <textarea> to update button state
document.getElementById('emailInput').addEventListener('input', updateButtonState);

// Initialize button state based on the current textarea content
updateButtonState();

//Main function: generating the email content
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
      // Office is ready
      console.log("BaridAI is ready...")
  }
});

async function generateEmail() {
  // Define flags
  const frFlag = document.getElementById("fr_flag");
  const ukFlag = document.getElementById("uk_flag");
  try {

    // Show the placeholder video
    placeholderMessage = document.getElementById('placeholdermsg');
    placeholderMessage.style.display = 'flex';
    placeholderMessage.scrollIntoView({ behavior : 'smooth'});

    // Disable the 'generate email' button
    const buttonClick = document.getElementById('buttonClick');
    buttonClick.disabled = true;

    // Disable flags
    frFlag.classList.add('disabledFlags');
    ukFlag.classList.add('disabledFlags');

    // Get the current subject and check if it starts with 'RE: '
    const currentSubject = await new Promise((resolve, reject) => {
      Office.context.mailbox.item.subject.getAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          console.error("Failed to get email subject:", result.error);
          reject(result.error);
          Office.context.mailbox.item.body.prependAsync(errorMessage, {coercionType: Office.CoercionType.Html})
        }
      });
    });

    const isReply = currentSubject.startsWith('RE: ');
    let currentBodyText;
    // Inform the user in the console wether it's a reply mail or not
    if(isReply){
      console.log("Behaving as a reply mail...")
      currentBodyText = '';
      currentBodyText = await new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync("text", function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            console.error("Failed to get email body:", result.error);
            reject(result.error);
            Office.context.mailbox.item.body.prependAsync(errorMessage, {coercionType: Office.CoercionType.Html})
          }
        });
      });
    }

    // Get the user input from the textarea element
    const userInput = document.getElementById('emailInput').value;

    let response;

    if(isReply){
      const totalInput = userInput + "\nThis is a response mail to this one: " + currentBodyText;
      console.log('User input: ', totalInput);
      // Make a POST request to the backend with the user input added to the original mail
      response = await fetch('http://localhost:5000/generate-email', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ user_input: totalInput })
      });

      if (!response.ok) {
        throw new Error('Failed to generate email content');
      }
    }else {
      console.log('User input:', userInput);
      // Make a POST request to the backend with the user input
      response = await fetch('http://localhost:5000/generate-email', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ user_input: userInput })
      });

      if (!response.ok) {
        throw new Error('Failed to generate email content');
      }
    }

    const data = await response.json();
    let generatedContent = data.email_content;
    const generatedSubject = data.subject_content;
    console.log('Generated email content:\n', generatedSubject, generatedContent);

    // Update the email body with the generated content
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.prependAsync(generatedContent, { coercionType: Office.CoercionType.Html }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Email body updated successfully.");
          resolve();
        } else {
          console.error("Failed to update email body:", result.error);
          reject(result.error);
          Office.context.mailbox.item.body.prependAsync(errorMessage, {coercionType: Office.CoercionType.Html})
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
            Office.context.mailbox.item.body.prependAsync(errorMessage, {coercionType: Office.CoercionType.Html})
          }
        });
      });
    }

    // Clear the textarea after setting the email body
    document.getElementById('emailInput').value = '';

  } catch (error) {
    console.error("Error generating email:", error);
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.prependAsync(errorMessage, { coercionType: Office.CoercionType.Html }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Error message set successfully.");
          resolve();
        } else {
          console.error("Failed to update email body with error message:", result.error);
          reject(result.error);
        }
      });
    });
  } finally {
    // Hide the placeholder video
    placeholderMessage.style.display = 'none';

    // Enable or disable the button (referring to the content of the <textarea>)
    updateButtonState();

    // Unable flags
    frFlag.classList.remove('disabledFlags');
    ukFlag.classList.remove('disabledFlags');
  }
}





