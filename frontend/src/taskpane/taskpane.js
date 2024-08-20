// Define the placeholder message
let placeholderMessage = document.getElementById('placeholdermsg');

// Define an error message
let errorMessage = "<span style= 'color: red;'>Sorry :(, an error occured. Check the console for more info.</span>";

// Language-specific strings
const texts = {
  en: {
    fr_flag_title: "Switch to french",
    uk_flag_title: "English",
    ar_flag_title: "Switch to arabic",
    es_flag_title: "Switch to spanish",
    pr_flag_title: "Switch to protuguese",
    de_flag_title: "Switch to german",
    headerText: "BaridAI - Your Email Assistant",
    longDescription: "BaridAI leverages advanced AI to assist you in composing and generating emails effortlessly. Simply provide the key points in any language, and BaridAI will craft a polished email draft for you, saving time and enhancing productivity.",
    shortDescription: "BaridAI uses advanced AI to help you compose emails effortlessly. Save time and boost productivity!",
    longAText: "How to use BaridAI?",
    shortAText: "Show less",
    buttonTitle: "Write the main points of the email to start generating",
    emailPlaceholder: "Enter the main points of the email in your preferred language",
    buttonClickText: "Generate email",
    errorMessage: "<span style='color: red;'>Sorry :(, an error occurred. Check the console for more info.</span>",
    placeholderMessage: "Working on it..."
  },
  fr: {
    fr_flag_title: "Français",
    uk_flag_title: "Basculer vers l'anglais",
    ar_flag_title: "Basculer vers l'arabe",
    es_flag_title: "Basculer vers l'espagnol",
    pr_flag_title: "Basculer vers le portugais",
    de_flag_title: "Basculer vers l'allemand",
    headerText: "BaridAI - Votre Assistant de Messagerie",
    longDescription: "BaridAI exploite une IA avancée pour vous aider à rédiger et générer des e-mails sans effort. Fournissez simplement les points clés dans la langue de votre choix, et BaridAI rédigera pour vous un brouillon d'e-mail soigné, ce qui vous fera gagner du temps et améliorera la productivité.",
    shortDescription: "BaridAI utilise une IA avancée pour vous aider à rédiger des e-mails sans effort. Gagnez du temps et augmentez la productivité!",
    longAText: "Comment utiliser BaridAI?",
    shortAText: "Voir moins",
    buttonTitle: "Entrez les principaux points de l'email pour générer",
    emailPlaceholder: "Entrez les principaux points de l'email dans votre langue préférée",
    buttonClickText: "Générer l'email",
    errorMessage: "<span style='color: red;'>Désolé :(, une erreur s'est produite. Consultez la console pour plus d'informations.</span>",
    placeholderMessage: "Travail en cours..."
  },
  ar: {
    fr_flag_title: "التبديل إلى الفرنسية",
    uk_flag_title: "التبديل إلى الإنجليزية",
    ar_flag_title: "العربية",
    es_flag_title: "التبديل إلى الإسبانية",
    pr_flag_title: "التبديل إلى البرتغالية",
    de_flag_title: "التبديل إلى الألمانية",
    headerText: "BaridAI - مساعد البريد الإلكتروني الخاص بك",
    longDescription: "BaridAI يستفيد من الذكاء الاصطناعي المتقدم لمساعدتك في تأليف وتوليد رسائل البريد الإلكتروني بكل سهولة. ما عليك سوى تقديم النقاط الرئيسية بأي لغة، وسيتولى BaridAI صياغة مسودة بريد إلكتروني أنيقة، مما يوفر الوقت ويعزز الإنتاجية.",
    shortDescription: "BaridAI يستخدم الذكاء الاصطناعي المتقدم لمساعدتك في تأليف الرسائل الإلكترونية بسهولة. وفر الوقت وزيِّد الإنتاجية!",
    longAText: "كيفية استخدام BaridAI؟",
    shortAText: "عرض أقل",
    buttonTitle: "اكتب النقاط الرئيسية للبريد الإلكتروني لبدء الصياغة",
    emailPlaceholder: "اكتب النقاط الرئيسية للبريد الإلكتروني بلغتك المفضلة",
    buttonClickText: "صياغة البريد الإلكتروني",
    errorMessage: "<span style='color: red;'>عذرًا :(، حدث خطأ. تحقق من وحدة التحكم لمزيد من المعلومات.</span>",
    placeholderMessage: "جارٍ العمل...",
    rtl: true
  },
  es: {
    fr_flag_title: "Cambiar a francés",
    uk_flag_title: "Cambiar a inglés",
    ar_flag_title: "Cambiar a árabe",
    es_flag_title: "Español",
    pr_flag_title: "Cambiar a portugués",
    de_flag_title: "Cambiar a alemán",
    headerText: "BaridAI - Su asistente de correo electrónico",
    longDescription: "BaridAI utiliza inteligencia artificial avanzada para ayudarte a redactar y generar correos electrónicos sin esfuerzo. Simplemente proporciona los puntos clave en cualquier idioma, y BaridAI creará un borrador pulido del correo electrónico para ti, ahorrando tiempo y mejorando la productividad.",
    shortDescription: "BaridAI usa IA avanzada para ayudarte a redactar correos electrónicos sin esfuerzo. ¡Ahorra tiempo y aumenta la productividad!",
    longAText: "¿Cómo usar BaridAI?",
    shortAText: "Mostrar menos",
    buttonTitle: "Escribe los puntos principales del correo electrónico para comenzar a generar",
    emailPlaceholder: "Escriba los puntos principales del correo electrónico en su idioma preferido.",
    buttonClickText: "Generar correo electrónico",
    errorMessage: "<span style='color: red;'>Lo siento :(, ocurrió un error. Consulta la consola para más información.</span>",
    placeholderMessage: "Trabajando en ello..."
  },
  pt: {
    fr_flag_title: "Mudar para francês",
    uk_flag_title: "Mudar para inglês",
    ar_flag_title: "Mudar para árabe",
    es_flag_title: "Mudar para espanhol",
    pr_flag_title: "Português",
    de_flag_title: "Mudar para alemão",
    headerText: "BaridAI - Seu Assistente de Email",
    longDescription: "BaridAI usa inteligência artificial avançada para ajudá-lo a redigir e gerar e-mails sem esforço. Basta fornecer os pontos principais em qualquer idioma, e o BaridAI criará um rascunho polido do e-mail para você, economizando tempo e aumentando a produtividade.",
    shortDescription: "BaridAI usa inteligência artificial avançada para ajudá-lo a redigir e-mails com facilidade. Economize tempo e aumente a produtividade!",
    longAText: "Como usar o BaridAI?",
    shortAText: "Mostrar menos",
    buttonTitle: "Escreva os principais pontos do e-mail para começar a gerar",
    emailPlaceholder: "Introduza os pontos principais do e-mail no idioma da sua preferência",
    buttonClickText: "Gerar e-mail",
    errorMessage: "<span style='color: red;'>Desculpe :(, ocorreu um erro. Verifique o console para mais informações.</span>",
    placeholderMessage: "Trabalhando nisso..."
  },
  de: {
    fr_flag_title: "Wechseln Sie zu Französisch",
    uk_flag_title: "Zu Englisch wechseln",
    ar_flag_title: "Zu Arabisch wechseln",
    es_flag_title: "Zu Spanisch wechseln",
    pr_flag_title: "Zu Portugiesisch wechseln",
    de_flag_title: "Deutsch",
    headerText: "BaridAI - Ihr E-Mail-Assistent",
    longDescription: "BaridAI nutzt fortschrittliche KI, um Ihnen beim Verfassen und Erstellen von E-Mails mühelos zu helfen. Geben Sie einfach die Hauptpunkte in einer beliebigen Sprache ein, und BaridAI erstellt für Sie einen ausgefeilten E-Mail-Entwurf, spart Zeit und steigert die Produktivität.",
    shortDescription: "BaridAI verwendet fortschrittliche KI, um Ihnen beim Verfassen von E-Mails mühelos zu helfen. Sparen Sie Zeit und steigern Sie die Produktivität!",
    longAText: "Wie benutzt man BaridAI?",
    shortAText: "Weniger anzeigen",
    buttonTitle: "Geben Sie die Hauptpunkte der E-Mail ein, um mit der Generierung zu beginnen",
    emailPlaceholder: "Geben Sie die Hauptpunkte der E-Mail in Ihrer bevorzugten Sprache ein",
    buttonClickText: "E-Mail generieren",
    errorMessage: "<span style='color: red;'>Entschuldigung :(, ein Fehler ist aufgetreten. Überprüfen Sie die Konsole für weitere Informationen.</span>",
    placeholderMessage: "Arbeiten daran..."
  }
};

// Default language
let currentLanguage = 'en';

// Function to show more & less info
function showMoreLess(){
  const descriptionText = document.getElementById('descriptionText');
  const aText = document.getElementById('atext');

  if (aText.innerHTML === texts[currentLanguage].longAText) {
    aText.innerHTML = texts[currentLanguage].shortAText;
    descriptionText.textContent = texts[currentLanguage].longDescription;
    descriptionText.classList.remove('collapsed');
    descriptionText.classList.add('expanded');
  } else if (aText.innerHTML === texts[currentLanguage].shortAText) {
    aText.innerHTML = texts[currentLanguage].longAText;
    descriptionText.textContent = texts[currentLanguage].shortDescription;
    descriptionText.classList.remove('expanded');
    descriptionText.classList.add('collapsed');
  }
}


// Function to set the language of the add-in
function setLanguage(lang) {
  if (!texts[lang]) {
    console.error("Language not supported");
    return;
  }

  currentLanguage = lang;

  const fr_flag = document.getElementById('fr_flag');
  const uk_flag = document.getElementById('uk_flag');
  const ar_flag = document.getElementById('ar_flag');
  const es_flag = document.getElementById('es_flag');
  const pr_flag = document.getElementById('pr_flag');
  const de_flag = document.getElementById('de_flag');
  const headerText = document.getElementById('headerText');
  const descriptionText = document.getElementById('descriptionText');
  const aText = document.getElementById('atext');
  const emailInput = document.getElementById('emailInput');
  const buttonClick = document.getElementById('buttonClick');

  const langData = texts[lang];
  // Update flag titles
  fr_flag.title = langData.fr_flag_title;
  uk_flag.title = langData.uk_flag_title;
  ar_flag.title = langData.ar_flag_title;
  es_flag.title = langData.es_flag_title;
  pr_flag.title = langData.pr_flag_title;
  de_flag.title = langData.de_flag_title;
  
  // Update UI elements based on selected language
  headerText.textContent = langData.headerText;
  
  descriptionText.textContent = langData.shortDescription;
  aText.innerHTML = langData.longAText;
  emailInput.placeholder = langData.emailPlaceholder;
  buttonClick.innerText = langData.buttonClickText;
  
  errorMessage = langData.errorMessage;
  placeholderMessage = langData.placeholderMessage;
  buttonTitle = langData.buttonTitle;

  // Text align (right-to-left for arabic)
  if (langData.rtl) {
    document.body.style.direction = 'rtl';
    aText.style.paddingLeft = '0';
    aText.style.paddingRight = '20px';
  } else {
      document.body.style.direction = 'ltr';
      aText.style.paddingRight = '0';
      aText.style.paddingLeft = '20px'; 
  }

  // Update button state after language change
  updateButtonState();
}

// Function to enable or disable the button based on textarea content
function updateButtonState() {
  const textarea = document.getElementById('emailInput');
  const button = document.getElementById('buttonClick');
  button.disabled = textarea.value.trim() === '';
  if(button.disabled){
    button.title = texts[currentLanguage].buttonTitle;
  }else{
    button.title = '';
  }
}

// Event listener to <textarea> to update button state
document.getElementById('emailInput').addEventListener('input', updateButtonState);

// Initialize button state based on the current textarea content
updateButtonState();

//Main function: generating the email content
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
      // Office is ready
      console.log("BaridAI - Mail Generator is ready...")
  }
});

async function generateEmail() {
  // Define flags
  const flags = document.getElementById("flags");
  try {

    // Show the placeholder video
    placeholderMessage = document.getElementById('placeholdermsg');
    placeholderMessage.style.display = 'flex';
    placeholderMessage.scrollIntoView({ behavior : 'smooth'});

    // Disable the 'generate email' button
    const buttonClick = document.getElementById('buttonClick');
    buttonClick.disabled = true;

    // Disable flags
    flags.classList.add('disabledFlags');

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

      if(response.status === 422){ // 'Unprocessable entity' HTTP error status
        const data = await response.json();
        console.error("Backend error: ", data.email_content);

        // Update the mail body with the error message
        await new Promise((resolve, reject) => {
          Office.context.mailbox.item.body.prependAsync('<p style="color : red";>BaridAI faced problems, please try again later or check the console for more info.</p>',
            {coercionType: Office.CoercionType.Html},
            function(result){
              if(result.status === Office.AsyncResultStatus.Succeeded){
                resolve();
              }else{
                console.error("Failed to update email body with error message: ", result.error);
                reject(result.error);
              }
            }
          );
        });

        // Update the subject with the error message
        await new Promise((resolve, reject) => {
          Office.context.mailbox.item.subject.setAsync(data.subject_content,
            function(result){
              if(result.status === Office.AsyncResultStatus.Succeeded){
                resolve();
              }else{
                console.error("Failed to update email subject with error message: ", result.error);
                reject(result.error);
              }
            }
          );
        });

        return; // Prevent further error handling (if(!response.ok) statement)
      }

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

      if(response.status === 422){ // 'Unprocessable entity' HTTP error status
        const data = await response.json();
        console.error("Backend error: ", data.email_content);

        // Update the mail body with the error message
        await new Promise((resolve, reject) => {
          Office.context.mailbox.item.body.prependAsync('<p style="color : red";>BaridAI faced problems, please try again later or check the console for more info.</p>',
            {coercionType: Office.CoercionType.Html},
            function(result){
              if(result.status === Office.AsyncResultStatus.Succeeded){
                resolve();
              }else{
                console.error("Failed to update email body with error message: ", result.error);
                reject(result.error);
              }
            }
          );
        });

        // Update the subject with the error message
        await new Promise((resolve, reject) => {
          Office.context.mailbox.item.subject.setAsync(data.subject_content,
            function(result){
              if(result.status === Office.AsyncResultStatus.Succeeded){
                resolve();
              }else{
                console.error("Failed to update email subject with error message: ", result.error);
                reject(result.error);
              }
            }
          );
        });

        return; // Prevent further error handling (if(!response.ok) statement)
      }

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
    flags.classList.remove('disabledFlags');
  }
}





