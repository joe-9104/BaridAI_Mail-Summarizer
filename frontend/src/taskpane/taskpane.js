
// Define language-specific strings
const languageStrings = {
  french: {
      fr_flag_title: "Français",
      uk_flag_title: "Basculer vers l'anglais",
      ar_flag_title: "Basculer vers l'arabe",
      es_flag_title: "Basculer vers l'espagnol",
      pr_flag_title: "Basculer vers le portugais",
      de_flag_title: "Basculer vers l'allemand",
      headerText: "BaridAI - Votre Assistant de Messagerie",
      descriptionText: "BaridAI utilise une IA avancée pour résumer rapidement les e-mails. Gagnez du temps et augmentez la productivité!",
      longDescriptionText: "BaridAI utilise une IA améliorée pour résumer les e-mails en toute simplicité. Il vous suffit de cliquer sur le bouton et BaridAI génèrera pour vous un résumé concis de l'e-mail, rationalisant ainsi votre flux de travail et augmentant votre efficacité",
      aText: "Voir plus",
      oppositeAText: "Voir moins",
      buttonClick_text: "Résumer l'email",
      summary_placeholder: "Résumer en français...",
      opposite_summary_placeholder: "Summarize in english...",
  },
  english: {
      fr_flag_title: "Switch to french",
      uk_flag_title: "English",
      ar_flag_title: "Switch to arabic",
      es_flag_title: "Switch to spanish",
      pr_flag_title: "Switch to protuguese",
      de_flag_title: "Switch to german",
      headerText: "BaridAI - Your Email Assistant",
      descriptionText: "BaridAI uses advanced AI to quickly summarize emails. Streamline your workflow and increase productivity!",
      longDescriptionText: "BaridAI uses enhanced AI to summarize emails with ease. Just click the button and BaridAI will generate a concise email summary for you, streamlining your workflow and increasing your efficiency",
      aText: "Show more",
      oppositeAText: "Show less",
      buttonClick_text: "Summarize email",
      summary_placeholder: "Summarize in english...",
      opposite_summary_placeholder: "Résumer en français...",
  },
  arabic: {
      fr_flag_title: "التبديل إلى الفرنسية",
      uk_flag_title: "التبديل إلى الإنجليزية",
      ar_flag_title: "العربية",
      es_flag_title: "التبديل إلى الإسبانية",
      pr_flag_title: "التبديل إلى البرتغالية",
      de_flag_title: "التبديل إلى الألمانية",
      headerText: "BaridAI - مساعد البريد الإلكتروني الخاص بك",
      descriptionText: "BaridAI يستخدم الذكاء الاصطناعي المتقدم لتلخيص رسائل البريد الإلكتروني بسرعة. قم بتبسيط سير العمل وزيادة الإنتاجية!",
      longDescriptionText: "يستخدم BaridAI الذكاء الاصطناعي المحسّن لتلخيص رسائل البريد الإلكتروني بسهولة. ما عليك سوى النقر على الزر وسيقوم BaridAI بإنشاء ملخص موجز عبر البريد الإلكتروني لك، مما يؤدي إلى تبسيط سير عملك وزيادة كفاءتك",
      aText: "عرض المزيد",
      oppositeAText: "عرض اقل",
      buttonClick_text: "تلخيص البريد الإلكتروني",
      summary_placeholder: "تلخيص باللغة العربية...",
      opposite_summary_placeholder: "Summarize in english...",
      rtl: true
  },
  spanish: {
      fr_flag_title: "Cambiar a francés",
      uk_flag_title: "Cambiar a inglés",
      ar_flag_title: "Cambiar a árabe",
      es_flag_title: "Español",
      pr_flag_title: "Cambiar a portugués",
      de_flag_title: "Cambiar a alemán",
      headerText: "BaridAI - Su asistente de correo electrónico",
      descriptionText: "BaridAI utiliza avanzada AI para resumir correos electrónicos rápidamente. ¡Simplifique su flujo de trabajo y aumente la productividad!",
      longDescriptionText: "BaridAI utiliza IA mejorada para resumir correos electrónicos con facilidad. Simplemente haga clic en el botón y BaridAI generará un resumen conciso por correo electrónico, optimizando su flujo de trabajo y aumentando su eficiencia.",
      aText: "Ver más",
      oppositeAText: "mostrar menos",
      buttonClick_text: "Resumir el correo electrónico",
      summary_placeholder: "Resumir en español...",
      opposite_summary_placeholder: "Summarize in english...",
  },
  portuguese: {
      fr_flag_title: "Mudar para francês",
      uk_flag_title: "Mudar para inglês",
      ar_flag_title: "Mudar para árabe",
      es_flag_title: "Mudar para espanhol",
      pr_flag_title: "Português",
      de_flag_title: "Mudar para alemão",
      headerText: "BaridAI - Seu Assistente de Email",
      descriptionText: "BaridAI utiliza IA avançada para resumir emails rapidamente. Simplifique seu fluxo de trabalho e aumente a produtividade!",
      longDescriptionText: "BaridAI usa IA aprimorada para resumir e-mails com facilidade. Basta clicar no botão e BaridAI irá gerar um resumo conciso por e-mail para você, agilizando seu fluxo de trabalho e aumentando sua eficiência",
      aText: "Mostrar mais",
      oppositeAText: "mostrar menos",
      buttonClick_text: "Resumir o email",
      summary_placeholder: "Resumir em português...",
      opposite_summary_placeholder: "Summarize in english...",
  },
  german: {
      fr_flag_title: "Wechseln Sie zu Französisch",
      uk_flag_title: "Zu Englisch wechseln",
      ar_flag_title: "Zu Arabisch wechseln",
      es_flag_title: "Zu Spanisch wechseln",
      pr_flag_title: "Zu Portugiesisch wechseln",
      de_flag_title: "Deutsch",
      headerText: "BaridAI - Ihr E-Mail-Assistent",
      descriptionText: "BaridAI verwendet fortschrittliche KI, um E-Mails schnell zusammenzufassen. Optimieren Sie Ihren Workflow und steigern Sie die Produktivität!",
      longDescriptionText: "BaridAI nutzt erweiterte KI, um E-Mails einfach zusammenzufassen. Klicken Sie einfach auf die Schaltfläche und BaridAI erstellt für Sie eine prägnante E-Mail-Zusammenfassung, die Ihren Arbeitsablauf optimiert und Ihre Effizienz steigert",
      aText: "Mehr anzeigen",
      oppositeAText: "weniger zeigen",
      buttonClick_text: "E-Mail zusammenfassen",
      summary_placeholder: "Zusammenfassen auf Deutsch...",
      opposite_summary_placeholder: "Summarize in english...",
  }
};

// Define current language
let currentLanguage = 'english'; // Default language
// Define summary language
let summaryLanguage = 'english';// Default language

// Function to show more & less info
function showMoreLess() {
  const descriptionText = document.getElementById('descriptionText');
  const aText = document.getElementById('atext');

  // Get the current language strings
  const langData = languageStrings[currentLanguage];

  if (aText.innerHTML == langData.aText) {
    aText.innerHTML = langData.oppositeAText; // Switching between "Show more" and "Show less"
    descriptionText.textContent = langData.longDescriptionText; // Show the long description
    descriptionText.classList.remove('collapsed');
    descriptionText.classList.add('expanded');
  } else {
    aText.innerHTML = langData.aText; // Revert to the original "Show more" text
    descriptionText.textContent = langData.descriptionText; // Show the short description
    descriptionText.classList.remove('expanded');
    descriptionText.classList.add('collapsed');
  }
}

// Function to toggle flags
async function toggleFlags() {
  const flagsDiv = await document.getElementById('flags');
  const toggleText = await document.getElementById('toggleFlags');

  if (flagsDiv.style.display === 'none' || flagsDiv.style.display === '') {
      flagsDiv.style.display = 'block';
      toggleText.innerText = 'Hide change language menu';
  } else {
      flagsDiv.style.display = 'none';
      toggleText.innerText = 'Change language';
  }
}


//Function to set the language of the add-in
function setLanguage(lang) {

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
  const summaryElement = document.getElementById('summaryText');
  const buttonClick = document.getElementById('buttonClick');

  if (!languageStrings[lang]) return; // Exit if language not supported

  const langData = languageStrings[lang];

  fr_flag.title = langData.fr_flag_title;
  uk_flag.title = langData.uk_flag_title;
  ar_flag.title = langData.ar_flag_title;
  es_flag.title = langData.es_flag_title;
  pr_flag.title = langData.pr_flag_title;
  de_flag.title = langData.de_flag_title;
  headerText.textContent = langData.headerText;
  descriptionText.textContent = langData.descriptionText;
  aText.innerHTML = langData.aText;
  // Change the summaryElement text only if it matches one of the summary_placeholders
  const isPlaceholderText = Object.values(languageStrings).some(langObj => langObj.summary_placeholder === summaryElement.innerText);

  if (isPlaceholderText) {
      summaryElement.innerText = langData.summary_placeholder;
  }
  buttonClick.innerText = langData.buttonClick_text;
  buttonClick.disabled = (summaryLanguage === currentLanguage && summaryElement.innerText !== langData.summary_placeholder);

  if (langData.rtl) {
      document.body.style.direction = 'rtl';
  } else {
      document.body.style.direction = 'ltr';
  }
}

//Main function: summarizing the email content
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    // Office is ready
    console.log("BaridAI - Mail Summarizer is ready...")
  }
});

async function summarizeEmail() {

  const item = Office.context.mailbox.item;

  // Define the button and its initial color
  const button = document.getElementById("buttonClick");
  const initialColor = button.style.backgroundColor;

  // Define the placeholder video
  const placeholderMessage = document.getElementById("placeholdermsg");

  // Define the summary text element
  const summaryElement = document.getElementById('summaryText');

  // Define flags
  const flags = document.getElementById("flags");

  try {

    // Disable button and change its color
    button.disabled = true;
    button.style.backgroundColor = '#95d7ef';

    // Disable flags
    flags.classList.add('disabledFlags');

    // Focus on the summary element
    summaryElement.classList.add('focused');

    // Show the placehloder video
    placeholderMessage.style.display = 'flex';
    placeholderMessage.scrollIntoView({ behavior : 'smooth'});

    //Stock the summary language
    summaryLanguage = currentLanguage;

    // Ensure the UI updates before the next async operation
    await new Promise(resolve => setTimeout(resolve, 0));

    // Get the body of the email
    await new Promise(resolve => {
      item.body.getAsync("text", async function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const emailContent = result.value;

          // Call backend to get the summary
          const response = await fetch('http://localhost:8000/summarize', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({ email_content: emailContent, language: currentLanguage })
          });

          if(response.ok){
            // Successful response
            const data = await response.json();

            // Clear the summary text before displaying the new summary
            summaryElement.innerText = '';

            // Change the color of the text to black
            summaryElement.style.color = 'black';

            // Split the summary text into words (regular expression to match each sequence of non-whitespaced characters followed by a whitespaced one)
            const words = data.summary.match(/\S+|[^\S\r\n]+|[\r\n]/g);

            // Display each word with a delay
            for (let i = 0; i < words.length; i++) {
              await new Promise(resolve => setTimeout(resolve, 25)); // Adjust the delay time here
              summaryElement.innerText += words[i] + (i < words.length - 1 ? ' ' : '');
            }
          }else{
            const errorData = await response.json();
            summaryElement.style.color = 'red';
            summaryElement.innerText = 'BaridAI faced problems. Please try again later or check the console for more info.'
            console.error("Backend error: ", errorData.error)
          }
        }
        resolve();
      });
    });
  } catch (error) {
    summaryElement.style.color = 'red';
    summaryElement.innerText = 'Sorry :(, an error occured. Check the console for more info.'
    console.error("Frontend error: \nAn error occurred! ", error);
  } finally {
    // Revert button color to the initial color
    button.style.backgroundColor = initialColor;

    // Hide the placeholder video
    placeholderMessage.style.display = 'none';

    // Unfocus on the summary element
    summaryElement.classList.remove('focused');

    // Unable flags
    flags.classList.remove('disabledFlags');
  }
}



