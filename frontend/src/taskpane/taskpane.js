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

//Main function: summarizing the email content
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
      // Office is ready
      console.log("Summarizer is ready...")
  }
});

async function summarizeEmail() {
  try{
    const item = Office.context.mailbox.item;

    // Show the placeholder video
    placeholderMessage = document.getElementById('placeholdermsg');
    placeholderMessage.style.display = 'flex';
    placeholderMessage.scrollIntoView({ behavior : 'smooth'});
    
    // Get the body of the email
    item.body.getAsync("text", function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailContent = result.value;

            // Call backend to get the summary
            fetch('http://localhost:5000/summarize', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ email_content: emailContent })
            })
            .then(response => response.json())
            .then(data => {
                document.getElementById('summaryText').innerText = data.summary;
            })
            .catch(error => console.error('Error:', error));
        }
    }); 
  }catch(error){
    console.log("an error occured!")
  }finally{
    // Hide the placeholder video
    placeholderMessage.style.display = 'none';
  }
}


