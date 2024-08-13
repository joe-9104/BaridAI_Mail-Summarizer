
// Define long and short strings
let longDescriptionTextFr = "BaridAI utilise une IA améliorée pour résumer les e-mails en toute simplicité. Il vous suffit de cliquer sur le bouton et BaridAI génèrera pour vous un résumé concis de l'e-mail, rationalisant ainsi votre flux de travail et augmentant votre efficacité."
let longDescriptionTextEn = "BaridAI uses enhanced AI to summarize emails with ease. Just click the button, and BaridAI will generate a concise email summary for you, streamlining your workflow and increasing efficiency."
let shortDescriptionTextFr = "BaridAI utilise une IA avancée pour résumer rapidement les e-mails. Gagnez du temps et augmentez la productivité!"
let shortDescriptionTextEn = "BaridAI uses advanced AI to quickly summarize emails. Streamline your workflow and increase productivity!"
let longATextFr = "Voir plus"
let longATextEn = "Show more"
let shortATextFr = "Voir moins"
let shortAtextEn = "Show less"

//Function to show more & less info
function showMoreLess() {
  const descriptionText = document.getElementById('descriptionText');
  const aText = document.getElementById('atext');
  if (aText.innerHTML == longATextEn) {
    aText.innerHTML = shortAtextEn;
    descriptionText.textContent = longDescriptionTextEn;
    descriptionText.classList.remove('collapsed');
    descriptionText.classList.add('expanded');
  } else if (aText.innerHTML == shortAtextEn) {
    aText.innerHTML = longATextEn;
    descriptionText.textContent = shortDescriptionTextEn;
    descriptionText.classList.remove('expanded');
    descriptionText.classList.add('collapsed');
  } else if (aText.innerHTML == shortATextFr) {
    aText.innerHTML = longATextFr;
    descriptionText.textContent = shortDescriptionTextFr;
    descriptionText.classList.remove('expanded');
    descriptionText.classList.add('collapsed');
  } else if (aText.innerHTML == longATextFr) {
    aText.innerHTML = shortATextFr;
    descriptionText.textContent = longDescriptionTextFr;
    descriptionText.classList.remove('collapsed');
    descriptionText.classList.add('expanded');
  }
}

//Function to set the language of the add-in
function setLanguage(lang) {
  const fr_flag = document.getElementById('fr_flag');
  const uk_flag = document.getElementById('uk_flag');
  const headerText = document.getElementById('headerText');
  const descriptionText = document.getElementById('descriptionText');
  const aText = document.getElementById('atext');
  const summaryElement = document.getElementById('summaryText');
  const buttonClick = document.getElementById('buttonClick');

  if (lang == 'fr') {
    fr_flag.title = "Français";
    uk_flag.title = "Basculer vers l'anglais";
    headerText.textContent = "BaridAI - Votre Assistant de Messagerie";
    descriptionText.textContent = shortDescriptionTextFr;
    aText.innerHTML = longATextFr;
    if(summaryElement.innerText === 'Click to summarize...'){
      summaryElement.innerText = "Cliquer pour résumer..."
    }
    buttonClick.innerText = "Résumer l'email"
  } else if (lang === 'en') {
    fr_flag.title = "Switch to french";
    uk_flag.title = "English";
    headerText.textContent = "BaridAI - Your Email Assistant";
    descriptionText.textContent = shortDescriptionTextEn;
    aText.innerHTML = longATextEn;
    if(summaryElement.innerText === 'Cliquer pour résumer...'){
      summaryElement.innerText = "Click to summarize..."
    }
    buttonClick.innerText = "Summarize email"
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

  const item = Office.context.mailbox.item;

  // Define the button and its initial color
  const button = document.getElementById("buttonClick");
  const initialColor = button.style.backgroundColor;

  // Define the placeholder video
  const placeholderMessage = document.getElementById("placeholdermsg");

  // Define the summary text element
  const summaryElement = document.getElementById('summaryText');

  // Define flags
  const frFlag = document.getElementById("fr_flag");
  const ukFlag = document.getElementById("uk_flag");

  try {

    // Disable button and change its color
    button.disabled = true;
    button.style.backgroundColor = '#c2a1c1';

    // Disable flags
    frFlag.classList.add('disabledFlags');
    ukFlag.classList.add('disabledFlags');

    // Show the placehloder video
    placeholderMessage.style.display = 'flex';
    placeholderMessage.scrollIntoView({ behavior : 'smooth'});

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
            body: JSON.stringify({ email_content: emailContent })
          });
          const data = await response.json();

          // Clear the summary text before displaying the new summary
          summaryElement.innerHTML = '';

          // Change the color of the text to black
          summaryElement.style.color = 'black';

          // Split the summary text into words
          const words = data.summary.match(/\S+|[^\S\r\n]+|[\r\n]/g);

          // Display each word with a delay
          for (let i = 0; i < words.length; i++) {
            await new Promise(resolve => setTimeout(resolve, 100)); // Adjust the delay time here
            summaryElement.innerHTML += words[i] + (i < words.length - 1 ? ' ' : '');
          }
        }
        resolve();
      });
    });
  } catch (error) {
    console.error("An error occurred! ", error);
  } finally {
    // Revert button color to the initial color
    button.style.backgroundColor = initialColor;

    // Hide the placeholder video
    placeholderMessage.style.display = 'none';

    // Unable flags
    frFlag.classList.remove('disabledFlags');
    ukFlag.classList.remove('disabledFlags');
  }
}



