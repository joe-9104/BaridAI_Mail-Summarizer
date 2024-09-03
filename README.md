# BaridAI - Mail Summarizer

**BaridAI - Your Intelligent Mail Assistant**

**BaridAI** is an *Outlook add-in* designed to streamline email summarization, leveraging advanced AI technologies. This add-in is highly customizable, supporting multiple languages and offering a user-friendly interface to boost productivity and save time.


***Features***
- *Quick Summaries*: **BaridAI** provides concise summaries of emails, making it easier to grasp the content without reading lengthy texts.
- *Dynamic Display*: The summary text appears dynamically, word by word, creating a smooth and engaging reading experience.
- *Multi-Language Support*: The add-in supports multiple languages, including English, French, Arabic, Spanish, Portuguese, and German, enabling users to compose emails in their native language. Users can switch between languages seamlessly using a language selector with flag icons representing each language.
- *Intuitive Interface*: The add-in features a clean, responsive design that adjusts to different screen sizes, ensuring a consistent user experience across devices.
- *Customizable UI/UX*: Elements such as text alignment (right-to-left for Arabic), and dynamic button states are tailored to enhance usability.


***Technology Stack***
- *Frontend*: The add-in's interface is crafted with HTML, CSS, and JavaScript, delivering a seamless and responsive user experience directly within Outlook.
- *Backend*: A Flask-based Python server powers the backend, utilizing the genai library to interact with AI models for email content generation and summarization.
- *AI Models*: The project employs Google's Gemini 1.5 Flash model to generate both the body and subject of emails.


***Compatibility***

**BaridAI** is compatible with the following versions of Outlook:
- Outlook *2013* or later on *Windows* and *Mac*
- Outlook *2016* or later on *Windows* and *Mac*
- Outlook *2019* or later on *Windows* and *Mac*
- Outlook on the *web*
- Outlook on Windows and Mac (*Microsoft 365*)

***Installation & Usage***

You can install **BaridAI** from the Office Store. It will be automatically deployed in your Outlook environment.
