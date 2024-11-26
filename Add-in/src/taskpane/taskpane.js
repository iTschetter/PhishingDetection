/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
const { GoogleGenerativeAI } = require("@google/generative-ai"); // importing Google AI SDK
// API key is not pushed to github for security
// TODO Figure out a solution for production
// Dotenv doesn't work in broswer
const { GEMINI_API_KEY } = require("../../config.js");
const genAI = new GoogleGenerativeAI(GEMINI_API_KEY); // Creates a new instance, using our API key, of the Gemini AI


Office.onReady((info) => {
  // Occurs when everything is fully loaded (i.e. when ready)
  if (info.host === Office.HostType.Outlook) {
    // Ensuring the host application is outlook
    document.getElementById("sideload-msg").style.display = "none"; // Hides sideload message
    document.getElementById("app-body").style.display = "flex"; // Makes the main app body visible with flex display
    document.getElementById("run").onclick = run; // Sets up a button titled "run"
  }
});

export async function analyze(emailContent, metadata) {
  // Medium of communication with Gemini
  try {
    const model = genAI.getGenerativeModel({
      model: "gemini-1.5-flash",
    });

    // Prompt engineering:
    const prompt = `
    For each email provided, analyze it and return a JSON response with exactly these fields:
    - confidence: integer from 0-100 representing how confident you are that this is a scam
    - elements: array of strings listing specific elements that indicate potential fraud
    - reasoning: string explaining your detailed analysis of why these elements are suspicious

    Format all responses as valid JSON objects with these exact field names. Make sure you don't return a 

    Example response:
    {
        "confidence": 85,
        "elements": [
            "Urgency in subject line",
            "Grammatical errors",
            "Request for sensitive information",
            "Suspicious sender address"
        ],
        "reasoning": "The email exhibits multiple red flags typical of scam attempts. The urgent subject line creates pressure to act quickly. Multiple grammatical errors suggest non-native English speakers. The request for sensitive banking information is never legitimate from real institutions. The sender address mimics but doesn't exactly match the real company domain."
    }

     Email Metadata:
     ${JSON.stringify(metadata, null, 2)}

     Email Content:
     ${emailContent}
     `;

    const result = await model.generateContent(prompt); // Prompting the AI
    const response = await result.response; // Capturing it's response
    return response.text();
  } catch (error) {
    // Error handling
    console.error("Error: ", error.message);
    return "Error analyzing email";
  }
}

export function cleanGeminiResponse(response) {

  if (response === 'Error analyzing email') {
    return response;
  }

  // Removing backticks and 'json' identifier
  let cleaning = response.replace(/```json/g, '').replace(/```/g, '').trim();
  
  // Handling any potential leading or trailing whitespace and newlines
  cleaning = cleaning.replace(/^\s+|\s+$/g, '');
  
  try {
      // Parsing the cleaned string into a JSON object
      const cleaned = JSON.parse(cleaning);
      return cleaned;
  } catch (error) {
      console.error('Error parsing JSON:', error);
      throw error;
  }
}

/* eslint-disable node/no-unsupported-features/es-syntax */
export async function run() {
  // Occurs when the "run" button is pressed
  if (analysisHasOccurred) {
    return;
  }
  analysisHasOccurred = true;

  const item = Office.context.mailbox.item; // Current email item selected by user
  const metadata = {
    // Grabs the address, subject, and whether or not there is attachments
    sender: item.from?.emailAddress,
    subject: item.subject,
    hasAttachments: item.attachments.length > 0,
  };

  Office.context.mailbox.item.body.getAsync(
    // Grabbing the selected email
    "text",
    { asyncContext: "This is passed to the callback" },
    async function callback(result) {
      // Passing the email as "result"
      
      let insertAt = document.getElementById("item-subject");
      insertAt.innerHTML = ""; //Clear previous results

      let label = document.createElement("b").appendChild(document.createTextNode("AI Analysis: "));
      insertAt.appendChild(label);
      let results = await analyze(result.value, metadata); // Calling Gemini to analyze the email (result.value is the body)
      const cleaned = cleanGeminiResponse(results);

      // Add defensive checks
      if (typeof cleaned === 'string') {
        // It's an error message
        insertAt.appendChild(document.createTextNode(cleaned));
        analysisHasOccurred = false;
        return;
      }

      insertAt.appendChild(document.createElement("br"));
      
      // Check if confidence exists
      if (cleaned.confidence !== undefined) {
        insertAt.appendChild(document.createTextNode(`Confidence Score: ${cleaned.confidence}`));
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createElement("br"));
      }

      // Check if elements exists and is an array
      if (cleaned.elements && Array.isArray(cleaned.elements) && cleaned.elements.length > 0) {
        insertAt.appendChild(document.createTextNode(`Suspicious Elements: `));
        const ul = document.createElement("ul");
        cleaned.elements.forEach(item => {
          const lst = document.createElement('li');
          lst.textContent = item;
          ul.appendChild(lst);
        });
        insertAt.appendChild(ul);
        insertAt.appendChild(document.createElement("br"));
      }

      // Check if reasoning exists
      if (cleaned.reasoning) {
        insertAt.appendChild(document.createTextNode(`Reason: ${cleaned.reasoning}`));
        insertAt.appendChild(document.createElement("br"));
      }

      analysisHasOccurred = false;
    }
  );
}
