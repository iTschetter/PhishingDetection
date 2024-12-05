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
let analysisHasOccurred = false;

Office.onReady( (info) => {
  if ( info.host === Office.HostType.Outlook ) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // event handler for item selection
    Office.context.mailbox.addHandlerAsync( Office.EventType.ItemChanged , itemChangedHandler );

    // run ai analysis for the email
    run();
  }
});

// Handler function for when email selection changes
function itemChangedHandler(eventArgs) {
  run();
}

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
      
      const insertAt = document.getElementById("item-subject");
      insertAt.innerHTML = ""; // Clear previous results
      
      const results = await analyze(result.value, metadata);
      const cleaned = cleanGeminiResponse(results);

      if (typeof cleaned === 'string') {
        insertAt.innerHTML = `<div class="error">${cleaned}</div>`;
        analysisHasOccurred = false;
        return;
      }
      
      // Create the HTML structure
      const html = `
        <div class="results">
          <div class="sectionContainer">
            <div class="containerTitleSpecial">Risk Confidence Score:</div>
            <div class="confidenceScore ${cleaned.confidence >= 75 ? 'highRisk' : cleaned.confidence >= 50 ? 'mediumRisk' : 'lowRisk'}">
              ${cleaned.confidence >= 75 ? 'High' : cleaned.confidence >= 50 ? 'Medium' : 'Low'} (${cleaned.confidence}%)
            </div>
          </div>

          ${cleaned.elements.length > 0 ? `
            <div class="sectionContainer">
              <div class="containerTitle">Suspicious Elements:</div>
              <ul class="elementsList">
                ${cleaned.elements.map(element => `<li>${element}</li>`).join('')}
              </ul>
            </div>
          ` : ''}

          <div class="sectionContainer">
            <div class="containerTitle">Analysis:</div>
            <div class="AIAnalysisText">${cleaned.reasoning}</div>
          </div>
        </div>
      `;
      
      insertAt.innerHTML = html;
      analysisHasOccurred = false;
    }
  );
}
