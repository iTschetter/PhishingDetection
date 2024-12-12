/* eslint-disable node/no-unsupported-features/es-syntax */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
const { GoogleGenerativeAI } = require("@google/generative-ai"); // importing Google AI SDK
// API key is not pushed to github for security
// TODO Figure out a solution for production
// Dotenv doesn't work in broswer
// Global Vars:
const { GEMINI_API_KEY } = require("../../config.js");
const genAI = new GoogleGenerativeAI(GEMINI_API_KEY); // Creates a new instance, using our API key, of the Gemini AI
let analysisHasOccurred = false;

Office.onReady((info) => {
  // Startup code (this is what begins everything  for our program!!!)
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none"; // Hiding the sideload-msg!! ()
    document.getElementById("app-body").style.display = "flex";

    // event handler for item selection
    // Sets up for the handler function later
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChangedHandler);

    // run ai analysis for the email
    run();
  }
});

// Handler function for when email selection changes
function itemChangedHandler() {
  run();
}

// Sends the emails info to Gemini and receives its response:
export async function analyze(emailContent, metadata) {
  // Medium of communication with Gemini
  try {
    const model = genAI.getGenerativeModel({
      model: "gemini-1.5-flash",
    }); // We found this to be the best model for our project (we are using a free account)

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
    // Bug/Limit of our system: occurs often when Gemini is overloaded
    // Error handling
    console.error("Error: ", error.message);
    return "Error analyzing email";
  }
}
// Gemini responds with a string, so we have to clean it up to turn it into a JSON object:
export function cleanGeminiResponse(response) {
  if (response === "Error analyzing email") {
    return response;
  }

  // Removing backticks (from the front and back) and 'json' letters from the string Gemini returns to us
  let cleaning = response
    .replace(/```json/g, "")
    .replace(/```/g, "")
    .trim();

  // Matches and removes all whitespace/new lines by looking at the start (denoted by '^\s+') or the end of the string (denoted by '\s+$')
  cleaning = cleaning.replace(/^\s+|\s+$/g, "");

  try {
    // Parsing the cleaned string into a JSON object
    const cleaned = JSON.parse(cleaning);
    return cleaned;
  } catch (error) {
    console.error("Error parsing JSON:", error);
    throw error;
  }
}

/* eslint-disable node/no-unsupported-features/es-syntax */
export async function run() {
  // This is where the magic happens!
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

      if (typeof cleaned === "string") {
        // checking to see if the JSON conversion failed (throwing error if it did)
        insertAt.innerHTML = `<div class="error">${cleaned}</div>`;
        analysisHasOccurred = false;
        return;
      }

      // using the following html as our app-bodys structure (it will be inserted right after this)
      // Using div tags to 'divide' each section of our app (using classes to connect to our CSS elements)
      const html = `
        <div class="results">

          <div class="sectionContainer">

            <div class="containerTitleSpecial">Risk Confidence Score:</div>
            <div class="confidenceScore ${cleaned.confidence >= 75 ? "highRisk" : cleaned.confidence >= 50 ? "mediumRisk" : "lowRisk"}">
              ${cleaned.confidence >= 75 ? "High" : cleaned.confidence >= 50 ? "Medium" : "Low"} (${cleaned.confidence}%)

            </div>
          </div>

          ${
            cleaned.elements.length > 0
              ? `
            <div class="sectionContainer">

              <div class="containerTitle">Suspicious Elements:</div>
              <ul class="elementsList">
                ${cleaned.elements.map((element) => `<li>${element}</li>`).join("")}
              </ul>

            </div>
          `
              : ""
          }

          <div class="sectionContainer">

            <div class="containerTitle">Analysis:</div>
            <div class="AIAnalysisText">${cleaned.reasoning}</div>

          </div>

        </div>
      `;

      insertAt.innerHTML = html; // inserting the above HTML
      analysisHasOccurred = false; // Global var (may need to be deleted)
    }
  );
}
