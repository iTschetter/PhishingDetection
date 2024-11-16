/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
const { GoogleGenerativeAI } = require("@google/generative-ai"); // importing Google AI SDK
const { GEMINI_API_KEY } = require("../../config.js");
// API key is not pushed to github for security
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

async function analyze(emailContent, metadata) {
  // Medium of communication with Gemini
  try {
    const model = genAI.getGenerativeModel({
      model: "gemini-1.5-flash",
    });

    // Prompt engineering:
    const prompt = `
    You are a cybersecurity expert analyzing an email for phishing attempts.
    Analyze the following email content and metadata for signs of phishing.
    Provide your response in the following format:
    - Confidence Score: (0-100, where 100 is definitely phishing)
    - Suspicious Elements: (bullet point list)
    - Reasoning: (brief explanation)

    Please response with a JSON object with no additional text. The JSON elements should be the three items previously listed.

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
    console.error("Error: ", error);
    return "Error analyzing email";
  }
}

export async function run() {
  // Occurs when the "run" button is pressed

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
      let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
      insertAt.appendChild(label);

      const results = await analyze(result.value, metadata); // Calling Gemini to analyze the email (result.value is the body)
      insertAt.appendChild(document.createElement("br"));
      insertAt.appendChild(document.createTextNode(results)); // Displaying the results from gemini's analysis of the body of the email into the UI (app-body)
      insertAt.appendChild(document.createElement("br"));
    }
  );
}
