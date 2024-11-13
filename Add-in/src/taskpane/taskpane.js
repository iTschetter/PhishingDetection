/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

const GEMINI_API_KEY = 'YOUR_API_KEY_HERE';
const { GoogleGenerativeAI } = require("@google/generative-ai");
const genAI = new GoogleGenerativeAI(GEMINI_API_KEY);

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function analyze(emailContent, metadata) {
  try {
    const model = genAI.getGenerativeModel({ model: "gemini-pro" });

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

     const result = await model.generalContent(prompt);
     const response = await result.response;
     return response.text();
  } catch (error) {
    console.error("Error: ", error);
    return "Error analyzing email";
  }
}

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;

  const metadata = {
    sender: item.from?.emailAddress,
    subject: item.subject,
    hasAttachments: item.attachments.length > 0
  };

  Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: "This is passed to the callback" },
    async function callback(result) {
      // Do something with the result.

      let insertAt = document.getElementById("item-subject");
      let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
      insertAt.appendChild(label);
      insertAt.appendChild(document.createElement("br"));
      insertAt.appendChild(document.createTextNode(result.value));
      insertAt.appendChild(document.createElement("br"));

      const results = await analyze(result.value, metadata);

      // DO SOMETHING WITH RESULTS

    }
  );
}
