# group-projects-group-3

# Getting the repo up and running
1. Install Vs code https://code.visualstudio.com/
1. Install Node.js https://nodejs.org/en/download/prebuilt-installer
1. Install npm https://docs.npmjs.com/downloading-and-installing-node-js-and-npm
1. Open repo in VS code
1. Open terminal
1.  > `git clone https://github.com/dsu-cs/group-projects-group-3.git`
1.  > `git checkout dev`
1.  > `git pull`
1.  > `cd Add-in`
1.  > `npm install`
1.  > In the Add-in directory, create a config.js file with your new API key. It should look like this
```// config.js
const config = {
  GEMINI_API_KEY: "API KEY GOES HERE",
};

module.exports = config;
```
1.  > `npm run start` (Make sure you do this when in the Add-in directory)
1. Open outlook client or https://outlook.office.com/mail/
1. Select an email
1. Locate the App button or drop down option
   
     ![Screenshot 2024-12-13 142558](https://github.com/user-attachments/assets/79a8f5bd-e7a6-4195-b745-fdfb843d5794)
1. Select Phishnet
   
     ![image](https://github.com/user-attachments/assets/d3a03761-405a-4b33-bbdf-26bc639758b3)
1. Select Show Taskpane
   
     ![image](https://github.com/user-attachments/assets/f9566393-f6e1-4b25-865e-e6aacc1eb835)
    

### Setting up the test suite

Ensure that you have installed jest, chai, sinon, and jsdom. To do so, run the following command:
`npm install --save-dev jest chai sinon jsdom`.
