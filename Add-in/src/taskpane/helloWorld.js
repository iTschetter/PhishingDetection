/* eslint-disable @typescript-eslint/no-unused-vars */
/* global document, Office */

const helloWorld = {
  addHelloWorldText: async () => {
    Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
  },
};

module.exports = helloWorld;
