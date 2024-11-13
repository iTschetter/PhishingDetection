const OfficeAddinMock = require("office-addin-mock");
const myFeature = require("../src/taskpane/helloWorld.js");

// Create the seed mock object.
const mockData = {
  // Identify the host to the mock library (required for Outlook).
  host: "outlook",
  context: {
    mailbox: {
      item: {
        setSelectedDataAsync: function (data) {
          this.data = data;
        },
      },
    },
  },
};

// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in message should be set to 'Hello World'", async function () {
  await myFeature.addHelloWorldText();
  expect(officeMock.context.mailbox.item.data).toBe("Hello World!");
});
