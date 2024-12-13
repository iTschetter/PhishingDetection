const sinon = require('sinon');
const { JSDOM } = require('jsdom');

// Setup DOM environment
const dom = new JSDOM(`
  <!DOCTYPE html>
  <html>
    <body>
      <div id="sideload-msg"></div>
      <div id="app-body"></div>
      <div id="item-subject"></div>
      <button id="run"></button>
    </body>
  </html>
`);

global.document = dom.window.document;
global.window = dom.window;

// Mock Office.js
global.Office = {
  onReady: callback => callback({ host: 'Outlook' }),
  HostType: { Outlook: 'Outlook' },
  EventType: { ItemChanged: 'itemChanged' },
  context: {
    mailbox: {
      item: {
        from: { emailAddress: 'test@example.com' },
        subject: 'Test Subject',
        attachments: [],
        body: {
          getAsync: (format, options, callback) => {
            callback({
              value: 'Test email body',
              status: 'succeeded'
            });
          }
        }
      },
      addHandlerAsync: (eventType, handler, callback) => {
        if (callback) callback();
        return Promise.resolve();
      }
    }
  }
};

// Mock console to prevent noise in tests
global.console = {
  log: jest.fn(),
  error: jest.fn(),
  warn: jest.fn()
};

// Create standard test response
const testResponse = {
  confidence: 85,
  elements: ['Suspicious sender', 'Urgent language'],
  reasoning: 'Test reasoning'
};

// Mock Gemini AI
const mockGenerateContent = sinon.stub().resolves({
  response: {
    text: () => JSON.stringify(testResponse)
  }
});

const mockGetGenerativeModel = sinon.stub().returns({
  generateContent: mockGenerateContent
});

// Mock the Gemini module
jest.mock('@google/generative-ai', () => ({
  GoogleGenerativeAI: function() {
    return {
      getGenerativeModel: mockGetGenerativeModel
    };
  }
}));

// Import after mocks are set up
const { analyze, cleanGeminiResponse, run } = require('../src/taskpane/taskpane');

describe('Email Analysis Add-in', () => {
  beforeEach(() => {
    // Reset the DOM
    document.getElementById('item-subject').innerHTML = '';
    global.analysisHasOccurred = false;

    // Reset mocks
    mockGenerateContent.reset();
    mockGenerateContent.resolves({
      response: {
        text: () => JSON.stringify(testResponse)
      }
    });
  });

  describe('analyze()', () => {
    it('should analyze email content and return formatted results', async () => {
      const emailContent = 'Test email content';
      const metadata = {
        sender: 'test@example.com',
        subject: 'Test Subject',
        hasAttachments: false
      };

      const result = await analyze(emailContent, metadata);
      const parsed = JSON.parse(result);
      
      expect(parsed).toHaveProperty('confidence', 85);
      expect(parsed.elements).toHaveLength(2);
      expect(parsed.elements).toContain('Suspicious sender');
      expect(typeof parsed.reasoning).toBe('string');
    });

    it('should handle API errors gracefully', async () => {
      mockGenerateContent.rejects(new Error('API Error'));
      
      const result = await analyze('content', {});
      expect(result).toBe('Error analyzing email');
    });

    it('should send correct prompt to AI', async () => {
      const emailContent = 'Test content';
      const metadata = { subject: 'Test' };
      
      await analyze(emailContent, metadata);
      
      expect(mockGenerateContent.called).toBeTruthy();
      const call = mockGenerateContent.getCall(0);
      // Check that the email content is in the prompt
      expect(call.args[0]).toContain(emailContent);
      
      // Check that the metadata is in the prompt with proper formatting
      const formattedMetadata = JSON.stringify(metadata, null, 2);
      expect(call.args[0]).toContain(formattedMetadata);
    });
  });

  describe('cleanGeminiResponse()', () => {
    it('should handle error message responses', () => {
      const result = cleanGeminiResponse('Error analyzing email');
      expect(result).toBe('Error analyzing email');
    });

    it('should clean and parse JSON responses', () => {
      const response = '```json\n{"confidence": 85,"elements": [],"reasoning": "test"}\n```';
      const result = cleanGeminiResponse(response);
      
      expect(result).toEqual({
        confidence: 85,
        elements: [],
        reasoning: 'test'
      });
    });

    it('should throw error for invalid JSON', () => {
      const response = '```json\ninvalid json\n```';
      expect(() => cleanGeminiResponse(response)).toThrow();
    });

    it('should handle responses without code blocks', () => {
      const response = '{"confidence": 85,"elements": [],"reasoning": "test"}';
      const result = cleanGeminiResponse(response);
      
      expect(result).toEqual({
        confidence: 85,
        elements: [],
        reasoning: 'test'
      });
    });
  });

  describe('run()', () => {
    beforeEach(() => {
      // Reset everything before each test
      document.getElementById('item-subject').innerHTML = '';
      global.analysisHasOccurred = false;

      // Setup default Office.js mock
      Office.context.mailbox.item.body.getAsync = (format, options, callback) => {
        callback({
          value: 'Test email body',
          status: 'succeeded'
        });
      };

      // Reset Gemini mock
      mockGenerateContent.reset();
    });

    it('should not run if analysis is already in progress', async () => {
      // Set the flag before running
      global.analysisHasOccurred = true;
      
      // Run the function
      await run();
      
      // Instead of checking mockGenerateContent, check if the DOM was updated
      const resultDiv = document.getElementById('item-subject');
      expect(resultDiv.innerHTML).toBe(''); // Should remain empty if analysis didn't run
    });

    it('should handle error responses', async () => {
      // Setup the mock to reject with an error
      mockGenerateContent.resolves({
        response: {
          text: () => 'Error analyzing email'
        }
      });

      // Run the function
      await run();
      
      // Add a small delay to let async operations complete
      await new Promise(resolve => setTimeout(resolve, 0));
      
      // Check the DOM
      const resultDiv = document.getElementById('item-subject');
      expect(resultDiv.innerHTML).toContain('error');
      
      // Verify the flag was reset
      expect(global.analysisHasOccurred).toBe(false);
    });

    it('should handle successful analysis with no suspicious elements', async () => {
      // Create response with low confidence and no elements
      const noElementsResponse = {
        confidence: 20,
        elements: [],
        reasoning: 'No suspicious elements found'
      };

      // Setup the mock with this specific response
      mockGenerateContent.resolves({
        response: {
          text: () => JSON.stringify(noElementsResponse)
        }
      });

      // Run the function
      await run();
      
      // Add a small delay to let async operations complete
      await new Promise(resolve => setTimeout(resolve, 0));

      // Check the DOM
      const resultDiv = document.getElementById('item-subject');
      const html = resultDiv.innerHTML;

      // Verify expected content
      expect(html).toContain('Risk Confidence Score:');
      expect(html).toContain('Low');
      expect(html).toContain('(20%)');
      expect(html).not.toContain('Suspicious Elements:');
      expect(html).toContain('No suspicious elements found');
    });
  });
});