const { expect } = require('chai');
const sinon = require('sinon');
const { JSDOM } = require('jsdom');

// Setup DOM environment
const dom = new JSDOM('<!DOCTYPE html><div id="sideload-msg"></div><div id="app-body"></div><div id="item-subject"></div>');
global.document = dom.window.document;
global.window = dom.window;

// Mock Office.js
global.Office = {
  onReady: callback => callback({ host: 'Outlook' }),
  HostType: { Outlook: 'Outlook' },
  context: {
    mailbox: {
      item: {
        from: { emailAddress: 'test@example.com' },
        subject: 'Test Subject',
        attachments: [],
        body: {
          getAsync: null // Will be stubbed in individual tests
        }
      }
    }
  }
};

// Mock console methods
global.console = {
  log: jest.fn(),
  error: jest.fn(),
  warn: jest.fn(),
};

// Mock Google AI
const mockGenerateContent = sinon.stub().resolves({
  response: {
    text: () => JSON.stringify({
      confidence: 85,
      elements: ['Suspicious sender', 'Urgent language'],
      reasoning: 'Test reasoning'
    })
  }
});

const mockGetGenerativeModel = sinon.stub().returns({
  generateContent: mockGenerateContent
});

jest.mock('@google/generative-ai', () => ({
  GoogleGenerativeAI: function() {
    return {
      getGenerativeModel: mockGetGenerativeModel
    };
  }
}));

describe('Email Analysis Add-in', () => {
  let sandbox;

  beforeEach(() => {
    sandbox = sinon.createSandbox();
    // Reset DOM elements
    document.getElementById('item-subject').innerHTML = '';
    // Reset analysis flag
    global.analysisHasOccurred = false;
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe('analyze()', () => {
    it('should return properly formatted analysis results', async () => {
      const emailContent = 'Test email content';
      const metadata = {
        sender: 'test@example.com',
        subject: 'Test Subject',
        hasAttachments: false
      };

      const result = await analyze(emailContent, metadata);
      const parsed = JSON.parse(result);

      expect(parsed).to.have.property('confidence').that.is.a('number');
      expect(parsed).to.have.property('elements').that.is.an('array');
      expect(parsed).to.have.property('reasoning').that.is.a('string');
    });

    it('should handle API errors gracefully', async () => {
      mockGenerateContent.rejects(new Error('API Error'));
      
      const result = await analyze('content', {});
      expect(result).to.equal('Error analyzing email');
    });
  });

  describe('cleanGeminiResponse()', () => {
    it('should clean and parse JSON response correctly', () => {
      const response = '```json\n{"confidence": 85,"elements": [],"reasoning": "test"}\n```';
      const result = cleanGeminiResponse(response);
      
      expect(result).to.deep.equal({
        confidence: 85,
        elements: [],
        reasoning: 'test'
      });
    });

    it('should throw error for invalid JSON', () => {
      const response = '```json\ninvalid json\n```';
      expect(() => cleanGeminiResponse(response)).to.throw();
    });
  });

  describe('run()', () => {
    beforeEach(() => {
      // Mock getAsync to simulate email body retrieval
      Office.context.mailbox.item.body.getAsync = (format, options, callback) => {
        callback({
          value: 'Test email body',
          status: 'succeeded'
        });
      };
    });

    it('should not run analysis if previous analysis is still running', async () => {
      global.analysisHasOccurred = true;
      await run();
      expect(mockGetGenerativeModel.called).to.be.false;
    });

    it('should update DOM with analysis results', async () => {
      await run();
      
      const resultDiv = document.getElementById('item-subject');
      expect(resultDiv.innerHTML).to.include('Confidence Score: 85');
      expect(resultDiv.innerHTML).to.include('Test reasoning');
    });

    it('should handle empty suspicious elements array', async () => {
      mockGenerateContent.resolves({
        response: {
          text: () => JSON.stringify({
            confidence: 85,
            elements: [],
            reasoning: 'Test reasoning'
          })
        }
      });

      await run();
      
      const resultDiv = document.getElementById('item-subject');
      expect(resultDiv.innerHTML).to.not.include('Suspicious Elements:');
    });
  });
});