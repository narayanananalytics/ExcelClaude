import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { perplexityService, PerplexityMessage } from '../../services/perplexityService';
import { ExcelHelpers } from '../../utils/excelHelpers';
import { VBAHelpers } from '../../utils/vbaHelpers';

interface Message extends PerplexityMessage {
  timestamp: Date;
}

const App: React.FC = () => {
  const [apiKey, setApiKey] = useState<string>('');
  const [isConfigured, setIsConfigured] = useState<boolean>(false);
  const [messages, setMessages] = useState<Message[]>([]);
  const [inputMessage, setInputMessage] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [success, setSuccess] = useState<string>('');
  const [pastedImage, setPastedImage] = useState<string | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    // Check if API key is already stored
    const storedKey = localStorage.getItem('perplexityApiKey');
    if (storedKey) {
      setApiKey(storedKey);
      perplexityService.setApiKey(storedKey);
      setIsConfigured(true);
    }
  }, []);

  useEffect(() => {
    // Auto-scroll to bottom when messages change
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  const handleSetApiKey = () => {
    if (!apiKey.trim()) {
      setError('Please enter an API key');
      return;
    }

    try {
      perplexityService.setApiKey(apiKey);
      localStorage.setItem('perplexityApiKey', apiKey);
      setIsConfigured(true);
      setSuccess('API key configured successfully!');
      setTimeout(() => setSuccess(''), 3000);
    } catch (err) {
      setError('Failed to configure API key');
    }
  };

  const handlePaste = async (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    const items = e.clipboardData?.items;
    if (!items) return;

    for (let i = 0; i < items.length; i++) {
      const item = items[i];

      // Check if the item is an image
      if (item.type.indexOf('image') !== -1) {
        e.preventDefault();
        const blob = item.getAsFile();
        if (blob) {
          const reader = new FileReader();
          reader.onload = (event) => {
            const base64Image = event.target?.result as string;
            setPastedImage(base64Image);
            setSuccess('Image pasted! Send your message to include it.');
            setTimeout(() => setSuccess(''), 3000);
          };
          reader.readAsDataURL(blob);
        }
        break;
      }
    }
  };

  const handleRemoveImage = () => {
    setPastedImage(null);
  };

  const handleSendMessage = async () => {
    if ((!inputMessage.trim() && !pastedImage) || isLoading) return;

    let messageContent = inputMessage;

    // If there's an image, add context about it
    if (pastedImage) {
      messageContent = `${inputMessage}\n\n[Image pasted - Note: Image analysis requires Perplexity API vision support. Describe what you see or what you want to do with Excel based on this image.]`;
    }

    const userMessage: Message = {
      role: 'user',
      content: messageContent,
      timestamp: new Date()
    };

    setMessages(prev => [...prev, userMessage]);
    setInputMessage('');
    setPastedImage(null);
    setIsLoading(true);
    setError('');

    try {
      const response = await perplexityService.sendMessage(inputMessage || 'Analyze this image and generate VBA code');
      const assistantMessage: Message = {
        role: 'assistant',
        content: response,
        timestamp: new Date()
      };
      setMessages(prev => [...prev, assistantMessage]);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to get response from Perplexity';
      setError(errorMessage);
    } finally {
      setIsLoading(false);
    }
  };

  const handleKeyPress = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  const handleClearChat = () => {
    setMessages([]);
    perplexityService.clearHistory();
    setSuccess('Chat history cleared');
    setTimeout(() => setSuccess(''), 3000);
  };

  const handleCreateOHLCChart = async () => {
    if (!isConfigured) {
      setError('Please configure your API key first');
      return;
    }

    setIsLoading(true);
    setError('');
    try {
      const rangeAddress = await ExcelHelpers.getSelectedRange();

      // Ask Perplexity to generate VBA code for OHLC chart
      const prompt = `Generate a VBA macro to create a properly formatted OHLC candlestick chart using the data in range ${rangeAddress}. The chart should have green candles for bullish periods and red candles for bearish periods.`;

      const userMessage: Message = {
        role: 'user',
        content: prompt,
        timestamp: new Date()
      };

      setMessages(prev => [...prev, userMessage]);

      const response = await perplexityService.sendMessage(prompt);
      const assistantMessage: Message = {
        role: 'assistant',
        content: response,
        timestamp: new Date()
      };
      setMessages(prev => [...prev, assistantMessage]);

      setSuccess('VBA code generated! Copy and run it in Excel (Alt+F11)');
      setTimeout(() => setSuccess(''), 5000);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Unknown error';
      setError('Failed to generate OHLC chart code: ' + errorMessage);
    } finally {
      setIsLoading(false);
    }
  };

  const handleInsertSampleData = async () => {
    setIsLoading(true);
    setError('');
    try {
      const rangeAddress = await ExcelHelpers.insertSampleOHLCData();
      setSuccess(`Sample OHLC data inserted at ${rangeAddress}`);
      setTimeout(() => setSuccess(''), 3000);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Unknown error';
      setError('Failed to insert sample data: ' + errorMessage);
    } finally {
      setIsLoading(false);
    }
  };

  const handleGetSelectedRange = async () => {
    try {
      const rangeAddress = await ExcelHelpers.getSelectedRange();
      setInputMessage(`Help me with the data in range ${rangeAddress}`);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Unknown error';
      setError('Failed to get selected range: ' + errorMessage);
    }
  };

  const handleCopyCode = async (code: string) => {
    const success = await VBAHelpers.copyToClipboard(code);
    if (success) {
      setSuccess('VBA code copied to clipboard!');
      setTimeout(() => setSuccess(''), 3000);
    } else {
      setError('Failed to copy code to clipboard');
      setTimeout(() => setError(''), 3000);
    }
  };

  const formatMessageContent = (content: string) => {
    // Enhanced formatting for code blocks with copy button
    const parts = content.split(/(```[\s\S]*?```)/g);
    return parts.map((part, index) => {
      if (part.startsWith('```') && part.endsWith('```')) {
        // Extract language and code
        const match = part.match(/```(\w+)?\n([\s\S]*?)```/);
        const language = match?.[1] || '';
        const code = match?.[2] || part.slice(3, -3);

        const isVBA = language.toLowerCase() === 'vba';

        return (
          <div key={index} className="code-block-container">
            {isVBA && (
              <div className="code-block-header">
                <span className="code-language">VBA Macro</span>
                <button
                  className="copy-code-btn"
                  onClick={() => handleCopyCode(code.trim())}
                >
                  📋 Copy Code
                </button>
              </div>
            )}
            <pre className="code-block">
              <code>{code.trim()}</code>
            </pre>
            {isVBA && (
              <div className="vba-instructions">
                <details>
                  <summary>How to run this macro in Excel</summary>
                  <div className="instructions-content">
                    {VBAHelpers.getExecutionInstructions()}
                  </div>
                </details>
              </div>
            )}
          </div>
        );
      }
      return <span key={index}>{part}</span>;
    });
  };

  return (
    <div className="app-container">
      <div className="app-header">
        <h1>Excel Perplexity Assistant</h1>
        <p>AI-powered VBA macro generator for Excel automation</p>
      </div>

      {!isConfigured && (
        <div className="api-key-section">
          <h3>Configure Perplexity API Key</h3>
          <div className="api-key-input-group">
            <input
              type="password"
              placeholder="Enter your Perplexity API key..."
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
              onKeyPress={(e) => e.key === 'Enter' && handleSetApiKey()}
            />
            <button onClick={handleSetApiKey}>Set API Key</button>
          </div>
          <div className="api-key-info">
            Get your API key from perplexity.ai/settings/api
          </div>
        </div>
      )}

      {error && <div className="error-message">{error}</div>}
      {success && <div className="success-message">{success}</div>}

      <div className="quick-actions">
        <h3>Quick Actions</h3>
        <div className="action-buttons">
          <button
            className="action-button"
            onClick={handleInsertSampleData}
            disabled={isLoading}
          >
            Insert Sample OHLC Data
          </button>
          <button
            className="action-button"
            onClick={handleCreateOHLCChart}
            disabled={isLoading || !isConfigured}
          >
            Generate OHLC Chart VBA
          </button>
          <button
            className="action-button primary"
            onClick={handleGetSelectedRange}
            disabled={isLoading}
          >
            Ask Perplexity about Selected Range
          </button>
        </div>
      </div>

      <div className="chat-section">
        <div className="chat-header">
          <h3>Chat with Perplexity</h3>
          {messages.length > 0 && (
            <button className="clear-chat-btn" onClick={handleClearChat}>
              Clear Chat
            </button>
          )}
        </div>

        <div className="chat-messages">
          {messages.map((msg, index) => (
            <div key={index} className={`message ${msg.role}`}>
              {formatMessageContent(msg.content)}
            </div>
          ))}
          {isLoading && (
            <div className="loading-indicator">
              <span>Perplexity is thinking</span>
              <div className="loading-dots">
                <span></span>
                <span></span>
                <span></span>
              </div>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>

        <div className="chat-input-container">
          {pastedImage && (
            <div className="pasted-image-preview">
              <img src={pastedImage} alt="Pasted" />
              <button className="remove-image-btn" onClick={handleRemoveImage}>
                ✕ Remove
              </button>
            </div>
          )}
          <div className="chat-input-group">
            <textarea
              ref={textareaRef}
              placeholder={
                isConfigured
                  ? "Ask Perplexity about Excel features, or paste an image (Ctrl+V)..."
                  : "Please configure your API key first"
              }
              value={inputMessage}
              onChange={(e) => setInputMessage(e.target.value)}
              onKeyPress={handleKeyPress}
              onPaste={handlePaste}
              disabled={!isConfigured || isLoading}
            />
            <button
              onClick={handleSendMessage}
              disabled={!isConfigured || isLoading || (!inputMessage.trim() && !pastedImage)}
            >
              Send
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;
