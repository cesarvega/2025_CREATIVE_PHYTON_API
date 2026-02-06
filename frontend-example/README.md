# Frontend Chatbot Integration Example

This example demonstrates how to integrate a Gemini-powered chatbot with your Creative API.

## Architecture

```
User Browser
    ↓
Chatbot UI (HTML/JS)
    ↓
Gemini API (with Function Calling)
    ↓
Your FastAPI (http://localhost:50100)
    ↓
SQL Server Databases
```

## How It Works

1. **User enters message** → Sent to Gemini API
2. **Gemini decides** → Should it call a function or respond directly?
3. **If function needed** → Frontend calls your FastAPI endpoint
4. **API returns data** → Sent back to Gemini
5. **Gemini formats response** → Shows friendly message to user

## Setup

### 1. Get Gemini API Key

1. Go to [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Create a new API key
3. Copy the key (starts with `AIza...`)

### 2. Configure CORS in Your API

Add CORS headers to allow frontend access. Update your [app/main.py](../app/main.py):

```python
from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, specify your frontend domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
```

### 3. Run the Chatbot

1. Open `chatbot-integration.html` in a web browser
2. Enter your Gemini API key
3. Start chatting!

## Example Conversations

### List Projects
```
User: Show me all projects containing "Coca"
Bot: I found 3 projects matching "Coca":
     1. Coca-Cola Brand Refresh (NW-2024-001)
     2. Coca-Cola Zero Campaign (NW-2024-005)
     3. Coca-Cola Packaging (BSR-2024-003)
```

### Check Task Status
```
User: What's the status of task abc-123-def?
Bot: Your presentation is currently processing:
     - Status: in_progress
     - Progress: 65%
     - Message: Generating slides 13/20
```

### Get Analytics
```
User: Generate an analytics report for NW projects from January 2024
Bot: I've generated the analytics report! It includes:
     - Total projects: 45
     - By region breakdown
     - Monthly trends
     The report is ready for download.
```

## Available Commands

The chatbot can understand natural language requests for:

- **Listing projects**: "Show me projects", "Find projects with X"
- **Checking presentations**: "Does project X have a presentation?"
- **Task monitoring**: "Check task status", "What's happening with task Y?"
- **Queue info**: "How many tasks are running?"
- **Analytics**: "Generate report for NW projects"

## Customization

### Change API URL

Update the `API_BASE_URL` in `chatbot-integration.html`:

```javascript
const API_BASE_URL = 'http://your-server:50100';
```

### Add More Tools

To add new tools:

1. Add to `MCP_TOOLS` array in the HTML file
2. Add corresponding case in `executeMCPTool()` function
3. The chatbot will automatically understand the new capability!

### Styling

The chatbot uses CSS variables for easy theming. Modify the gradient colors:

```css
background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
```

## Production Deployment

For production use:

1. **Secure API key storage**: Store in environment variables, not localStorage
2. **Add authentication**: Implement user login and session management
3. **Rate limiting**: Add rate limits to prevent API abuse
4. **CORS**: Restrict CORS to your specific domain
5. **HTTPS**: Use HTTPS for all API calls
6. **Error handling**: Add comprehensive error handling and logging

## Alternative: React Implementation

For a more robust solution, consider using React with these libraries:

```bash
npm install @google/generative-ai
npm install axios
```

Example React component:

```jsx
import { GoogleGenerativeAI } from "@google/generative-ai";

function ChatBot() {
  const genAI = new GoogleGenerativeAI(process.env.REACT_APP_GEMINI_KEY);
  const model = genAI.getGenerativeModel({ model: "gemini-2.0-flash-exp" });

  const chat = model.startChat({
    tools: [{ functionDeclarations: MCP_TOOLS }],
    systemInstruction: "You are a helpful assistant..."
  });

  // Handle messages and function calls
  // ...
}
```

## Troubleshooting

### CORS Errors
- Ensure CORS middleware is configured in your FastAPI app
- Check browser console for specific CORS error messages

### API Key Issues
- Verify your Gemini API key is valid
- Check for billing/quota issues in Google Cloud Console

### Connection Refused
- Ensure your FastAPI server is running on port 50100
- Check firewall settings

### Function Calls Not Working
- Verify tool names match exactly between frontend and API
- Check browser console for function call details

## Next Steps

1. Add file upload capability for Excel/PPTX files
2. Implement user authentication
3. Add conversation history persistence
4. Create admin dashboard for monitoring
5. Add support for creating presentations through chat
