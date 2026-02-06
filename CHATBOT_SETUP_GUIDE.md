# Complete Chatbot Setup Guide

This guide shows you how to add an AI chatbot to your Creative API frontend using Gemini.

## 🏗️ Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                        USER BROWSER                          │
│                                                              │
│  ┌────────────────────────────────────────────────────┐   │
│  │         Chatbot UI (HTML/JavaScript)                │   │
│  │  - Chat interface                                   │   │
│  │  - Message handling                                 │   │
│  │  - Gemini API integration                          │   │
│  └─────────────────┬──────────────────────────────────┘   │
│                    │                                         │
└────────────────────┼─────────────────────────────────────────┘
                     │
                     │ HTTPS
                     ▼
┌─────────────────────────────────────────────────────────────┐
│                    GEMINI API                                │
│  - Natural language understanding                            │
│  - Function calling (decides which tools to use)             │
│  - Response generation                                       │
└─────────────────┬───────────────────────────────────────────┘
                  │
                  │ Tool Execution Request
                  ▼
┌─────────────────────────────────────────────────────────────┐
│              YOUR FASTAPI SERVER                             │
│              (http://localhost:50100)                        │
│                                                              │
│  Endpoints:                                                  │
│  • GET  /api/daymaster/projects                             │
│  • GET  /api/bi_guidelines/nw-active-presentations          │
│  • POST /api/presentations/create                           │
│  • GET  /api/tasks/{task_id}                               │
│  • GET  /api/analytics/generate/{type}                      │
│  • ... and many more                                        │
└─────────────────┬───────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────────┐
│                 SQL SERVER DATABASES                         │
│  • BI_GUIDELINES (presentations, projects)                   │
│  • DAYMASTER (project master list)                          │
└─────────────────────────────────────────────────────────────┘
```

## 📋 What You Need

### Required

1. **Google Gemini API Key**
   - Free tier available
   - Get it here: https://makersuite.google.com/app/apikey

2. **Your FastAPI Server**
   - Already running at `http://localhost:50100`
   - No authentication required (currently)

3. **Web Browser**
   - Chrome, Firefox, Safari, or Edge
   - JavaScript enabled

### Optional (for advanced features)

- Node.js (for MCP server)
- React/Vue framework (for production frontend)

## 🚀 Quick Start (3 Options)

### Option 1: Simple HTML Chatbot (Easiest)

**Best for:** Quick testing, demos, proof of concept

**Setup time:** 5 minutes

1. Open `frontend-example/chatbot-integration.html` in a browser
2. Enter your Gemini API key
3. Start chatting!

**Pros:**
- No installation required
- Works immediately
- Perfect for testing

**Cons:**
- API key stored in browser localStorage (not secure)
- No file upload support
- Limited customization

---

### Option 2: MCP Server + Claude Desktop (Recommended)

**Best for:** Using Claude Desktop as your chatbot interface

**Setup time:** 10 minutes

1. Install the MCP server:
   ```bash
   cd mcp-server
   npm install
   ```

2. Configure Claude Desktop (edit `claude_desktop_config.json`):
   ```json
   {
     "mcpServers": {
       "creative-api": {
         "command": "node",
         "args": ["C:\\inetpub\\wwwroot\\CreativePythonAPI\\mcp-server\\index.js"],
         "env": {
           "API_BASE_URL": "http://localhost:50100"
         }
       }
     }
   }
   ```

3. Restart Claude Desktop
4. Chat with Claude - it now has access to your API!

**Pros:**
- Professional UI (Claude Desktop)
- Secure (no API keys in browser)
- Full MCP protocol support
- File handling built-in

**Cons:**
- Requires Claude Desktop installation
- Node.js required
- Not embedded in your app

---

### Option 3: Custom React/Vue Frontend (Production)

**Best for:** Production deployment, custom branding

**Setup time:** 1-2 hours

See `frontend-example/README.md` for React implementation details.

**Pros:**
- Full customization
- Secure authentication
- File upload support
- Your own branding

**Cons:**
- More development required
- Need to manage infrastructure

---

## 🔧 Detailed Setup: HTML Chatbot

### Step 1: Enable CORS in FastAPI

Your API needs to accept requests from the browser. Update `app/main.py`:

```python
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="Report Generator API")

# Add this AFTER creating the app
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # For production, specify your domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
```

Restart your FastAPI server after making this change.

### Step 2: Get Gemini API Key

1. Go to https://makersuite.google.com/app/apikey
2. Sign in with Google account
3. Click "Create API Key"
4. Copy the key (starts with `AIza...`)

### Step 3: Open Chatbot

1. Navigate to `frontend-example/chatbot-integration.html`
2. Open in your browser (double-click or drag to browser)
3. Paste your API key in the input field
4. Click "Save Key"

### Step 4: Start Chatting!

Try these example prompts:

```
"Show me all projects"
"Find projects containing 'Coca'"
"How many tasks are in the queue?"
"Check if a presentation exists for project XYZ"
"Generate an NW analytics report"
```

---

## 🎯 What the Chatbot Can Do

### Current Capabilities

The chatbot currently supports these operations:

#### 1. Project Management
- **List projects**: "Show me all projects", "Find projects with X in the name"
- **Search projects**: "Search for Coca-Cola projects"
- **Pagination**: "Show me page 2 of projects"

#### 2. Presentation Management
- **List NW presentations**: "What NW presentations do we have?"
- **Check existence**: "Does project X have a presentation?"
- **Search presentations**: "Find all presentations for region Y"

#### 3. Task Monitoring
- **Check status**: "What's the status of task abc-123?"
- **Queue info**: "How many tasks are running?"
- **Wait times**: "How long until my task runs?"

#### 4. Analytics
- **Generate reports**: "Create an NW analytics report"
- **Date filtering**: "Report for January 2024"
- **Project type**: "BSR analytics report"

### Not Yet Supported (Would require additional work)

- Creating presentations (requires file upload)
- Downloading files (requires binary handling)
- Updating presentation categories
- Deleting presentations
- User authentication

---

## 🔒 Security Considerations

### Current Setup (Development)

**⚠️ Not secure for production:**
- No API authentication
- API key stored in localStorage
- CORS allows all origins
- No rate limiting

### For Production

You MUST add:

1. **API Authentication**
   ```python
   # Add to FastAPI
   from fastapi.security import APIKeyHeader

   API_KEY = APIKeyHeader(name="X-API-Key")

   @app.get("/api/projects")
   async def get_projects(api_key: str = Depends(API_KEY)):
       if api_key != os.getenv("API_KEY"):
           raise HTTPException(401)
       # ... rest of code
   ```

2. **Secure Key Storage**
   - Use environment variables
   - Never commit keys to git
   - Use secret management service

3. **CORS Restrictions**
   ```python
   allow_origins=["https://yourdomain.com"]
   ```

4. **Rate Limiting**
   ```python
   from slowapi import Limiter

   limiter = Limiter(key_func=get_remote_address)

   @app.get("/api/projects")
   @limiter.limit("10/minute")
   async def get_projects():
       # ...
   ```

5. **HTTPS**
   - Use HTTPS in production
   - Configure SSL certificates
   - Redirect HTTP → HTTPS

---

## 🛠️ Customization Guide

### Adding New Capabilities

To add a new action the chatbot can perform:

#### 1. Add Tool Definition (Frontend)

Edit `chatbot-integration.html`, add to `MCP_TOOLS` array:

```javascript
{
    name: 'delete_presentation',
    description: 'Delete a presentation by ID',
    parameters: {
        type: 'object',
        properties: {
            presentation_id: {
                type: 'string',
                description: 'ID of presentation to delete'
            }
        },
        required: ['presentation_id']
    }
}
```

#### 2. Add Execution Logic

Add case to `executeMCPTool()` function:

```javascript
case 'delete_presentation':
    response = await fetch(
        `${API_BASE_URL}/api/presentations/${args.presentation_id}`,
        { method: 'DELETE' }
    );
    break;
```

#### 3. Test

Restart the chatbot and try:
```
"Delete presentation with ID abc-123"
```

That's it! Gemini automatically understands the new capability.

---

## 📊 Monitoring & Debugging

### Check API Calls

Open browser console (F12) to see:
- Gemini API requests/responses
- Function calls being made
- API errors

### Test API Directly

Use the API docs at `http://localhost:50100/docs` to:
- Test endpoints directly
- Verify responses
- Check authentication

### MCP Server Logs

If using MCP server:
```bash
# Run with logging
node index.js 2>&1 | tee mcp-server.log
```

---

## 🚦 Common Issues & Solutions

### Issue: CORS Error

**Symptom:** "Access to fetch blocked by CORS policy"

**Solution:**
1. Add CORS middleware to FastAPI (see Step 1)
2. Restart FastAPI server
3. Clear browser cache
4. Try again

---

### Issue: Gemini API Error 400

**Symptom:** "Invalid API key"

**Solution:**
1. Verify key starts with `AIza...`
2. Check key hasn't expired
3. Verify billing enabled in Google Cloud Console

---

### Issue: Function Not Called

**Symptom:** Chatbot doesn't call the tool

**Solution:**
1. Check tool description is clear
2. Make parameter descriptions specific
3. Try more explicit prompt: "Use the list_projects tool to show me projects"

---

### Issue: FastAPI Connection Refused

**Symptom:** "Failed to fetch" or "ERR_CONNECTION_REFUSED"

**Solution:**
1. Verify FastAPI is running: `http://localhost:50100/health`
2. Check port 50100 is not blocked
3. Verify API_BASE_URL in chatbot matches your setup

---

## 📈 Next Steps

### Immediate Improvements

1. **Add File Upload**
   - Allow users to upload Excel/PPTX through chat
   - "Create presentation from this file"

2. **Add Authentication**
   - User login system
   - API key authentication
   - Role-based access control

3. **Enhance UI**
   - Add file attachment support
   - Show download links for reports
   - Display presentation thumbnails

### Advanced Features

1. **Conversation History**
   - Save chat history to database
   - Allow users to resume conversations
   - Search through past chats

2. **Multi-step Workflows**
   - "Create presentation and generate report"
   - "Check if exists, if not create new"
   - Progress tracking for long operations

3. **Notifications**
   - WebSocket real-time updates
   - Email notifications when tasks complete
   - Slack/Teams integration

4. **Admin Dashboard**
   - Monitor all chatbot conversations
   - Analytics on tool usage
   - User behavior insights

---

## 📚 Additional Resources

### Documentation

- **Your API Docs**: http://localhost:50100/docs
- **Gemini API**: https://ai.google.dev/docs
- **MCP Protocol**: https://modelcontextprotocol.io/

### Example Projects

- **MCP Server Examples**: https://github.com/modelcontextprotocol/servers
- **Gemini Function Calling**: https://ai.google.dev/docs/function_calling

### Support

- **Issues**: Create an issue in your project repo
- **API Status**: Check FastAPI health endpoint
- **Gemini Status**: https://status.cloud.google.com/

---

## 🎉 You're Done!

You now have:
✅ Working chatbot UI
✅ Gemini AI integration
✅ Connection to your API
✅ Natural language interface

Your users can now interact with your API using plain English!

**Example conversation:**

```
User: "Hey, show me projects related to Coca-Cola"
Bot: "I found 3 projects:
     1. Coca-Cola Brand Refresh (NW-2024-001)
     2. Coca-Cola Zero Campaign (NW-2024-005)
     3. Coca-Cola Packaging Study (BSR-2024-003)"

User: "Does the first one have a presentation?"
Bot: "Yes! The Coca-Cola Brand Refresh project has an active
     NW presentation created on 2024-01-15."

User: "Great! What's in the queue right now?"
Bot: "Currently there are 2 tasks running and 3 tasks waiting.
     The system has capacity for 5 concurrent tasks."
```

Simple, conversational, and powerful! 🚀
