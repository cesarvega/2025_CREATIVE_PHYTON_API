# Creative API MCP Server

Model Context Protocol (MCP) server that wraps the Creative Python API for AI chatbot integration.

## Setup

1. Install dependencies:
```bash
npm install
```

2. Configure environment:
```bash
cp .env.example .env
# Edit .env with your API URL
```

3. Run the server:
```bash
npm start
```

## Available Tools

The MCP server exposes these tools to AI assistants:

### Project Management
- `list_projects` - List projects from DayMaster with search/pagination
- `list_nw_presentations` - List active NW presentations

### Presentation Creation
- `create_nw_presentation` - Create Name Evaluation presentations
- `create_bsr_presentation` - Create Brand Strategy Reports
- `create_feedback_template` - Generate feedback templates

### Task Management
- `check_task_status` - Monitor background task progress
- `get_queue_status` - Check system queue status

### Analytics
- `get_analytics_report` - Generate analytics reports (NW/BSR)

### Utilities
- `check_nw_presentation_exists` - Check if presentation exists
- `download_presentation` - Download presentation files

## Integration with Claude Desktop

Add to your Claude Desktop config (`~/Library/Application Support/Claude/claude_desktop_config.json` on Mac):

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

## Usage with Gemini Frontend

Your frontend can use the Gemini API with function calling to interact with this MCP server:

```javascript
// Example: Frontend calls Gemini API
const response = await gemini.generateContent({
  contents: [{
    role: 'user',
    parts: [{ text: 'List all projects containing "Brand"' }]
  }],
  tools: [{
    functionDeclarations: [{
      name: 'list_projects',
      description: 'List projects from DayMaster database',
      parameters: {
        type: 'object',
        properties: {
          search: { type: 'string' }
        }
      }
    }]
  }]
});

// If Gemini requests a function call, execute it via MCP
if (response.functionCall) {
  const result = await callMCPTool(response.functionCall.name, response.functionCall.args);
  // Send result back to Gemini
}
```

## Testing

Test the MCP server directly:

```bash
# Install MCP inspector
npm install -g @modelcontextprotocol/inspector

# Run inspector
mcp-inspector node index.js
```
