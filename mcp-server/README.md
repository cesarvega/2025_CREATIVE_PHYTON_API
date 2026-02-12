# Creative API MCP Server v3.0

Comprehensive Model Context Protocol (MCP) server that wraps the Creative Python API for AI chatbot integration. Provides **50+ tools** for complete API access.

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

## Available Tools (50+)

### System & Health (2 tools)
| Tool | Description |
|------|-------------|
| `health_check` | Check API health status |
| `get_system_info` | Get server version and capabilities |

### DayMaster Projects (1 tool)
| Tool | Description |
|------|-------------|
| `list_daymaster_projects` | Search projects with filters (search, type, status, dates, pagination) |

### NW Presentations (10 tools)
| Tool | Description |
|------|-------------|
| `list_nw_presentations` | List active NW presentations with optional search |
| `get_nw_presentation_details` | Get full details of an NW presentation |
| `check_nw_presentation_exists` | Check if NW presentation exists |
| `create_nw_presentation` | Create new NW presentation from Excel file |
| `update_nw_presentation` | Update existing NW presentation |
| `delete_nw_presentation` | Delete NW presentation |
| `download_nw_presentation` | Download NW presentation file |
| `create_nw_feedback_template` | Generate NW feedback template |
| `search_nw_group_names` | Search group names within presentation |
| `get_nw_presentation_groups` | Get all groups with details |

### BSR Presentations (9 tools)
| Tool | Description |
|------|-------------|
| `list_bsr_presentations` | List active BSR presentations |
| `get_bsr_presentation_details` | Get full details of BSR presentation |
| `check_bsr_presentation_exists` | Check if BSR presentation exists |
| `create_bsr_presentation` | Create new BSR presentation |
| `update_bsr_presentation` | Update existing BSR presentation |
| `delete_bsr_presentation` | Delete BSR presentation |
| `download_bsr_presentation` | Download BSR presentation file |
| `create_bsr_feedback_template` | Generate BSR feedback template |
| `get_bsr_concepts` | Get project concepts and names |

### NSR Configuration (6 tools)
| Tool | Description |
|------|-------------|
| `get_nsr_config` | Get NSR configuration for presentation |
| `update_nsr_config` | Update NSR configuration |
| `create_nsr_report` | Generate NSR report |
| `download_nsr_report` | Download generated NSR report |
| `get_nsr_status` | Check NSR generation status |
| `list_available_nsr_types` | List available NSR report types |

### DW Presentations (1 tool)
| Tool | Description |
|------|-------------|
| `create_dw_presentation` | Create Design Workshop presentation |

### Backup & File Operations (3 tools)
| Tool | Description |
|------|-------------|
| `create_backup` | Create backup of presentation files |
| `list_backups` | List available backups |
| `restore_backup` | Restore from backup |

### Background Templates (6 tools)
| Tool | Description |
|------|-------------|
| `list_background_templates` | List available background templates |
| `get_template_details` | Get template configuration |
| `upload_background_template` | Upload new background image |
| `delete_background_template` | Delete background template |
| `preview_template` | Get template preview image |
| `apply_template_to_presentation` | Apply template to presentation |

### Analytics & Statistics (5 tools)
| Tool | Description |
|------|-------------|
| `get_nw_analytics` | Get NW projects analytics |
| `get_nw_region_analytics` | Get NW analytics by region |
| `get_bsr_analytics` | Get BSR projects analytics |
| `get_bsr_region_analytics` | Get BSR analytics by region |
| `generate_analytics_report` | Generate complete analytics Excel |

### Task Management (7 tools)
| Tool | Description |
|------|-------------|
| `check_task_status` | Check background task status |
| `list_active_tasks` | List all active tasks |
| `cancel_task` | Cancel running task |
| `get_task_result` | Get task result/download |
| `get_queue_status` | Get task queue status |
| `retry_failed_task` | Retry a failed task |
| `cleanup_old_tasks` | Clean up completed tasks |

### File Processing (2 tools)
| Tool | Description |
|------|-------------|
| `process_excel_file` | Process/validate Excel file |
| `convert_pptx_to_images` | Convert PPTX to images |

### Search & Discovery (3 tools)
| Tool | Description |
|------|-------------|
| `search_all_presentations` | Search across all presentation types |
| `get_recent_activity` | Get recent system activity |
| `get_presentation_statistics` | Get aggregated statistics |

### Testing (1 tool)
| Tool | Description |
|------|-------------|
| `test_database_connection` | Test database connectivity |

## Integration with Claude Desktop

Add to your Claude Desktop config:

**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
**Mac:** `~/Library/Application Support/Claude/claude_desktop_config.json`

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

## Usage Examples

### Natural Language Queries

The AI chatbot can now handle complex queries like:

- "List all NW presentations from the last month"
- "Create a new BSR presentation for project ABC123"
- "Show me analytics for all NW projects this quarter"
- "What's the status of task abc-123?"
- "Generate a feedback template for presentation XYZ"
- "Search for presentations containing 'Brand'"
- "Download the NSR report for presentation P001"

### Frontend Integration with Gemini

```javascript
import { GoogleGenerativeAI } from "@google/generative-ai";

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// Define function declarations from MCP tools
const tools = [{
  functionDeclarations: [
    {
      name: 'list_nw_presentations',
      description: 'List active NW presentations with optional search',
      parameters: {
        type: 'object',
        properties: {
          search: { type: 'string', description: 'Search term' },
          page: { type: 'number', description: 'Page number' },
          page_size: { type: 'number', description: 'Items per page' }
        }
      }
    },
    {
      name: 'get_nw_analytics',
      description: 'Get analytics for all NW projects',
      parameters: {
        type: 'object',
        properties: {
          start_date: { type: 'string', description: 'Start date YYYY-MM-DD' },
          end_date: { type: 'string', description: 'End date YYYY-MM-DD' }
        }
      }
    }
    // ... more tools
  ]
}];

// Chat with function calling
const model = genAI.getGenerativeModel({
  model: "gemini-pro",
  tools
});

const chat = model.startChat();
const result = await chat.sendMessage("Show me all NW presentations");

// Handle function calls
if (result.response.functionCalls) {
  for (const call of result.response.functionCalls) {
    const mcpResult = await callMCPServer(call.name, call.args);
    // Send result back to continue conversation
  }
}
```

## Testing

### Using MCP Inspector
```bash
# Install MCP inspector
npm install -g @modelcontextprotocol/inspector

# Run inspector
mcp-inspector node index.js
```

### Direct API Testing
```bash
# Test health endpoint
curl http://localhost:50100/health

# Test NW presentations list
curl http://localhost:50100/api/bi-guidelines/nw/presentations
```

## Architecture

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   Frontend      │     │   MCP Server    │     │  Creative API   │
│   (Gemini AI)   │────▶│   (Node.js)     │────▶│   (FastAPI)     │
└─────────────────┘     └─────────────────┘     └─────────────────┘
                              │                        │
                              ▼                        ▼
                        ┌─────────────────┐     ┌─────────────────┐
                        │  Tool Registry  │     │   SQL Server    │
                        │   (50+ tools)   │     │   Databases     │
                        └─────────────────┘     └─────────────────┘
```

## Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `API_BASE_URL` | Creative API base URL | `http://localhost:50100` |

## Changelog

### v3.0.0
- Added 50+ tools covering all API endpoints
- File upload support for presentation creation
- NSR configuration and report generation
- Background template management
- Complete task management system
- Analytics and statistics tools
- Search across all presentation types

### v2.0.0
- Added advanced filtering (dates, project type)
- Expanded to 27 tools
- Region analytics support

### v1.0.0
- Initial release with 10 basic tools
