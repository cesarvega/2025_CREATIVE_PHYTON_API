import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import axios from "axios";

// Configure your API base URL
const API_BASE_URL = process.env.API_BASE_URL || "http://localhost:50100";

class CreativeAPIServer {
  constructor() {
    this.server = new Server(
      {
        name: "creative-api-mcp-server",
        version: "2.0.0",
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.axiosInstance = axios.create({
      baseURL: API_BASE_URL,
      timeout: 300000, // 5 minutes for long-running tasks
    });

    this.setupToolHandlers();

    this.server.onerror = (error) => console.error("[MCP Error]", error);
    process.on("SIGINT", async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  setupToolHandlers() {
    // List all available tools
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        // ==================== PROJECT MANAGEMENT ====================
        {
          name: "list_daymaster_projects",
          description: "List projects from DayMaster database with pagination and search. Use this to find projects by name, code, or browse all available projects. Returns project information including names, codes, and details.",
          inputSchema: {
            type: "object",
            properties: {
              page: { type: "number", description: "Page number (default: 1)" },
              limit: { type: "number", description: "Items per page (default: 10, max: 100)" },
              search: { type: "string", description: "Search term to filter projects by name or code" }
            },
          },
        },
        {
          name: "get_project_info",
          description: "Get detailed information about a specific project by its ID. Automatically detects if it's NW, BSR, or NSR type and returns all project details including presentation type, status, upload info, and categories.",
          inputSchema: {
            type: "object",
            properties: {
              project_id: { type: "number", description: "The project/presentation ID to get info for" }
            },
            required: ["project_id"],
          },
        },

        // ==================== NW PRESENTATIONS ====================
        {
          name: "list_nw_presentations",
          description: "List active Name Evaluation (NW) presentations with pagination and search. Returns presentations sorted by last update date (newest first). Includes project name, display name, status, and dates.",
          inputSchema: {
            type: "object",
            properties: {
              page: { type: "number", description: "Page number (default: 1)" },
              limit: { type: "number", description: "Items per page (default: 50, max: 500)" },
              search: { type: "string", description: "Search term to filter by project or display name" }
            },
          },
        },
        {
          name: "check_nw_presentation_exists",
          description: "Check if a Name Evaluation (NW) presentation already exists for a specific project. Use this before creating a new presentation to avoid duplicates.",
          inputSchema: {
            type: "object",
            properties: {
              display_name: { type: "string", description: "Display name of the project to check" }
            },
            required: ["display_name"],
          },
        },
        {
          name: "update_nw_project",
          description: "Update NW project details including display name, status (OPEN/CLOSED), and BSR display name association.",
          inputSchema: {
            type: "object",
            properties: {
              presentation_id: { type: "number", description: "The presentation ID to update" },
              display_name: { type: "string", description: "New display name" },
              presentation_status: { type: "string", enum: ["OPEN", "CLOSED"], description: "Project status" },
              bsr_display_name: { type: "string", description: "Associated BSR display name (optional)" }
            },
            required: ["presentation_id", "display_name", "presentation_status"],
          },
        },

        // ==================== BSR PRESENTATIONS ====================
        {
          name: "list_bsr_presentations",
          description: "List active Brand Strategy Report (BSR) presentations with pagination and search. Returns BSR presentations sorted by last update date.",
          inputSchema: {
            type: "object",
            properties: {
              page: { type: "number", description: "Page number (default: 1)" },
              limit: { type: "number", description: "Items per page (default: 50, max: 500)" },
              search: { type: "string", description: "Search term to filter presentations" }
            },
          },
        },
        {
          name: "get_bsr_project_info",
          description: "Get detailed information about a specific BSR project including presentation details, slide configuration, and project categories.",
          inputSchema: {
            type: "object",
            properties: {
              project_id: { type: "number", description: "The BSR presentation ID" }
            },
            required: ["project_id"],
          },
        },
        {
          name: "list_bsr_display_names",
          description: "Get list of BSR display names for dropdown selection. Returns distinct BSR display names sorted alphabetically.",
          inputSchema: {
            type: "object",
            properties: {
              page: { type: "number", description: "Page number (default: 1)" },
              limit: { type: "number", description: "Items per page (default: 50)" },
              search: { type: "string", description: "Search term to filter display names" }
            },
          },
        },
        {
          name: "update_bsr_project",
          description: "Update BSR project details including display name and status (OPEN/CLOSED).",
          inputSchema: {
            type: "object",
            properties: {
              presentation_id: { type: "number", description: "The BSR presentation ID to update" },
              display_name: { type: "string", description: "New display name" },
              status: { type: "string", enum: ["OPEN", "CLOSED"], description: "Project status" }
            },
            required: ["presentation_id", "display_name", "status"],
          },
        },

        // ==================== NSR CONFIGURATION ====================
        {
          name: "list_nsr_projects",
          description: "Get list of active NSR (Name Safety Report) projects. Returns all active NSR projects for dropdown selection.",
          inputSchema: {
            type: "object",
            properties: {},
          },
        },
        {
          name: "get_nsr_config",
          description: "Get NSR validation rule configuration for a specific project. Returns all rules (101-120) and their ON/OFF states.",
          inputSchema: {
            type: "object",
            properties: {
              project_name: { type: "string", description: "The NSR project name" }
            },
            required: ["project_name"],
          },
        },
        {
          name: "update_nsr_rule",
          description: "Update or insert an NSR validation rule for a project. Rules control validation checks like USAN compliance, prefix rules, etc.",
          inputSchema: {
            type: "object",
            properties: {
              project_name: { type: "string", description: "The NSR project name" },
              rule_id: {
                type: "number",
                description: "Rule ID (101=Contains Y, 102=Contains H, 103=Contains W, 104=Contains J, 105=Contains K, 106=Check USAN Violation, 108=Check USAN Nomenclature, 109-115=Prefix rules, 116-117=USAN checks)"
              },
              is_on: { type: "number", enum: [0, 1], description: "1=ON (active), 0=OFF (inactive)" }
            },
            required: ["project_name", "rule_id", "is_on"],
          },
        },
        {
          name: "initialize_nsr_rules",
          description: "Initialize all NSR validation rules for a new project. Creates all 19 rules (101-120 except 107) in ON state.",
          inputSchema: {
            type: "object",
            properties: {
              project_name: { type: "string", description: "The NSR project name to initialize" }
            },
            required: ["project_name"],
          },
        },

        // ==================== ANALYTICS & STATISTICS ====================
        {
          name: "get_nw_analytics",
          description: "Get analytics data for ALL NW projects. Returns vote percentages (retained, positive, neutral, negative), newly created names count, and active status for each project.",
          inputSchema: {
            type: "object",
            properties: {},
          },
        },
        {
          name: "get_nw_region_analytics",
          description: "Get aggregated NW analytics by region/active lead. Returns project counts and average statistics per region.",
          inputSchema: {
            type: "object",
            properties: {},
          },
        },
        {
          name: "get_bsr_analytics",
          description: "Get analytics data for ALL BSR projects. Returns PC access count, mobile access count, and total access for each project.",
          inputSchema: {
            type: "object",
            properties: {},
          },
        },
        {
          name: "get_bsr_region_analytics",
          description: "Get aggregated BSR analytics by region/active lead. Returns total access counts per region.",
          inputSchema: {
            type: "object",
            properties: {},
          },
        },
        {
          name: "generate_analytics_report",
          description: "Generate and download complete analytics Excel report. Creates a formatted Excel file with Project-Specific and Region-Specific sheets.",
          inputSchema: {
            type: "object",
            properties: {
              project_type: { type: "string", enum: ["NW", "BSR"], description: "Type of analytics report" },
              start_date: { type: "string", description: "Filter from date (YYYY-MM-DD format, optional)" },
              end_date: { type: "string", description: "Filter to date (YYYY-MM-DD format, optional)" }
            },
            required: ["project_type"],
          },
        },

        // ==================== TASK MANAGEMENT ====================
        {
          name: "check_task_status",
          description: "Check the status of a background task (presentation creation, backup generation, etc.). Returns status (pending/processing/completed/failed), progress percentage, and result or error.",
          inputSchema: {
            type: "object",
            properties: {
              task_id: { type: "string", description: "Task ID returned from creation operations" }
            },
            required: ["task_id"],
          },
        },
        {
          name: "list_tasks",
          description: "List all background tasks with optional filtering by type and status. Tasks are sorted by creation time (newest first).",
          inputSchema: {
            type: "object",
            properties: {
              task_type: { type: "string", description: "Filter by type: create_presentation, generate_backup, create_bsr, etc." },
              status: { type: "string", enum: ["pending", "processing", "completed", "failed"], description: "Filter by status" },
              limit: { type: "number", description: "Max tasks to return (default: 100)" }
            },
          },
        },
        {
          name: "cancel_task",
          description: "Cancel a queued task or delete a completed/failed task from the system.",
          inputSchema: {
            type: "object",
            properties: {
              task_id: { type: "string", description: "Task ID to cancel or delete" }
            },
            required: ["task_id"],
          },
        },
        {
          name: "get_queue_status",
          description: "Get current task queue status including active tasks, queued tasks, available slots, and system capacity. Use this to check if the system can accept new tasks.",
          inputSchema: {
            type: "object",
            properties: {},
          },
        },

        // ==================== TEMPLATES & THEMES ====================
        {
          name: "list_template_groups",
          description: "Get list of available background templates grouped by category. Returns template IDs, names, categories, and file paths.",
          inputSchema: {
            type: "object",
            properties: {},
          },
        },
        {
          name: "list_background_templates",
          description: "Get list of background templates with optional filtering by group. Returns templates with preview URLs and thumbnail URLs.",
          inputSchema: {
            type: "object",
            properties: {
              template_group: { type: "string", description: "Filter by template group/category" },
              page: { type: "number", description: "Page number (default: 1)" },
              limit: { type: "number", description: "Items per page (default: 50, max: 100)" }
            },
          },
        },
        {
          name: "get_background_template",
          description: "Get detailed information about a specific background template by ID.",
          inputSchema: {
            type: "object",
            properties: {
              template_id: { type: "number", description: "The template ID to retrieve" }
            },
            required: ["template_id"],
          },
        },

        // ==================== REPORT GENERATION ====================
        {
          name: "create_feedback_template",
          description: "Generate NW feedback template Word document for a presentation. Creates a formatted document with name evaluation feedback sections.",
          inputSchema: {
            type: "object",
            properties: {
              display_name: { type: "string", description: "Display name of the NW project" }
            },
            required: ["display_name"],
          },
        },
        {
          name: "reload_project_sounds",
          description: "Reload MP3 file paths for a NW project. Use this when audio files have been updated or moved and need to be reindexed in the database.",
          inputSchema: {
            type: "object",
            properties: {
              display_name: { type: "string", description: "Display name of the project" }
            },
            required: ["display_name"],
          },
        },

        // ==================== SYSTEM HEALTH ====================
        {
          name: "check_api_health",
          description: "Check if the API server is running and healthy. Returns API name, version, and documentation URLs.",
          inputSchema: {
            type: "object",
            properties: {},
          },
        },

        // ==================== STATISTICS QUERIES ====================
        {
          name: "count_presentations_by_type",
          description: "Get counts of presentations by project type (NW, BSR, NSR). Useful for dashboard statistics.",
          inputSchema: {
            type: "object",
            properties: {
              project_type: { type: "string", enum: ["NW", "BSR", "all"], description: "Type to count or 'all' for all types" }
            },
            required: ["project_type"],
          },
        },
        {
          name: "get_recent_presentations",
          description: "Get recently created or updated presentations across all types. Useful for activity feed or dashboard.",
          inputSchema: {
            type: "object",
            properties: {
              project_type: { type: "string", enum: ["NW", "BSR", "all"], description: "Filter by type or 'all'" },
              days: { type: "number", description: "Number of days to look back (default: 7)" },
              limit: { type: "number", description: "Max results (default: 10)" }
            },
          },
        },
        {
          name: "search_presentations",
          description: "Search across all presentation types (NW, BSR) by project name, display name, or uploaded by. Returns combined results from all tables.",
          inputSchema: {
            type: "object",
            properties: {
              query: { type: "string", description: "Search query (searches project name, display name, uploaded by)" },
              project_type: { type: "string", enum: ["NW", "BSR", "all"], description: "Filter by type or 'all' (default)" },
              status: { type: "string", enum: ["OPEN", "CLOSED", "all"], description: "Filter by status or 'all' (default)" },
              limit: { type: "number", description: "Max results (default: 20)" }
            },
            required: ["query"],
          },
        },
      ],
    }));

    // Handle tool calls
    this.server.setRequestHandler(CallToolRequestSchema, async (request) =>
      this.handleToolCall(request.params.name, request.params.arguments || {})
    );
  }

  async handleToolCall(toolName, args) {
    try {
      switch (toolName) {
        // Project Management
        case "list_daymaster_projects":
          return await this.listDaymasterProjects(args);
        case "get_project_info":
          return await this.getProjectInfo(args);

        // NW Presentations
        case "list_nw_presentations":
          return await this.listNWPresentations(args);
        case "check_nw_presentation_exists":
          return await this.checkNWPresentationExists(args);
        case "update_nw_project":
          return await this.updateNWProject(args);

        // BSR Presentations
        case "list_bsr_presentations":
          return await this.listBSRPresentations(args);
        case "get_bsr_project_info":
          return await this.getBSRProjectInfo(args);
        case "list_bsr_display_names":
          return await this.listBSRDisplayNames(args);
        case "update_bsr_project":
          return await this.updateBSRProject(args);

        // NSR Configuration
        case "list_nsr_projects":
          return await this.listNSRProjects();
        case "get_nsr_config":
          return await this.getNSRConfig(args);
        case "update_nsr_rule":
          return await this.updateNSRRule(args);
        case "initialize_nsr_rules":
          return await this.initializeNSRRules(args);

        // Analytics
        case "get_nw_analytics":
          return await this.getNWAnalytics();
        case "get_nw_region_analytics":
          return await this.getNWRegionAnalytics();
        case "get_bsr_analytics":
          return await this.getBSRAnalytics();
        case "get_bsr_region_analytics":
          return await this.getBSRRegionAnalytics();
        case "generate_analytics_report":
          return await this.generateAnalyticsReport(args);

        // Task Management
        case "check_task_status":
          return await this.checkTaskStatus(args);
        case "list_tasks":
          return await this.listTasks(args);
        case "cancel_task":
          return await this.cancelTask(args);
        case "get_queue_status":
          return await this.getQueueStatus();

        // Templates
        case "list_template_groups":
          return await this.listTemplateGroups();
        case "list_background_templates":
          return await this.listBackgroundTemplates(args);
        case "get_background_template":
          return await this.getBackgroundTemplate(args);

        // Reports
        case "create_feedback_template":
          return await this.createFeedbackTemplate(args);
        case "reload_project_sounds":
          return await this.reloadProjectSounds(args);

        // System
        case "check_api_health":
          return await this.checkAPIHealth();

        // Statistics
        case "count_presentations_by_type":
          return await this.countPresentationsByType(args);
        case "get_recent_presentations":
          return await this.getRecentPresentations(args);
        case "search_presentations":
          return await this.searchPresentations(args);

        default:
          throw new Error(`Unknown tool: ${toolName}`);
      }
    } catch (error) {
      const errorMessage = error.response?.data?.detail || error.message;
      return {
        content: [{ type: "text", text: `Error: ${errorMessage}` }],
        isError: true,
      };
    }
  }

  // ==================== PROJECT MANAGEMENT ====================

  async listDaymasterProjects(args) {
    const { page = 1, limit = 10, search = "" } = args;
    const response = await this.axiosInstance.get("/api/daymaster/projects", {
      params: { page, limit, search },
    });

    const data = response.data;
    let text = `Found ${data.total} projects (showing page ${data.page} of ${Math.ceil(data.total/data.limit)})\n\n`;

    if (data.projects && data.projects.length > 0) {
      data.projects.forEach((p, i) => {
        text += `${i+1}. ${p.ProjectName || p.project_name || 'N/A'}\n`;
        if (p.ProjectCode || p.project_code) text += `   Code: ${p.ProjectCode || p.project_code}\n`;
      });
    } else {
      text += "No projects found.";
    }

    return { content: [{ type: "text", text }] };
  }

  async getProjectInfo(args) {
    const response = await this.axiosInstance.get(`/api/bi_guidelines/project-info/${args.project_id}`);
    const data = response.data;

    let text = `Project Information (ID: ${args.project_id})\n`;
    text += `${'='.repeat(40)}\n`;
    text += `Type: ${data.project_type || 'Unknown'}\n`;
    text += `Project: ${data.project || data.Project || 'N/A'}\n`;
    text += `Display Name: ${data.displayname || data.DisplayName || 'N/A'}\n`;
    text += `Status: ${data.presentationstatus || data.PresentationStatus || 'N/A'}\n`;
    text += `Presentation Type: ${data.presentationtype || data.PresentationType || 'N/A'}\n`;
    text += `Uploaded By: ${data.uploadedby || data.UploadedBy || 'N/A'}\n`;
    text += `Page Number: ${data.page_number || 'N/A'}\n`;

    if (data.categories && data.categories.length > 0) {
      text += `\nCategories:\n`;
      data.categories.forEach(c => {
        text += `  - ${c.category}: ${c.elements}\n`;
      });
    }

    return { content: [{ type: "text", text }] };
  }

  // ==================== NW PRESENTATIONS ====================

  async listNWPresentations(args) {
    const { page = 1, limit = 50, search = "" } = args;
    const response = await this.axiosInstance.get("/api/bi_guidelines/nw-active-presentations", {
      params: { page, limit, search },
    });

    const data = response.data;
    let text = `NW Presentations: ${data.total} total (page ${data.page})\n\n`;

    if (data.presentations && data.presentations.length > 0) {
      data.presentations.forEach((p, i) => {
        text += `${i+1}. ${p.display_name || p.DisplayName}\n`;
        text += `   Project: ${p.project || p.Project}\n`;
        text += `   ID: ${p.presentation_id || p.PresentationId}\n`;
        text += `   Status: ${p.status || p.PresentationStatus || 'OPEN'}\n\n`;
      });
    } else {
      text += "No NW presentations found.";
    }

    return { content: [{ type: "text", text }] };
  }

  async checkNWPresentationExists(args) {
    const response = await this.axiosInstance.get("/api/presentations/exists", {
      params: { display_name: args.display_name },
    });

    const exists = response.data.exists;
    const text = exists
      ? `Yes, a presentation already exists for "${args.display_name}" (ID: ${response.data.presentation_id})`
      : `No presentation found for "${args.display_name}" - you can create one.`;

    return { content: [{ type: "text", text }] };
  }

  async updateNWProject(args) {
    const response = await this.axiosInstance.put("/api/bi_guidelines/update-project-details", {
      presentation_id: args.presentation_id,
      display_name: args.display_name,
      presentation_status: args.presentation_status,
      bsr_display_name: args.bsr_display_name || null,
    });

    return { content: [{ type: "text", text: `NW project ${args.presentation_id} updated successfully.\n\nNew values:\n- Display Name: ${args.display_name}\n- Status: ${args.presentation_status}` }] };
  }

  // ==================== BSR PRESENTATIONS ====================

  async listBSRPresentations(args) {
    const { page = 1, limit = 50, search = "" } = args;
    const response = await this.axiosInstance.get("/api/bi_guidelines/bsr-active-presentations", {
      params: { page, limit, search },
    });

    const data = response.data;
    let text = `BSR Presentations: ${data.total} total (page ${data.page})\n\n`;

    if (data.presentations && data.presentations.length > 0) {
      data.presentations.forEach((p, i) => {
        text += `${i+1}. ${p.display_name || p.DisplayName}\n`;
        text += `   Project: ${p.project || p.Project}\n`;
        text += `   ID: ${p.presentation_id || p.PresentationId}\n\n`;
      });
    } else {
      text += "No BSR presentations found.";
    }

    return { content: [{ type: "text", text }] };
  }

  async getBSRProjectInfo(args) {
    const response = await this.axiosInstance.get(`/api/bi_guidelines/bsr-project-info/${args.project_id}`);
    const data = response.data;

    let text = `BSR Project Information (ID: ${args.project_id})\n`;
    text += `${'='.repeat(40)}\n`;
    text += `Project: ${data.project || 'N/A'}\n`;
    text += `Display Name: ${data.displayname || 'N/A'}\n`;
    text += `Status: ${data.presentationstatus || 'N/A'}\n`;
    text += `Type: ${data.presentationtype || 'N/A'}\n`;
    text += `Uploaded By: ${data.uploadedby || 'N/A'}\n`;
    text += `Slide Number: ${data.slidenumber || data.page_number || 'N/A'}\n`;
    text += `Wide Screen: ${data.iswideppt ? 'Yes' : 'No'}\n`;

    if (data.categories && data.categories.length > 0) {
      text += `\nCategories:\n`;
      data.categories.forEach(c => {
        text += `  - ${c.category}: ${c.elements}\n`;
      });
    }

    return { content: [{ type: "text", text }] };
  }

  async listBSRDisplayNames(args) {
    const { page = 1, limit = 50, search = "" } = args;
    const response = await this.axiosInstance.get("/api/bi_guidelines/display-names", {
      params: { page, limit, search },
    });

    const data = response.data;
    let text = `BSR Display Names: ${data.total} total\n\n`;

    if (data.display_names && data.display_names.length > 0) {
      data.display_names.forEach((name, i) => {
        text += `${i+1}. ${name}\n`;
      });
    }

    return { content: [{ type: "text", text }] };
  }

  async updateBSRProject(args) {
    const response = await this.axiosInstance.put("/api/bi_guidelines/update-bsr-project", {
      presentation_id: args.presentation_id,
      display_name: args.display_name,
      status: args.status,
    });

    return { content: [{ type: "text", text: `BSR project ${args.presentation_id} updated successfully.\n\nNew values:\n- Display Name: ${args.display_name}\n- Status: ${args.status}` }] };
  }

  // ==================== NSR CONFIGURATION ====================

  async listNSRProjects() {
    const response = await this.axiosInstance.get("/api/bi_guidelines/nsr-projects");
    const data = response.data;

    let text = `Active NSR Projects:\n\n`;
    if (data.data && data.data.length > 0) {
      data.data.forEach((p, i) => {
        text += `${i+1}. ${p}\n`;
      });
    } else {
      text += "No active NSR projects found.";
    }

    return { content: [{ type: "text", text }] };
  }

  async getNSRConfig(args) {
    const response = await this.axiosInstance.get(`/api/bi_guidelines/nsr-config/${encodeURIComponent(args.project_name)}`);
    const data = response.data;

    let text = `NSR Configuration for: ${data.project_name}\n`;
    text += `${'='.repeat(40)}\n`;

    const ruleDescriptions = {
      101: "Contains Y", 102: "Contains H", 103: "Contains W", 104: "Contains J", 105: "Contains K",
      106: "Check USAN Violation", 108: "Check USAN Nomenclature",
      109: "Prefix AR", 110: "Prefix DEX", 111: "Prefix ES", 112: "Prefix STR",
      113: "Prefix LEV", 114: "Prefix X", 115: "Prefix RAC",
      116: "Check USAN Search", 117: "Check USAN MedNet",
      118: "Double Letter", 119: "Double Letter 2", 120: "Prefix RAC2"
    };

    if (data.rules && data.rules.length > 0) {
      data.rules.forEach(r => {
        const desc = ruleDescriptions[r.rule_id] || `Rule ${r.rule_id}`;
        const status = r.is_on ? "ON ✓" : "OFF ✗";
        text += `${r.rule_id}: ${desc} - ${status}\n`;
      });
    } else {
      text += "No rules configured (all OFF by default)";
    }

    return { content: [{ type: "text", text }] };
  }

  async updateNSRRule(args) {
    const response = await this.axiosInstance.put("/api/bi_guidelines/nsr-config", {
      project_name: args.project_name,
      rule_id: args.rule_id,
      is_on: args.is_on,
    });

    const status = args.is_on ? "ON" : "OFF";
    return { content: [{ type: "text", text: `NSR rule ${args.rule_id} for "${args.project_name}" set to ${status}` }] };
  }

  async initializeNSRRules(args) {
    const response = await this.axiosInstance.post(`/api/bi_guidelines/nsr-config/${encodeURIComponent(args.project_name)}/initialize`);
    const data = response.data;

    return { content: [{ type: "text", text: `NSR rules initialized for "${args.project_name}".\n\nRules created: ${data.rulesCreated}\nAll rules set to ON.` }] };
  }

  // ==================== ANALYTICS ====================

  async getNWAnalytics() {
    const response = await this.axiosInstance.get("/api/analytics/nw/projects");
    const data = response.data;

    let text = `NW Analytics: ${data.total} projects\n`;
    text += `${'='.repeat(40)}\n\n`;

    if (data.data && data.data.length > 0) {
      // Show first 10 projects
      const toShow = data.data.slice(0, 10);
      toShow.forEach((p, i) => {
        text += `${i+1}. ${p.display_name}\n`;
        text += `   Retained: ${p.retained_percentage?.toFixed(1) || 0}%\n`;
        text += `   Positive: ${p.positive_percentage?.toFixed(1) || 0}%\n`;
        text += `   Neutral: ${p.neutral_percentage?.toFixed(1) || 0}%\n`;
        text += `   Negative: ${p.negative_percentage?.toFixed(1) || 0}%\n`;
        text += `   New Names: ${p.newly_created_count || 0}\n\n`;
      });

      if (data.data.length > 10) {
        text += `... and ${data.data.length - 10} more projects`;
      }
    }

    return { content: [{ type: "text", text }] };
  }

  async getNWRegionAnalytics() {
    const response = await this.axiosInstance.get("/api/analytics/nw/regions");
    const data = response.data;

    let text = `NW Region Analytics: ${data.total} regions\n`;
    text += `${'='.repeat(40)}\n\n`;

    if (data.data && data.data.length > 0) {
      data.data.forEach((r, i) => {
        text += `${i+1}. ${r.region || r.active_lead || 'Unknown'}\n`;
        text += `   Projects: ${r.project_count || 0}\n`;
        text += `   Avg Retained: ${r.avg_retained?.toFixed(1) || 0}%\n`;
        text += `   Avg Positive: ${r.avg_positive?.toFixed(1) || 0}%\n\n`;
      });
    }

    return { content: [{ type: "text", text }] };
  }

  async getBSRAnalytics() {
    const response = await this.axiosInstance.get("/api/analytics/bsr/projects");
    const data = response.data;

    let text = `BSR Analytics: ${data.total} projects\n`;
    text += `${'='.repeat(40)}\n\n`;

    if (data.data && data.data.length > 0) {
      const toShow = data.data.slice(0, 10);
      toShow.forEach((p, i) => {
        text += `${i+1}. ${p.display_name}\n`;
        text += `   PC Access: ${p.pc_access_count || 0}\n`;
        text += `   Mobile Access: ${p.mobile_access_count || 0}\n`;
        text += `   Total: ${p.total_access_count || 0}\n\n`;
      });

      if (data.data.length > 10) {
        text += `... and ${data.data.length - 10} more projects`;
      }
    }

    return { content: [{ type: "text", text }] };
  }

  async getBSRRegionAnalytics() {
    const response = await this.axiosInstance.get("/api/analytics/bsr/regions");
    const data = response.data;

    let text = `BSR Region Analytics: ${data.total} regions\n`;
    text += `${'='.repeat(40)}\n\n`;

    if (data.data && data.data.length > 0) {
      data.data.forEach((r, i) => {
        text += `${i+1}. ${r.region || r.active_lead || 'Unknown'}\n`;
        text += `   Projects: ${r.project_count || 0}\n`;
        text += `   Total Access: ${r.total_access || 0}\n\n`;
      });
    }

    return { content: [{ type: "text", text }] };
  }

  async generateAnalyticsReport(args) {
    const { project_type, start_date, end_date } = args;
    const params = {};
    if (start_date) params.start_date = start_date;
    if (end_date) params.end_date = end_date;

    // This endpoint returns a file, so we just confirm it's available
    const response = await this.axiosInstance.get(`/api/analytics/generate/${project_type}`, {
      params,
      responseType: 'arraybuffer',
    });

    const buffer = Buffer.from(response.data);
    const filename = response.headers['content-disposition']?.match(/filename=([^;]+)/)?.[1] || `${project_type}AnalyticsReport.xlsx`;

    return { content: [{ type: "text", text: `Analytics report generated successfully!\n\nFile: ${filename}\nSize: ${(buffer.length / 1024).toFixed(1)} KB\nType: ${project_type}\n\nThe report includes:\n- Sheet 1: Project-Specific data\n- Sheet 2: Region-Specific aggregations\n\nNote: The file is ready for download through the API.` }] };
  }

  // ==================== TASK MANAGEMENT ====================

  async checkTaskStatus(args) {
    const response = await this.axiosInstance.get(`/api/presentations/tasks/${args.task_id}`);
    const data = response.data;

    let text = `Task Status: ${data.status?.toUpperCase()}\n`;
    text += `${'='.repeat(30)}\n`;
    text += `Task ID: ${data.task_id}\n`;
    text += `Type: ${data.task_type || 'N/A'}\n`;
    text += `Progress: ${data.progress || 0}%\n`;

    if (data.message) text += `Message: ${data.message}\n`;
    if (data.error) text += `Error: ${data.error}\n`;
    if (data.position) text += `Queue Position: ${data.position}\n`;
    if (data.estimated_wait_seconds) text += `Est. Wait: ${Math.round(data.estimated_wait_seconds)}s\n`;

    if (data.result) {
      text += `\nResult:\n`;
      if (data.result.presentation_id) text += `  Presentation ID: ${data.result.presentation_id}\n`;
      if (data.result.total_slides) text += `  Total Slides: ${data.result.total_slides}\n`;
    }

    return { content: [{ type: "text", text }] };
  }

  async listTasks(args) {
    const { task_type, status, limit = 100 } = args;
    const params = { limit };
    if (task_type) params.task_type = task_type;
    if (status) params.status = status;

    const response = await this.axiosInstance.get("/api/presentations/tasks", { params });
    const data = response.data;

    let text = `Tasks: ${data.total} found\n`;
    text += `${'='.repeat(30)}\n\n`;

    if (data.tasks && data.tasks.length > 0) {
      data.tasks.forEach((t, i) => {
        const statusIcon = { pending: "⏳", processing: "🔄", completed: "✅", failed: "❌" }[t.status] || "?";
        text += `${i+1}. ${statusIcon} ${t.task_id.substring(0, 8)}...\n`;
        text += `   Type: ${t.task_type}\n`;
        text += `   Status: ${t.status} (${t.progress || 0}%)\n\n`;
      });
    } else {
      text += "No tasks found.";
    }

    return { content: [{ type: "text", text }] };
  }

  async cancelTask(args) {
    const response = await this.axiosInstance.delete(`/api/presentations/tasks/${args.task_id}`);
    return { content: [{ type: "text", text: response.data.message || `Task ${args.task_id} cancelled/deleted.` }] };
  }

  async getQueueStatus() {
    const response = await this.axiosInstance.get("/api/presentations/concurrency-status");
    const data = response.data;

    let text = `System Queue Status\n`;
    text += `${'='.repeat(30)}\n`;
    text += `Active Tasks: ${data.active_tasks || 0}\n`;
    text += `Queued Tasks: ${data.queued_tasks || 0}\n`;
    text += `Available Slots: ${data.available_slots || 0}\n`;
    text += `Max Workers: ${data.max_workers || 0}\n`;
    text += `Queue Capacity: ${data.queue_capacity || 0}\n`;

    if (data.statistics) {
      text += `\nStatistics:\n`;
      text += `  Total Submitted: ${data.statistics.total_submitted || 0}\n`;
      text += `  Total Completed: ${data.statistics.total_completed || 0}\n`;
      text += `  Total Failed: ${data.statistics.total_failed || 0}\n`;
    }

    const canAccept = (data.available_slots || 0) > 0 || (data.queued_tasks || 0) < (data.queue_capacity || 10);
    text += `\n${canAccept ? '✅ System can accept new tasks' : '⚠️ System at capacity - try again later'}`;

    return { content: [{ type: "text", text }] };
  }

  // ==================== TEMPLATES ====================

  async listTemplateGroups() {
    const response = await this.axiosInstance.get("/api/bi_guidelines/template-groups");
    const data = response.data;

    let text = `Template Groups: ${data.total} templates\n`;
    text += `(${data.custom_count} custom, ${data.system_count} system)\n`;
    text += `${'='.repeat(40)}\n\n`;

    // Group by category
    const byCategory = {};
    if (data.template_groups) {
      data.template_groups.forEach(t => {
        const cat = t.category || 'Uncategorized';
        if (!byCategory[cat]) byCategory[cat] = [];
        byCategory[cat].push(t);
      });
    }

    Object.keys(byCategory).forEach(cat => {
      text += `📁 ${cat}:\n`;
      byCategory[cat].slice(0, 5).forEach(t => {
        text += `   - ${t.template_name} (ID: ${t.template_group_id})\n`;
      });
      if (byCategory[cat].length > 5) {
        text += `   ... and ${byCategory[cat].length - 5} more\n`;
      }
      text += '\n';
    });

    return { content: [{ type: "text", text }] };
  }

  async listBackgroundTemplates(args) {
    const { template_group, page = 1, limit = 50 } = args;
    const params = { page, limit };
    if (template_group) params.template_group = template_group;

    const response = await this.axiosInstance.get("/api/bi_guidelines/background-templates", { params });
    const data = response.data;

    let text = `Background Templates: ${data.total} total\n`;
    text += `${'='.repeat(40)}\n\n`;

    if (data.templates && data.templates.length > 0) {
      data.templates.forEach((t, i) => {
        text += `${i+1}. ${t.template_name}\n`;
        text += `   Category: ${t.category}\n`;
        text += `   ID: ${t.template_group_id}\n\n`;
      });
    }

    return { content: [{ type: "text", text }] };
  }

  async getBackgroundTemplate(args) {
    const response = await this.axiosInstance.get(`/api/bi_guidelines/background-templates/${args.template_id}`);
    const data = response.data;

    let text = `Template Details (ID: ${args.template_id})\n`;
    text += `${'='.repeat(30)}\n`;
    text += `Name: ${data.template_name}\n`;
    text += `Category: ${data.category}\n`;
    text += `File: ${data.template_file_name}\n`;

    return { content: [{ type: "text", text }] };
  }

  // ==================== REPORTS ====================

  async createFeedbackTemplate(args) {
    const response = await this.axiosInstance.post("/api/presentations/create-feedback-template", {
      display_name: args.display_name,
    });

    const data = response.data;
    return { content: [{ type: "text", text: `Feedback template task created!\n\nTask ID: ${data.task_id}\nStatus: ${data.status}\n\nUse check_task_status to monitor progress.` }] };
  }

  async reloadProjectSounds(args) {
    const response = await this.axiosInstance.post("/api/bi_guidelines/reload-project-sounds", {
      display_name: args.display_name,
    });

    return { content: [{ type: "text", text: `Sounds reloaded for "${args.display_name}"` }] };
  }

  // ==================== SYSTEM ====================

  async checkAPIHealth() {
    const response = await this.axiosInstance.get("/health");
    const data = response.data;

    let text = `API Health: ✅ Healthy\n`;
    text += `${'='.repeat(30)}\n`;
    text += `Name: ${data.name || data.app_name || 'Report Generator API'}\n`;
    text += `Version: ${data.version || 'N/A'}\n`;
    text += `Status: ${data.status || 'healthy'}\n`;

    return { content: [{ type: "text", text }] };
  }

  // ==================== STATISTICS ====================

  async countPresentationsByType(args) {
    const { project_type } = args;
    let text = `Presentation Counts\n${'='.repeat(30)}\n\n`;

    if (project_type === 'all' || project_type === 'NW') {
      const nwResponse = await this.axiosInstance.get("/api/bi_guidelines/nw-active-presentations", { params: { limit: 1 } });
      text += `NW Presentations: ${nwResponse.data.total}\n`;
    }

    if (project_type === 'all' || project_type === 'BSR') {
      const bsrResponse = await this.axiosInstance.get("/api/bi_guidelines/bsr-active-presentations", { params: { limit: 1 } });
      text += `BSR Presentations: ${bsrResponse.data.total}\n`;
    }

    return { content: [{ type: "text", text }] };
  }

  async getRecentPresentations(args) {
    const { project_type = 'all', limit = 10 } = args;
    let text = `Recent Presentations\n${'='.repeat(30)}\n\n`;

    const presentations = [];

    if (project_type === 'all' || project_type === 'NW') {
      const nwResponse = await this.axiosInstance.get("/api/bi_guidelines/nw-active-presentations", { params: { limit } });
      if (nwResponse.data.presentations) {
        nwResponse.data.presentations.forEach(p => {
          presentations.push({ ...p, type: 'NW' });
        });
      }
    }

    if (project_type === 'all' || project_type === 'BSR') {
      const bsrResponse = await this.axiosInstance.get("/api/bi_guidelines/bsr-active-presentations", { params: { limit } });
      if (bsrResponse.data.presentations) {
        bsrResponse.data.presentations.forEach(p => {
          presentations.push({ ...p, type: 'BSR' });
        });
      }
    }

    // Sort by update date (assuming it's available) and take top N
    const topN = presentations.slice(0, limit);

    topN.forEach((p, i) => {
      text += `${i+1}. [${p.type}] ${p.display_name || p.DisplayName}\n`;
      text += `   Project: ${p.project || p.Project}\n`;
      text += `   ID: ${p.presentation_id || p.PresentationId}\n\n`;
    });

    if (topN.length === 0) {
      text += "No recent presentations found.";
    }

    return { content: [{ type: "text", text }] };
  }

  async searchPresentations(args) {
    const { query, project_type = 'all', status = 'all', limit = 20 } = args;
    let text = `Search Results for "${query}"\n${'='.repeat(30)}\n\n`;

    const results = [];

    if (project_type === 'all' || project_type === 'NW') {
      const nwResponse = await this.axiosInstance.get("/api/bi_guidelines/nw-active-presentations", {
        params: { search: query, limit }
      });
      if (nwResponse.data.presentations) {
        nwResponse.data.presentations.forEach(p => {
          results.push({ ...p, type: 'NW' });
        });
      }
    }

    if (project_type === 'all' || project_type === 'BSR') {
      const bsrResponse = await this.axiosInstance.get("/api/bi_guidelines/bsr-active-presentations", {
        params: { search: query, limit }
      });
      if (bsrResponse.data.presentations) {
        bsrResponse.data.presentations.forEach(p => {
          results.push({ ...p, type: 'BSR' });
        });
      }
    }

    // Filter by status if needed
    let filtered = results;
    if (status !== 'all') {
      filtered = results.filter(p => (p.status || p.PresentationStatus || 'OPEN').toUpperCase() === status);
    }

    text += `Found ${filtered.length} results:\n\n`;

    filtered.slice(0, limit).forEach((p, i) => {
      text += `${i+1}. [${p.type}] ${p.display_name || p.DisplayName}\n`;
      text += `   Project: ${p.project || p.Project}\n`;
      text += `   Status: ${p.status || p.PresentationStatus || 'OPEN'}\n`;
      text += `   ID: ${p.presentation_id || p.PresentationId}\n\n`;
    });

    if (filtered.length === 0) {
      text += "No presentations found matching your criteria.";
    }

    return { content: [{ type: "text", text }] };
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error("Creative API MCP Server v2.0 running on stdio");
    console.error(`Connected to API at: ${API_BASE_URL}`);
  }
}

const server = new CreativeAPIServer();
server.run().catch(console.error);
