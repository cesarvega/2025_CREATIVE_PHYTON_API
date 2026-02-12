import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import axios from "axios";
import FormData from "form-data";
import fs from "fs";
import path from "path";

// Configure your API base URL
const API_BASE_URL = process.env.API_BASE_URL || "http://localhost:50100";

class CreativeAPIServer {
  constructor() {
    this.server = new Server(
      {
        name: "creative-api-mcp-server",
        version: "3.0.0",
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.axiosInstance = axios.create({
      baseURL: API_BASE_URL,
      timeout: 600000, // 10 minutes for long-running tasks
    });

    this.setupToolHandlers();

    this.server.onerror = (error) => console.error("[MCP Error]", error);
    process.on("SIGINT", async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: this.getToolDefinitions(),
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) =>
      this.handleToolCall(request.params.name, request.params.arguments || {})
    );
  }

  getToolDefinitions() {
    return [
      // ============================================================
      // SYSTEM & HEALTH
      // ============================================================
      {
        name: "check_api_health",
        description: "Check if the API server is running and healthy. Returns API name, version, and status.",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "get_api_info",
        description: "Get API information including documentation URLs (Swagger, Scalar).",
        inputSchema: { type: "object", properties: {} },
      },

      // ============================================================
      // DAYMASTER PROJECTS
      // ============================================================
      {
        name: "list_daymaster_projects",
        description: "List projects from DayMaster database (8000+ projects). Supports search by name/code and pagination. Use this to find projects before creating presentations.",
        inputSchema: {
          type: "object",
          properties: {
            search: { type: "string", description: "Search term to filter by project name or code" },
            page: { type: "number", description: "Page number (default: 1)" },
            limit: { type: "number", description: "Items per page (default: 10, max: 100)" },
          },
        },
      },

      // ============================================================
      // NW PRESENTATIONS (Name Evaluation)
      // ============================================================
      {
        name: "list_nw_presentations",
        description: "List active NW (Name Evaluation) presentations. Sorted by last update date (newest first). Use for browsing existing NW projects.",
        inputSchema: {
          type: "object",
          properties: {
            search: { type: "string", description: "Search by project or display name" },
            page: { type: "number", description: "Page number (default: 1)" },
            limit: { type: "number", description: "Items per page (default: 50, max: 500)" },
          },
        },
      },
      {
        name: "check_nw_exists",
        description: "Check if an NW/NSR/DW presentation already exists for a project. Returns presentation ID if found. Use before creating to avoid duplicates.",
        inputSchema: {
          type: "object",
          properties: {
            project_name: { type: "string", description: "Project name" },
            display_name: { type: "string", description: "Display name to check" },
          },
          required: ["display_name"],
        },
      },
      {
        name: "get_project_info",
        description: "Get detailed project information by ID. Auto-detects if NW, BSR, or NSR type. Returns all details including categories, status, upload info.",
        inputSchema: {
          type: "object",
          properties: {
            project_id: { type: "number", description: "Presentation/Project ID" },
          },
          required: ["project_id"],
        },
      },
      {
        name: "update_nw_project",
        description: "Update NW project details: display name, status (OPEN/CLOSED), and BSR association. Use to rename projects or close completed ones.",
        inputSchema: {
          type: "object",
          properties: {
            presentation_id: { type: "number", description: "Presentation ID to update" },
            display_name: { type: "string", description: "New display name" },
            presentation_status: { type: "string", enum: ["OPEN", "CLOSED"], description: "Project status" },
            bsr_display_name: { type: "string", description: "Associated BSR display name (optional)" },
          },
          required: ["presentation_id", "display_name", "presentation_status"],
        },
      },
      {
        name: "create_nw_presentation",
        description: "Create a new NW presentation from Excel + PowerPoint files. This is a BACKGROUND TASK - returns task_id immediately. Poll task status to check progress.",
        inputSchema: {
          type: "object",
          properties: {
            project: { type: "string", description: "Project name (from DayMaster)" },
            display_name: { type: "string", description: "Display name for the presentation" },
            excel_file_path: { type: "string", description: "Path to Excel file with name data" },
            pptx_file_path: { type: "string", description: "Path to PowerPoint template" },
            page_number: { type: "number", description: "Starting slide number (default: 1)" },
            is_phonetics: { type: "boolean", description: "Process phonetics data (default: false)" },
            has_groups: { type: "boolean", description: "Data has groups (default: false)" },
            presentation_type: { type: "string", enum: ["Normal", "Normal-NoNeutral", "Phonetics", "Katakana", "Katakana_BigJap", "Nonproprietary", "Tagline", "Design"], description: "Presentation type" },
            overwrite_existing: { type: "boolean", description: "Overwrite if exists (default: false)" },
          },
          required: ["project", "display_name", "excel_file_path", "pptx_file_path"],
        },
      },
      {
        name: "download_nw_results",
        description: "Generate and download NW results (Excel + Word in ZIP). BACKGROUND TASK. Use after presentation is completed to get results.",
        inputSchema: {
          type: "object",
          properties: {
            presentation_id: { type: "number", description: "Presentation ID" },
          },
          required: ["presentation_id"],
        },
      },
      {
        name: "create_feedback_template",
        description: "Generate NW Feedback Template document (Word). BACKGROUND TASK. Creates InputDocumentRationales with name evaluation sections.",
        inputSchema: {
          type: "object",
          properties: {
            presentation_id: { type: "number", description: "Presentation ID" },
          },
          required: ["presentation_id"],
        },
      },
      {
        name: "create_feedback_template_with_backup",
        description: "Generate Feedback Template + backup PPTX in ZIP. BACKGROUND TASK. Bundle feedback doc with presentation backup.",
        inputSchema: {
          type: "object",
          properties: {
            presentation_id: { type: "number", description: "Presentation ID" },
          },
          required: ["presentation_id"],
        },
      },
      {
        name: "reload_project_sounds",
        description: "Reload/resynchronize MP3 audio file paths for an NW project. Use when audio files have been updated or moved.",
        inputSchema: {
          type: "object",
          properties: {
            display_name: { type: "string", description: "Project display name" },
          },
          required: ["display_name"],
        },
      },

      // ============================================================
      // BSR PRESENTATIONS (Brand Strategy Report)
      // ============================================================
      {
        name: "list_bsr_presentations",
        description: "List active BSR (Brand Strategy Report) presentations. Supports search and pagination.",
        inputSchema: {
          type: "object",
          properties: {
            search: { type: "string", description: "Search term" },
            page: { type: "number", description: "Page number (default: 1)" },
            limit: { type: "number", description: "Items per page (default: 50)" },
          },
        },
      },
      {
        name: "list_bsr_display_names",
        description: "Get list of BSR display names for dropdown/autocomplete. Returns distinct names sorted alphabetically.",
        inputSchema: {
          type: "object",
          properties: {
            search: { type: "string", description: "Search term" },
            page: { type: "number", description: "Page number" },
            limit: { type: "number", description: "Items per page" },
          },
        },
      },
      {
        name: "check_bsr_exists",
        description: "Check if a BSR presentation exists. Returns presentation ID if found.",
        inputSchema: {
          type: "object",
          properties: {
            project_name: { type: "string", description: "Project name" },
            display_name: { type: "string", description: "Display name" },
          },
          required: ["display_name"],
        },
      },
      {
        name: "get_bsr_project_info",
        description: "Get detailed BSR project information including categories, slide config, and status.",
        inputSchema: {
          type: "object",
          properties: {
            project_id: { type: "number", description: "BSR Presentation ID" },
          },
          required: ["project_id"],
        },
      },
      {
        name: "update_bsr_project",
        description: "Update BSR project: display name and status (OPEN/CLOSED). Changing to CLOSED may trigger notifications.",
        inputSchema: {
          type: "object",
          properties: {
            presentation_id: { type: "number", description: "Presentation ID" },
            display_name: { type: "string", description: "New display name" },
            status: { type: "string", enum: ["OPEN", "CLOSED"], description: "Project status" },
          },
          required: ["presentation_id", "display_name", "status"],
        },
      },
      {
        name: "create_bsr_presentation",
        description: "Create a new BSR presentation from PowerPoint file. BACKGROUND TASK. BSR doesn't require Excel, only PPTX.",
        inputSchema: {
          type: "object",
          properties: {
            project_name: { type: "string", description: "Project name" },
            display_name: { type: "string", description: "Display name" },
            pptx_file_path: { type: "string", description: "Path to PowerPoint file" },
            slide_number: { type: "number", description: "Starting slide (default: 1)" },
            presentation_type: { type: "string", enum: ["BSR", "BSR-Japan"], description: "Type (default: BSR)" },
            user_name: { type: "string", description: "Uploader name" },
            is_wide_ppt: { type: "boolean", description: "Wide screen format" },
            categories: { type: "array", items: { type: "object" }, description: "Project categories" },
          },
          required: ["project_name", "display_name", "pptx_file_path"],
        },
      },
      {
        name: "generate_bsr_report",
        description: "Generate BSR report. BACKGROUND TASK.",
        inputSchema: {
          type: "object",
          properties: {
            presentation_id: { type: "number", description: "BSR Presentation ID" },
          },
          required: ["presentation_id"],
        },
      },
      {
        name: "update_presentation_categories",
        description: "Update categories for a presentation (NW or BSR).",
        inputSchema: {
          type: "object",
          properties: {
            presentation_id: { type: "number", description: "Presentation ID" },
            categories: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  category: { type: "string" },
                  elements: { type: "string" },
                },
              },
              description: "Array of category objects with 'category' and 'elements' fields",
            },
          },
          required: ["presentation_id", "categories"],
        },
      },

      // ============================================================
      // NSR CONFIGURATION (Name Safety Report)
      // ============================================================
      {
        name: "list_nsr_projects",
        description: "Get list of active NSR (Name Safety Report) projects. Use for NSR rule configuration.",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "get_nsr_config",
        description: "Get NSR validation rules for a project. Rules 101-120 control checks like USAN compliance, prefix rules, etc.",
        inputSchema: {
          type: "object",
          properties: {
            project_name: { type: "string", description: "NSR project name" },
          },
          required: ["project_name"],
        },
      },
      {
        name: "update_nsr_rule",
        description: "Update single NSR rule ON/OFF. Rules: 101-105=Contains (Y,H,W,J,K), 106/108/116/117=USAN checks, 109-115=Prefix rules.",
        inputSchema: {
          type: "object",
          properties: {
            project_name: { type: "string", description: "NSR project name" },
            rule_id: { type: "number", description: "Rule ID (101-120, except 107)" },
            is_on: { type: "number", enum: [0, 1], description: "1=ON, 0=OFF" },
          },
          required: ["project_name", "rule_id", "is_on"],
        },
      },
      {
        name: "check_nsr_rule_exists",
        description: "Check if specific NSR rule exists for project. Useful before inserting new rules.",
        inputSchema: {
          type: "object",
          properties: {
            project_name: { type: "string", description: "NSR project name" },
            rule_id: { type: "number", description: "Rule ID to check" },
          },
          required: ["project_name", "rule_id"],
        },
      },
      {
        name: "initialize_nsr_rules",
        description: "Initialize ALL NSR rules (101-120 except 107) for a new project. All rules start as ON.",
        inputSchema: {
          type: "object",
          properties: {
            project_name: { type: "string", description: "NSR project name to initialize" },
          },
          required: ["project_name"],
        },
      },
      {
        name: "bulk_update_nsr_rules",
        description: "Update multiple NSR rules at once. Efficient for toggling several rules.",
        inputSchema: {
          type: "object",
          properties: {
            project_name: { type: "string", description: "NSR project name" },
            rules: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  rule_id: { type: "number" },
                  is_on: { type: "number", enum: [0, 1] },
                },
              },
              description: "Array of {rule_id, is_on} objects",
            },
          },
          required: ["project_name", "rules"],
        },
      },

      // ============================================================
      // DW PRESENTATIONS (Design Workshop)
      // ============================================================
      {
        name: "create_dw_presentation",
        description: "Create a simple DW (Design Workshop) presentation. Simplified metadata, PPTX only.",
        inputSchema: {
          type: "object",
          properties: {
            project_name: { type: "string", description: "Project name" },
            display_name: { type: "string", description: "Display name" },
            pptx_file_path: { type: "string", description: "Path to PowerPoint file" },
            user_name: { type: "string", description: "Uploader name" },
            slide_type: { type: "string", description: "Slide type" },
            presentation_type: { type: "string", description: "Presentation type" },
            wide_presentation: { type: "boolean", description: "Wide screen format" },
          },
          required: ["project_name", "display_name", "pptx_file_path"],
        },
      },

      // ============================================================
      // BACKUP & FILE OPERATIONS
      // ============================================================
      {
        name: "generate_backup",
        description: "Generate backup PowerPoint from saved presentation files. BACKGROUND TASK.",
        inputSchema: {
          type: "object",
          properties: {
            presentation_id: { type: "number", description: "Presentation ID" },
          },
          required: ["presentation_id"],
        },
      },
      {
        name: "replace_project_images",
        description: "Regenerate project using new PPTX. Updates slide images without changing data. Optionally provide new Excel too.",
        inputSchema: {
          type: "object",
          properties: {
            project_name: { type: "string", description: "Existing project folder name" },
            project_type: { type: "string", enum: ["nw", "bsr", "dw", "nsr", "bipresents"], description: "Project type" },
            pptx_file_path: { type: "string", description: "New PPTX file path (optional, uses latest if omitted)" },
            excel_file_path: { type: "string", description: "New Excel file path (optional)" },
            page_number: { type: "number", description: "Starting slide number (optional)" },
          },
          required: ["project_name", "project_type"],
        },
      },
      {
        name: "download_presentation",
        description: "Get download URL/info for a presentation file.",
        inputSchema: {
          type: "object",
          properties: {
            presentation_id: { type: "number", description: "Presentation ID" },
          },
          required: ["presentation_id"],
        },
      },

      // ============================================================
      // BACKGROUND TEMPLATES
      // ============================================================
      {
        name: "list_template_groups",
        description: "Get background templates grouped by category. Returns template IDs, names, and counts.",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "list_background_templates",
        description: "Get paginated background templates with optional category filter. Includes preview URLs.",
        inputSchema: {
          type: "object",
          properties: {
            template_group: { type: "string", description: "Filter by group/category" },
            page: { type: "number", description: "Page number" },
            limit: { type: "number", description: "Items per page (max 100)" },
          },
        },
      },
      {
        name: "get_background_template",
        description: "Get single background template details by ID.",
        inputSchema: {
          type: "object",
          properties: {
            template_id: { type: "number", description: "Template ID" },
          },
          required: ["template_id"],
        },
      },
      {
        name: "create_background_template",
        description: "Upload new background template (JPG, PNG, or PPTX). PPTX first slide extracted as JPG.",
        inputSchema: {
          type: "object",
          properties: {
            template_name: { type: "string", description: "Template name (3-100 chars)" },
            template_group: { type: "string", description: "Group/category name (3-100 chars)" },
            file_path: { type: "string", description: "Path to image or PPTX file" },
          },
          required: ["template_name", "template_group", "file_path"],
        },
      },
      {
        name: "update_background_template",
        description: "Update template name or group. Cannot change the image itself.",
        inputSchema: {
          type: "object",
          properties: {
            template_id: { type: "number", description: "Template ID" },
            template_name: { type: "string", description: "New name (optional)" },
            template_group: { type: "string", description: "New group (optional)" },
          },
          required: ["template_id"],
        },
      },
      {
        name: "delete_background_template",
        description: "Delete background template and its physical file.",
        inputSchema: {
          type: "object",
          properties: {
            template_id: { type: "number", description: "Template ID to delete" },
          },
          required: ["template_id"],
        },
      },

      // ============================================================
      // ANALYTICS & STATISTICS
      // ============================================================
      {
        name: "get_nw_analytics",
        description: "Get analytics for ALL NW projects: vote percentages (retained, positive, neutral, negative), new names count.",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "get_nw_region_analytics",
        description: "Get NW analytics aggregated by region/active lead. Shows project counts and averages per region.",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "get_bsr_analytics",
        description: "Get analytics for ALL BSR projects: PC access, mobile access, total access counts.",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "get_bsr_region_analytics",
        description: "Get BSR analytics aggregated by region. Shows access totals per region.",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "generate_analytics_report",
        description: "Generate complete Analytics Excel report. Two sheets: Project-Specific and Region-Specific. Optional date filtering.",
        inputSchema: {
          type: "object",
          properties: {
            project_type: { type: "string", enum: ["NW", "BSR"], description: "Report type" },
            start_date: { type: "string", description: "Start date YYYY-MM-DD (optional)" },
            end_date: { type: "string", description: "End date YYYY-MM-DD (optional)" },
          },
          required: ["project_type"],
        },
      },

      // ============================================================
      // TASK MANAGEMENT
      // ============================================================
      {
        name: "get_task_status",
        description: "Get detailed status of a background task. Returns status (pending/processing/completed/failed), progress %, result or error.",
        inputSchema: {
          type: "object",
          properties: {
            task_id: { type: "string", description: "Task ID from creation response" },
          },
          required: ["task_id"],
        },
      },
      {
        name: "list_tasks",
        description: "List background tasks with optional filtering. Sorted by creation time (newest first).",
        inputSchema: {
          type: "object",
          properties: {
            task_type: { type: "string", description: "Filter: create_presentation, generate_backup, create_bsr, etc." },
            status: { type: "string", enum: ["pending", "processing", "completed", "failed"], description: "Filter by status" },
            limit: { type: "number", description: "Max tasks (default: 100)" },
          },
        },
      },
      {
        name: "cancel_task",
        description: "Cancel a queued task or delete completed/failed task. Cannot cancel already processing tasks.",
        inputSchema: {
          type: "object",
          properties: {
            task_id: { type: "string", description: "Task ID to cancel" },
          },
          required: ["task_id"],
        },
      },
      {
        name: "get_queue_status",
        description: "Get system queue status: active tasks, queued count, available slots, capacity, statistics.",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "get_task_position",
        description: "Get specific task's position in queue with estimated wait time.",
        inputSchema: {
          type: "object",
          properties: {
            task_id: { type: "string", description: "Task ID" },
          },
          required: ["task_id"],
        },
      },
      {
        name: "get_task_history",
        description: "Get recent task history with filtering options.",
        inputSchema: {
          type: "object",
          properties: {
            limit: { type: "number", description: "Max results" },
            task_type: { type: "string", description: "Filter by type" },
            status: { type: "string", description: "Filter by status" },
          },
        },
      },
      {
        name: "cleanup_old_tasks",
        description: "Remove old completed/failed tasks from system. Frees up memory.",
        inputSchema: {
          type: "object",
          properties: {
            max_age_hours: { type: "number", description: "Remove tasks older than X hours (default: 24)" },
          },
        },
      },

      // ============================================================
      // FILE PROCESSING
      // ============================================================
      {
        name: "process_excel",
        description: "Process Excel file and extract name data. Returns processed arrays for slide generation.",
        inputSchema: {
          type: "object",
          properties: {
            file_path: { type: "string", description: "Path to Excel file" },
            is_phonetics: { type: "boolean", description: "Process phonetics (default: false)" },
            has_groups: { type: "boolean", description: "Has group data (default: false)" },
          },
          required: ["file_path"],
        },
      },
      {
        name: "convert_pptx_to_images",
        description: "Convert PowerPoint slides to images. Returns image paths and slide titles.",
        inputSchema: {
          type: "object",
          properties: {
            file_path: { type: "string", description: "Path to PPTX file" },
            display_name: { type: "string", description: "Project display name" },
            project_type: { type: "string", enum: ["nw", "bsr", "dw", "nsr", "bipresents"], description: "Project type" },
          },
          required: ["file_path", "display_name", "project_type"],
        },
      },

      // ============================================================
      // SEARCH & DISCOVERY
      // ============================================================
      {
        name: "search_all_presentations",
        description: "Search across ALL presentation types (NW, BSR) by name. Returns combined results with type labels.",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Search query" },
            project_type: { type: "string", enum: ["NW", "BSR", "all"], description: "Filter by type (default: all)" },
            status: { type: "string", enum: ["OPEN", "CLOSED", "all"], description: "Filter by status (default: all)" },
            limit: { type: "number", description: "Max results (default: 20)" },
          },
          required: ["query"],
        },
      },
      {
        name: "get_recent_presentations",
        description: "Get recently updated presentations across all types. Good for activity feeds.",
        inputSchema: {
          type: "object",
          properties: {
            project_type: { type: "string", enum: ["NW", "BSR", "all"], description: "Filter by type" },
            limit: { type: "number", description: "Max results (default: 10)" },
          },
        },
      },
      {
        name: "count_presentations",
        description: "Get presentation counts by type. Useful for dashboard statistics.",
        inputSchema: {
          type: "object",
          properties: {
            project_type: { type: "string", enum: ["NW", "BSR", "all"], description: "Type to count" },
          },
          required: ["project_type"],
        },
      },

      // ============================================================
      // TESTING/DEBUG (for development)
      // ============================================================
      {
        name: "test_slide_generation",
        description: "Test slide generation logic with sample names. Development/debug tool.",
        inputSchema: {
          type: "object",
          properties: {
            template_name: { type: "string", description: "Template to test" },
            names: { type: "string", description: "Names separated by ##" },
          },
          required: ["template_name", "names"],
        },
      },
    ];
  }

  async handleToolCall(toolName, args) {
    try {
      const handler = this.toolHandlers[toolName];
      if (!handler) {
        throw new Error(`Unknown tool: ${toolName}`);
      }
      return await handler.call(this, args);
    } catch (error) {
      const errorMessage = error.response?.data?.detail || error.message;
      return {
        content: [{ type: "text", text: `❌ Error: ${errorMessage}` }],
        isError: true,
      };
    }
  }

  // Tool handler implementations
  toolHandlers = {
    // ============ SYSTEM ============
    check_api_health: async function() {
      const response = await this.axiosInstance.get("/health");
      return this.formatResponse(`API Health: ✅ Healthy\nName: ${response.data.name || 'Report Generator API'}\nVersion: ${response.data.version || 'N/A'}\nStatus: ${response.data.status || 'OK'}`);
    },

    get_api_info: async function() {
      const response = await this.axiosInstance.get("/");
      const d = response.data;
      return this.formatResponse(`API Information\n${'='.repeat(30)}\nName: ${d.name || d.app_name}\nVersion: ${d.version}\nDocs: ${d.docs_url || '/docs'}\nScalar: ${d.scalar_url || '/scalar'}`);
    },

    // ============ DAYMASTER ============
    list_daymaster_projects: async function(args) {
      const { search = "", page = 1, limit = 10 } = args;
      const response = await this.axiosInstance.get("/api/daymaster/projects", { params: { search, page, limit } });
      const d = response.data;

      let text = `📁 DayMaster Projects: ${d.total} total (page ${d.page}/${Math.ceil(d.total/d.limit)})\n\n`;
      if (d.projects?.length > 0) {
        d.projects.forEach((p, i) => {
          text += `${i+1}. ${p.ProjectName || p.project_name}\n`;
          if (p.ProjectCode) text += `   Code: ${p.ProjectCode}\n`;
        });
      } else {
        text += "No projects found.";
      }
      return this.formatResponse(text);
    },

    // ============ NW PRESENTATIONS ============
    list_nw_presentations: async function(args) {
      const { search = "", page = 1, limit = 50 } = args;
      const response = await this.axiosInstance.get("/api/bi_guidelines/nw-active-presentations", { params: { search, page, limit } });
      const d = response.data;

      let text = `📊 NW Presentations: ${d.total} total (page ${d.page})\n\n`;
      if (d.presentations?.length > 0) {
        d.presentations.forEach((p, i) => {
          text += `${i+1}. ${p.display_name || p.DisplayName}\n`;
          text += `   Project: ${p.project || p.Project}\n`;
          text += `   ID: ${p.presentation_id || p.PresentationId} | Status: ${p.status || 'OPEN'}\n\n`;
        });
      } else {
        text += "No NW presentations found.";
      }
      return this.formatResponse(text);
    },

    check_nw_exists: async function(args) {
      const params = { display_name: args.display_name };
      if (args.project_name) params.project_name = args.project_name;
      const response = await this.axiosInstance.get("/api/presentations/exists", { params });
      const exists = response.data.exists;
      return this.formatResponse(exists
        ? `✅ Presentation EXISTS\nDisplay Name: ${args.display_name}\nID: ${response.data.presentation_id}`
        : `❌ No presentation found for "${args.display_name}" - can be created.`);
    },

    get_project_info: async function(args) {
      const response = await this.axiosInstance.get(`/api/bi_guidelines/project-info/${args.project_id}`);
      const d = response.data;

      let text = `📋 Project Info (ID: ${args.project_id})\n${'='.repeat(35)}\n`;
      text += `Type: ${d.project_type || 'Unknown'}\n`;
      text += `Project: ${d.project || d.Project || 'N/A'}\n`;
      text += `Display Name: ${d.displayname || d.DisplayName || 'N/A'}\n`;
      text += `Status: ${d.presentationstatus || d.PresentationStatus || 'N/A'}\n`;
      text += `Pres. Type: ${d.presentationtype || d.PresentationType || 'N/A'}\n`;
      text += `Uploaded By: ${d.uploadedby || d.UploadedBy || 'N/A'}\n`;
      text += `Page Number: ${d.page_number || 'N/A'}\n`;

      if (d.categories?.length > 0) {
        text += `\nCategories:\n`;
        d.categories.forEach(c => text += `  • ${c.category}: ${c.elements}\n`);
      }
      return this.formatResponse(text);
    },

    update_nw_project: async function(args) {
      await this.axiosInstance.put("/api/bi_guidelines/update-project-details", {
        presentation_id: args.presentation_id,
        display_name: args.display_name,
        presentation_status: args.presentation_status,
        bsr_display_name: args.bsr_display_name || null,
      });
      return this.formatResponse(`✅ NW Project Updated\nID: ${args.presentation_id}\nDisplay Name: ${args.display_name}\nStatus: ${args.presentation_status}`);
    },

    create_nw_presentation: async function(args) {
      // Read files
      const excelContent = fs.readFileSync(args.excel_file_path);
      const pptxContent = fs.readFileSync(args.pptx_file_path);

      const formData = new FormData();
      formData.append('excel_file', excelContent, path.basename(args.excel_file_path));
      formData.append('pptx_file', pptxContent, path.basename(args.pptx_file_path));

      const metadata = {
        project: args.project,
        display_name: args.display_name,
        page_number: args.page_number || 1,
        is_phonetics: args.is_phonetics || false,
        has_groups: args.has_groups || false,
        presentation_type: args.presentation_type || "Normal",
        overwrite_existing: args.overwrite_existing || false,
      };
      formData.append('metadata', JSON.stringify(metadata));

      const response = await this.axiosInstance.post("/api/presentations/create", formData, {
        headers: formData.getHeaders(),
      });

      const d = response.data;
      return this.formatResponse(`🚀 NW Presentation Creation Started!\n\nTask ID: ${d.task_id}\nStatus: ${d.status}\nMessage: ${d.message}\n\n💡 Use get_task_status with this task_id to monitor progress.`);
    },

    download_nw_results: async function(args) {
      const response = await this.axiosInstance.post("/api/presentations/download-results", {
        presentation_id: args.presentation_id,
      });
      const d = response.data;
      return this.formatResponse(`📥 Download Results Task Created\n\nTask ID: ${d.task_id}\nStatus: ${d.status}\n\nUse get_task_status to check when ready.`);
    },

    create_feedback_template: async function(args) {
      const response = await this.axiosInstance.post("/api/presentations/create-feedback-template", {
        presentation_id: args.presentation_id,
      });
      const d = response.data;
      return this.formatResponse(`📝 Feedback Template Task Created\n\nTask ID: ${d.task_id}\nStatus: ${d.status}`);
    },

    create_feedback_template_with_backup: async function(args) {
      const response = await this.axiosInstance.post("/api/presentations/create-feedback-template-with-backup", {
        presentation_id: args.presentation_id,
      });
      const d = response.data;
      return this.formatResponse(`📦 Feedback + Backup Task Created\n\nTask ID: ${d.task_id}\nStatus: ${d.status}`);
    },

    reload_project_sounds: async function(args) {
      await this.axiosInstance.post("/api/bi_guidelines/reload-project-sounds", {
        display_name: args.display_name,
      });
      return this.formatResponse(`🔊 Sounds reloaded for "${args.display_name}"`);
    },

    // ============ BSR PRESENTATIONS ============
    list_bsr_presentations: async function(args) {
      const { search = "", page = 1, limit = 50 } = args;
      const response = await this.axiosInstance.get("/api/bi_guidelines/bsr-active-presentations", { params: { search, page, limit } });
      const d = response.data;

      let text = `📈 BSR Presentations: ${d.total} total (page ${d.page})\n\n`;
      if (d.presentations?.length > 0) {
        d.presentations.forEach((p, i) => {
          text += `${i+1}. ${p.display_name || p.DisplayName}\n`;
          text += `   Project: ${p.project || p.Project} | ID: ${p.presentation_id || p.PresentationId}\n\n`;
        });
      } else {
        text += "No BSR presentations found.";
      }
      return this.formatResponse(text);
    },

    list_bsr_display_names: async function(args) {
      const { search = "", page = 1, limit = 50 } = args;
      const response = await this.axiosInstance.get("/api/bi_guidelines/display-names", { params: { search, page, limit } });
      const d = response.data;

      let text = `📋 BSR Display Names: ${d.total} total\n\n`;
      if (d.display_names?.length > 0) {
        d.display_names.forEach((name, i) => text += `${i+1}. ${name}\n`);
      }
      return this.formatResponse(text);
    },

    check_bsr_exists: async function(args) {
      const params = { display_name: args.display_name };
      if (args.project_name) params.project_name = args.project_name;
      const response = await this.axiosInstance.get("/api/presentations/bsr/exists", { params });
      const exists = response.data.exists;
      return this.formatResponse(exists
        ? `✅ BSR Presentation EXISTS\nDisplay Name: ${args.display_name}\nID: ${response.data.presentation_id}`
        : `❌ No BSR presentation found for "${args.display_name}"`);
    },

    get_bsr_project_info: async function(args) {
      const response = await this.axiosInstance.get(`/api/bi_guidelines/bsr-project-info/${args.project_id}`);
      const d = response.data;

      let text = `📊 BSR Project Info (ID: ${args.project_id})\n${'='.repeat(35)}\n`;
      text += `Project: ${d.project || 'N/A'}\n`;
      text += `Display Name: ${d.displayname || 'N/A'}\n`;
      text += `Status: ${d.presentationstatus || 'N/A'}\n`;
      text += `Type: ${d.presentationtype || 'N/A'}\n`;
      text += `Uploaded By: ${d.uploadedby || 'N/A'}\n`;
      text += `Slide Number: ${d.slidenumber || d.page_number || 'N/A'}\n`;
      text += `Wide Screen: ${d.iswideppt ? 'Yes' : 'No'}\n`;

      if (d.categories?.length > 0) {
        text += `\nCategories:\n`;
        d.categories.forEach(c => text += `  • ${c.category}: ${c.elements}\n`);
      }
      return this.formatResponse(text);
    },

    update_bsr_project: async function(args) {
      await this.axiosInstance.put("/api/bi_guidelines/update-bsr-project", {
        presentation_id: args.presentation_id,
        display_name: args.display_name,
        status: args.status,
      });
      return this.formatResponse(`✅ BSR Project Updated\nID: ${args.presentation_id}\nDisplay Name: ${args.display_name}\nStatus: ${args.status}`);
    },

    create_bsr_presentation: async function(args) {
      const pptxContent = fs.readFileSync(args.pptx_file_path);

      const formData = new FormData();
      formData.append('pptx_file', pptxContent, path.basename(args.pptx_file_path));

      const metadata = {
        project_name: args.project_name,
        display_name: args.display_name,
        slide_number: args.slide_number || 1,
        presentation_type: args.presentation_type || "BSR",
        user_name: args.user_name || "",
        is_wide_ppt: args.is_wide_ppt || false,
        categories: args.categories || [],
      };
      formData.append('metadata', JSON.stringify(metadata));

      const response = await this.axiosInstance.post("/api/presentations/create-bsr", formData, {
        headers: formData.getHeaders(),
      });

      const d = response.data;
      return this.formatResponse(`🚀 BSR Presentation Creation Started!\n\nTask ID: ${d.task_id}\nStatus: ${d.status}`);
    },

    generate_bsr_report: async function(args) {
      const response = await this.axiosInstance.post("/api/presentations/generate-bsr-report", {
        presentation_id: args.presentation_id,
      });
      const d = response.data;
      return this.formatResponse(`📊 BSR Report Task Created\n\nTask ID: ${d.task_id}\nStatus: ${d.status}`);
    },

    update_presentation_categories: async function(args) {
      await this.axiosInstance.put(`/api/presentations/${args.presentation_id}/categories`, {
        categories: args.categories,
      });
      return this.formatResponse(`✅ Categories updated for presentation ${args.presentation_id}`);
    },

    // ============ NSR CONFIGURATION ============
    list_nsr_projects: async function() {
      const response = await this.axiosInstance.get("/api/bi_guidelines/nsr-projects");
      const d = response.data;

      let text = `🔬 Active NSR Projects:\n\n`;
      if (d.data?.length > 0) {
        d.data.forEach((p, i) => text += `${i+1}. ${p}\n`);
      } else {
        text += "No active NSR projects found.";
      }
      return this.formatResponse(text);
    },

    get_nsr_config: async function(args) {
      const response = await this.axiosInstance.get(`/api/bi_guidelines/nsr-config/${encodeURIComponent(args.project_name)}`);
      const d = response.data;

      const ruleDesc = {
        101: "Contains Y", 102: "Contains H", 103: "Contains W", 104: "Contains J", 105: "Contains K",
        106: "USAN Violation", 108: "USAN Nomenclature", 109: "Prefix AR", 110: "Prefix DEX",
        111: "Prefix ES", 112: "Prefix STR", 113: "Prefix LEV", 114: "Prefix X", 115: "Prefix RAC",
        116: "USAN Search", 117: "USAN MedNet", 118: "Double Letter", 119: "Double Letter 2", 120: "Prefix ZU"
      };

      let text = `🔧 NSR Config: ${d.project_name}\n${'='.repeat(35)}\n`;
      if (d.rules?.length > 0) {
        d.rules.forEach(r => {
          const desc = ruleDesc[r.rule_id] || `Rule ${r.rule_id}`;
          const icon = r.is_on ? "✅" : "❌";
          text += `${icon} ${r.rule_id}: ${desc}\n`;
        });
      } else {
        text += "No rules configured (all OFF by default)";
      }
      return this.formatResponse(text);
    },

    update_nsr_rule: async function(args) {
      await this.axiosInstance.put("/api/bi_guidelines/nsr-config", {
        project_name: args.project_name,
        rule_id: args.rule_id,
        is_on: args.is_on,
      });
      const status = args.is_on ? "ON ✅" : "OFF ❌";
      return this.formatResponse(`NSR Rule ${args.rule_id} for "${args.project_name}" set to ${status}`);
    },

    check_nsr_rule_exists: async function(args) {
      const response = await this.axiosInstance.get(`/api/bi_guidelines/nsr-config/${encodeURIComponent(args.project_name)}/${args.rule_id}/exists`);
      const exists = response.data.exists;
      return this.formatResponse(exists
        ? `✅ Rule ${args.rule_id} EXISTS for "${args.project_name}"`
        : `❌ Rule ${args.rule_id} NOT found for "${args.project_name}"`);
    },

    initialize_nsr_rules: async function(args) {
      const response = await this.axiosInstance.post(`/api/bi_guidelines/nsr-config/${encodeURIComponent(args.project_name)}/initialize`);
      const d = response.data;
      return this.formatResponse(`✅ NSR Rules Initialized\nProject: ${args.project_name}\nRules Created: ${d.rulesCreated}\nAll rules set to ON`);
    },

    bulk_update_nsr_rules: async function(args) {
      const results = [];
      for (const rule of args.rules) {
        await this.axiosInstance.put("/api/bi_guidelines/nsr-config", {
          project_name: args.project_name,
          rule_id: rule.rule_id,
          is_on: rule.is_on,
        });
        results.push(`Rule ${rule.rule_id}: ${rule.is_on ? 'ON' : 'OFF'}`);
      }
      return this.formatResponse(`✅ Bulk NSR Update Complete\nProject: ${args.project_name}\n\n${results.join('\n')}`);
    },

    // ============ DW PRESENTATIONS ============
    create_dw_presentation: async function(args) {
      const pptxContent = fs.readFileSync(args.pptx_file_path);

      const formData = new FormData();
      formData.append('powerpointFile', pptxContent, path.basename(args.pptx_file_path));

      const metadata = {
        projectName: args.project_name,
        displayName: args.display_name,
        userName: args.user_name || "",
        slideType: args.slide_type || "Image",
        presentationType: args.presentation_type || "DW",
        widePresentation: args.wide_presentation || false,
      };
      formData.append('metadata', JSON.stringify(metadata));

      const response = await this.axiosInstance.post("/api/presentations/create-simple-dw", formData, {
        headers: formData.getHeaders(),
      });

      return this.formatResponse(`✅ DW Presentation Created\n\n${JSON.stringify(response.data, null, 2)}`);
    },

    // ============ BACKUP & FILES ============
    generate_backup: async function(args) {
      const response = await this.axiosInstance.post("/api/presentations/backup", {
        presentation_id: args.presentation_id,
      });
      const d = response.data;
      return this.formatResponse(`📦 Backup Generation Started\n\nTask ID: ${d.task_id}\nStatus: ${d.status}`);
    },

    replace_project_images: async function(args) {
      const formData = new FormData();
      formData.append('project_name', args.project_name);
      formData.append('project_type', args.project_type);

      if (args.pptx_file_path) {
        const pptxContent = fs.readFileSync(args.pptx_file_path);
        formData.append('powerpointFile', pptxContent, path.basename(args.pptx_file_path));
      }
      if (args.excel_file_path) {
        const excelContent = fs.readFileSync(args.excel_file_path);
        formData.append('excelFile', excelContent, path.basename(args.excel_file_path));
      }
      if (args.page_number) formData.append('page_number', args.page_number.toString());

      const response = await this.axiosInstance.post("/api/presentations/replace-project-images/", formData, {
        headers: formData.getHeaders(),
      });

      return this.formatResponse(`✅ Project Images Replaced\n\n${JSON.stringify(response.data, null, 2)}`);
    },

    download_presentation: async function(args) {
      const response = await this.axiosInstance.get(`/api/presentations/download-presentation/${args.presentation_id}`, {
        responseType: 'arraybuffer',
      });
      const size = Buffer.from(response.data).length;
      return this.formatResponse(`📥 Presentation Download Ready\n\nID: ${args.presentation_id}\nSize: ${(size/1024).toFixed(1)} KB\nContent-Type: ${response.headers['content-type']}`);
    },

    // ============ TEMPLATES ============
    list_template_groups: async function() {
      const response = await this.axiosInstance.get("/api/bi_guidelines/template-groups");
      const d = response.data;

      let text = `🎨 Template Groups: ${d.total} templates\n(${d.custom_count} custom, ${d.system_count} system)\n${'='.repeat(35)}\n\n`;

      const byCategory = {};
      d.template_groups?.forEach(t => {
        const cat = t.category || 'Uncategorized';
        if (!byCategory[cat]) byCategory[cat] = [];
        byCategory[cat].push(t);
      });

      Object.keys(byCategory).forEach(cat => {
        text += `📁 ${cat} (${byCategory[cat].length}):\n`;
        byCategory[cat].slice(0, 3).forEach(t => text += `   • ${t.template_name}\n`);
        if (byCategory[cat].length > 3) text += `   ... +${byCategory[cat].length - 3} more\n`;
        text += '\n';
      });

      return this.formatResponse(text);
    },

    list_background_templates: async function(args) {
      const params = { page: args.page || 1, limit: args.limit || 50 };
      if (args.template_group) params.template_group = args.template_group;

      const response = await this.axiosInstance.get("/api/bi_guidelines/background-templates", { params });
      const d = response.data;

      let text = `🖼️ Background Templates: ${d.total} total\n\n`;
      d.templates?.forEach((t, i) => {
        text += `${i+1}. ${t.template_name}\n`;
        text += `   Category: ${t.category} | ID: ${t.template_group_id}\n\n`;
      });

      return this.formatResponse(text);
    },

    get_background_template: async function(args) {
      const response = await this.axiosInstance.get(`/api/bi_guidelines/background-templates/${args.template_id}`);
      const d = response.data;
      return this.formatResponse(`🖼️ Template Details\n\nID: ${d.template_group_id}\nName: ${d.template_name}\nCategory: ${d.category}\nFile: ${d.template_file_name}`);
    },

    create_background_template: async function(args) {
      const fileContent = fs.readFileSync(args.file_path);

      const formData = new FormData();
      formData.append('template_name', args.template_name);
      formData.append('template_group', args.template_group);
      formData.append('background_file', fileContent, path.basename(args.file_path));

      const response = await this.axiosInstance.post("/api/bi_guidelines/background-templates", formData, {
        headers: formData.getHeaders(),
      });

      return this.formatResponse(`✅ Template Created\n\nName: ${args.template_name}\nGroup: ${args.template_group}\nID: ${response.data.template?.template_group_id}`);
    },

    update_background_template: async function(args) {
      const formData = new FormData();
      if (args.template_name) formData.append('template_name', args.template_name);
      if (args.template_group) formData.append('template_group', args.template_group);

      await this.axiosInstance.patch(`/api/bi_guidelines/background-templates/${args.template_id}`, formData, {
        headers: formData.getHeaders(),
      });

      return this.formatResponse(`✅ Template ${args.template_id} updated`);
    },

    delete_background_template: async function(args) {
      await this.axiosInstance.delete(`/api/bi_guidelines/background-templates/${args.template_id}`);
      return this.formatResponse(`🗑️ Template ${args.template_id} deleted`);
    },

    // ============ ANALYTICS ============
    get_nw_analytics: async function() {
      const response = await this.axiosInstance.get("/api/analytics/nw/projects");
      const d = response.data;

      let text = `📊 NW Analytics: ${d.total} projects\n${'='.repeat(35)}\n\n`;
      d.data?.slice(0, 10).forEach((p, i) => {
        text += `${i+1}. ${p.display_name}\n`;
        text += `   Retained: ${p.retained_percentage?.toFixed(1) || 0}% | Positive: ${p.positive_percentage?.toFixed(1) || 0}%\n`;
        text += `   Neutral: ${p.neutral_percentage?.toFixed(1) || 0}% | Negative: ${p.negative_percentage?.toFixed(1) || 0}%\n`;
        text += `   New Names: ${p.newly_created_count || 0}\n\n`;
      });
      if (d.data?.length > 10) text += `... +${d.data.length - 10} more projects`;

      return this.formatResponse(text);
    },

    get_nw_region_analytics: async function() {
      const response = await this.axiosInstance.get("/api/analytics/nw/regions");
      const d = response.data;

      let text = `🌍 NW Region Analytics: ${d.total} regions\n${'='.repeat(35)}\n\n`;
      d.data?.forEach((r, i) => {
        text += `${i+1}. ${r.region || r.active_lead || 'Unknown'}\n`;
        text += `   Projects: ${r.project_count || 0}\n`;
        text += `   Avg Retained: ${r.avg_retained?.toFixed(1) || 0}%\n`;
        text += `   Avg Positive: ${r.avg_positive?.toFixed(1) || 0}%\n\n`;
      });

      return this.formatResponse(text);
    },

    get_bsr_analytics: async function() {
      const response = await this.axiosInstance.get("/api/analytics/bsr/projects");
      const d = response.data;

      let text = `📈 BSR Analytics: ${d.total} projects\n${'='.repeat(35)}\n\n`;
      d.data?.slice(0, 10).forEach((p, i) => {
        text += `${i+1}. ${p.display_name}\n`;
        text += `   PC: ${p.pc_access_count || 0} | Mobile: ${p.mobile_access_count || 0} | Total: ${p.total_access_count || 0}\n\n`;
      });
      if (d.data?.length > 10) text += `... +${d.data.length - 10} more projects`;

      return this.formatResponse(text);
    },

    get_bsr_region_analytics: async function() {
      const response = await this.axiosInstance.get("/api/analytics/bsr/regions");
      const d = response.data;

      let text = `🌍 BSR Region Analytics: ${d.total} regions\n${'='.repeat(35)}\n\n`;
      d.data?.forEach((r, i) => {
        text += `${i+1}. ${r.region || r.active_lead || 'Unknown'}\n`;
        text += `   Projects: ${r.project_count || 0} | Total Access: ${r.total_access || 0}\n\n`;
      });

      return this.formatResponse(text);
    },

    generate_analytics_report: async function(args) {
      const params = {};
      if (args.start_date) params.start_date = args.start_date;
      if (args.end_date) params.end_date = args.end_date;

      const response = await this.axiosInstance.get(`/api/analytics/generate/${args.project_type}`, {
        params,
        responseType: 'arraybuffer',
      });

      const size = Buffer.from(response.data).length;
      return this.formatResponse(`📊 Analytics Report Generated\n\nType: ${args.project_type}\nSize: ${(size/1024).toFixed(1)} KB\nSheets: Project-Specific, Region-Specific\n\n✅ Report ready for download through API`);
    },

    // ============ TASK MANAGEMENT ============
    get_task_status: async function(args) {
      const response = await this.axiosInstance.get(`/api/presentations/tasks/${args.task_id}`);
      const d = response.data;

      const statusIcons = { pending: "⏳", processing: "🔄", completed: "✅", failed: "❌" };
      let text = `${statusIcons[d.status] || "?"} Task Status: ${d.status?.toUpperCase()}\n${'='.repeat(30)}\n`;
      text += `Task ID: ${d.task_id}\n`;
      text += `Type: ${d.task_type || 'N/A'}\n`;
      text += `Progress: ${d.progress || 0}%\n`;
      if (d.message) text += `Message: ${d.message}\n`;
      if (d.error) text += `Error: ${d.error}\n`;
      if (d.position) text += `Queue Position: ${d.position}\n`;
      if (d.estimated_wait_seconds) text += `Est. Wait: ${Math.round(d.estimated_wait_seconds)}s\n`;

      if (d.result) {
        text += `\n📋 Result:\n`;
        if (d.result.presentation_id) text += `  Presentation ID: ${d.result.presentation_id}\n`;
        if (d.result.total_slides) text += `  Total Slides: ${d.result.total_slides}\n`;
        if (d.result.file_path) text += `  File: ${d.result.file_path}\n`;
      }

      return this.formatResponse(text);
    },

    list_tasks: async function(args) {
      const params = { limit: args.limit || 100 };
      if (args.task_type) params.task_type = args.task_type;
      if (args.status) params.status = args.status;

      const response = await this.axiosInstance.get("/api/presentations/tasks", { params });
      const d = response.data;

      const statusIcons = { pending: "⏳", processing: "🔄", completed: "✅", failed: "❌" };
      let text = `📋 Tasks: ${d.total} found\n${'='.repeat(30)}\n\n`;

      d.tasks?.forEach((t, i) => {
        text += `${i+1}. ${statusIcons[t.status] || "?"} ${t.task_id.substring(0, 12)}...\n`;
        text += `   Type: ${t.task_type} | Progress: ${t.progress || 0}%\n\n`;
      });

      if (!d.tasks?.length) text += "No tasks found.";

      return this.formatResponse(text);
    },

    cancel_task: async function(args) {
      await this.axiosInstance.delete(`/api/presentations/tasks/${args.task_id}`);
      return this.formatResponse(`🗑️ Task ${args.task_id} cancelled/deleted`);
    },

    get_queue_status: async function() {
      const response = await this.axiosInstance.get("/api/presentations/concurrency-status");
      const d = response.data;

      let text = `⚙️ Queue Status\n${'='.repeat(30)}\n`;
      text += `Active: ${d.active_tasks || 0}\n`;
      text += `Queued: ${d.queued_tasks || 0}\n`;
      text += `Available: ${d.available_slots || 0}\n`;
      text += `Max Workers: ${d.max_workers || 0}\n`;
      text += `Capacity: ${d.queue_capacity || 0}\n`;

      if (d.statistics) {
        text += `\n📊 Statistics:\n`;
        text += `  Submitted: ${d.statistics.total_submitted || 0}\n`;
        text += `  Completed: ${d.statistics.total_completed || 0}\n`;
        text += `  Failed: ${d.statistics.total_failed || 0}\n`;
      }

      const canAccept = (d.available_slots || 0) > 0;
      text += `\n${canAccept ? '✅ Can accept new tasks' : '⚠️ At capacity'}`;

      return this.formatResponse(text);
    },

    get_task_position: async function(args) {
      const response = await this.axiosInstance.get(`/api/tasks/queue/position/${args.task_id}`);
      const d = response.data;
      return this.formatResponse(`📍 Task Position\n\nTask: ${args.task_id}\nPosition: ${d.position || 'N/A'}\nEst. Wait: ${d.estimated_wait_seconds ? Math.round(d.estimated_wait_seconds) + 's' : 'N/A'}`);
    },

    get_task_history: async function(args) {
      const params = {};
      if (args.limit) params.limit = args.limit;
      if (args.task_type) params.task_type = args.task_type;
      if (args.status) params.status = args.status;

      const response = await this.axiosInstance.get("/api/tasks/queue/history", { params });
      const d = response.data;

      let text = `📜 Task History\n${'='.repeat(30)}\n\n`;
      d.tasks?.forEach((t, i) => {
        text += `${i+1}. ${t.task_type} - ${t.status}\n`;
        text += `   ID: ${t.task_id?.substring(0, 12)}...\n\n`;
      });

      return this.formatResponse(text);
    },

    cleanup_old_tasks: async function(args) {
      const params = {};
      if (args.max_age_hours) params.max_age_hours = args.max_age_hours;

      const response = await this.axiosInstance.post("/api/tasks/queue/cleanup", null, { params });
      return this.formatResponse(`🧹 Cleanup Complete\n\nRemoved: ${response.data.removed || 0} old tasks`);
    },

    // ============ FILE PROCESSING ============
    process_excel: async function(args) {
      const fileContent = fs.readFileSync(args.file_path);

      const formData = new FormData();
      formData.append('file_data', fileContent, path.basename(args.file_path));
      formData.append('is_phonetics', (args.is_phonetics || false).toString());
      formData.append('has_groups', (args.has_groups || false).toString());

      const response = await this.axiosInstance.post("/api/excel/process-excel", formData, {
        headers: formData.getHeaders(),
      });

      const d = response.data;
      let text = `📊 Excel Processed\n${'='.repeat(30)}\n`;
      text += `Names found: ${d.names?.length || 0}\n`;
      text += `Groups: ${d.groups?.length || 0}\n`;

      return this.formatResponse(text);
    },

    convert_pptx_to_images: async function(args) {
      const fileContent = fs.readFileSync(args.file_path);

      const formData = new FormData();
      formData.append('pptx_data', fileContent, path.basename(args.file_path));
      formData.append('display_name', args.display_name);
      formData.append('project_type', args.project_type);

      const response = await this.axiosInstance.post("/api/pptx/convert/", formData, {
        headers: formData.getHeaders(),
      });

      const d = response.data;
      let text = `🖼️ PPTX Converted\n${'='.repeat(30)}\n`;
      text += `Slides: ${d.images?.length || 0}\n`;
      text += `Thumbnails: ${d.thumbnails?.length || 0}\n`;

      if (d.slide_titles?.length > 0) {
        text += `\nSlide Titles:\n`;
        d.slide_titles.slice(0, 5).forEach((t, i) => text += `  ${i+1}. ${t}\n`);
        if (d.slide_titles.length > 5) text += `  ... +${d.slide_titles.length - 5} more\n`;
      }

      return this.formatResponse(text);
    },

    // ============ SEARCH & DISCOVERY ============
    search_all_presentations: async function(args) {
      const { query, project_type = 'all', status = 'all', limit = 20 } = args;
      const results = [];

      if (project_type === 'all' || project_type === 'NW') {
        const nw = await this.axiosInstance.get("/api/bi_guidelines/nw-active-presentations", { params: { search: query, limit } });
        nw.data.presentations?.forEach(p => results.push({ ...p, type: 'NW' }));
      }

      if (project_type === 'all' || project_type === 'BSR') {
        const bsr = await this.axiosInstance.get("/api/bi_guidelines/bsr-active-presentations", { params: { search: query, limit } });
        bsr.data.presentations?.forEach(p => results.push({ ...p, type: 'BSR' }));
      }

      let filtered = results;
      if (status !== 'all') {
        filtered = results.filter(p => (p.status || p.PresentationStatus || 'OPEN').toUpperCase() === status);
      }

      let text = `🔍 Search: "${query}"\nFound: ${filtered.length} results\n${'='.repeat(30)}\n\n`;

      filtered.slice(0, limit).forEach((p, i) => {
        text += `${i+1}. [${p.type}] ${p.display_name || p.DisplayName}\n`;
        text += `   Project: ${p.project || p.Project}\n`;
        text += `   ID: ${p.presentation_id || p.PresentationId} | Status: ${p.status || 'OPEN'}\n\n`;
      });

      if (filtered.length === 0) text += "No results found.";

      return this.formatResponse(text);
    },

    get_recent_presentations: async function(args) {
      const { project_type = 'all', limit = 10 } = args;
      const results = [];

      if (project_type === 'all' || project_type === 'NW') {
        const nw = await this.axiosInstance.get("/api/bi_guidelines/nw-active-presentations", { params: { limit } });
        nw.data.presentations?.forEach(p => results.push({ ...p, type: 'NW' }));
      }

      if (project_type === 'all' || project_type === 'BSR') {
        const bsr = await this.axiosInstance.get("/api/bi_guidelines/bsr-active-presentations", { params: { limit } });
        bsr.data.presentations?.forEach(p => results.push({ ...p, type: 'BSR' }));
      }

      let text = `🕐 Recent Presentations\n${'='.repeat(30)}\n\n`;

      results.slice(0, limit).forEach((p, i) => {
        text += `${i+1}. [${p.type}] ${p.display_name || p.DisplayName}\n`;
        text += `   Project: ${p.project || p.Project} | ID: ${p.presentation_id || p.PresentationId}\n\n`;
      });

      if (results.length === 0) text += "No recent presentations.";

      return this.formatResponse(text);
    },

    count_presentations: async function(args) {
      const { project_type } = args;
      let text = `📊 Presentation Counts\n${'='.repeat(30)}\n\n`;

      if (project_type === 'all' || project_type === 'NW') {
        const nw = await this.axiosInstance.get("/api/bi_guidelines/nw-active-presentations", { params: { limit: 1 } });
        text += `NW Presentations: ${nw.data.total}\n`;
      }

      if (project_type === 'all' || project_type === 'BSR') {
        const bsr = await this.axiosInstance.get("/api/bi_guidelines/bsr-active-presentations", { params: { limit: 1 } });
        text += `BSR Presentations: ${bsr.data.total}\n`;
      }

      return this.formatResponse(text);
    },

    // ============ TESTING ============
    test_slide_generation: async function(args) {
      const response = await this.axiosInstance.post("/api/presentations/test-slide-generation", null, {
        params: { template_name: args.template_name, names: args.names },
      });
      return this.formatResponse(`🧪 Slide Test Result\n\n${JSON.stringify(response.data, null, 2)}`);
    },
  };

  formatResponse(text) {
    return { content: [{ type: "text", text }] };
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error("🚀 Creative API MCP Server v3.0 running");
    console.error(`📡 Connected to: ${API_BASE_URL}`);
    console.error(`🔧 Tools available: ${this.getToolDefinitions().length}`);
  }
}

const server = new CreativeAPIServer();
server.run().catch(console.error);
