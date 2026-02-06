"""
Test script for Gemini chatbot integration with Creative API

This script tests the integration between Gemini AI and your FastAPI
without needing to set up the full MCP server.

Usage:
    python test_chatbot_integration.py
"""

import os
import json
import requests
from typing import Dict, Any, List

# Configuration
API_BASE_URL = os.getenv("API_BASE_URL", "http://localhost:50100")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")  # Set this in environment or replace here

if not GEMINI_API_KEY:
    print("⚠️  Please set GEMINI_API_KEY environment variable or edit the script")
    print("   Get your key from: https://makersuite.google.com/app/apikey")
    exit(1)

GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent"

# Tool definitions matching the MCP server
TOOLS = [
    {
        "name": "list_projects",
        "description": "List projects from DayMaster database with pagination and search",
        "parameters": {
            "type": "object",
            "properties": {
                "page": {"type": "integer", "description": "Page number (default: 1)"},
                "limit": {"type": "integer", "description": "Items per page (default: 10)"},
                "search": {"type": "string", "description": "Search term"}
            }
        }
    },
    {
        "name": "list_nw_presentations",
        "description": "List active NW presentations with pagination and search",
        "parameters": {
            "type": "object",
            "properties": {
                "page": {"type": "integer"},
                "limit": {"type": "integer"},
                "search": {"type": "string"}
            }
        }
    },
    {
        "name": "check_nw_presentation_exists",
        "description": "Check if an NW presentation exists for a project",
        "parameters": {
            "type": "object",
            "properties": {
                "display_name": {"type": "string", "description": "Display name of the project"}
            },
            "required": ["display_name"]
        }
    },
    {
        "name": "check_task_status",
        "description": "Check the status of a background task",
        "parameters": {
            "type": "object",
            "properties": {
                "task_id": {"type": "string", "description": "Task ID"}
            },
            "required": ["task_id"]
        }
    },
    {
        "name": "get_queue_status",
        "description": "Get current task queue status",
        "parameters": {
            "type": "object",
            "properties": {}
        }
    }
]


def execute_tool(tool_name: str, args: Dict[str, Any]) -> Dict[str, Any]:
    """Execute a tool by calling the FastAPI endpoint"""
    try:
        if tool_name == "list_projects":
            response = requests.get(f"{API_BASE_URL}/api/daymaster/projects", params=args)

        elif tool_name == "list_nw_presentations":
            response = requests.get(f"{API_BASE_URL}/api/bi_guidelines/nw-active-presentations", params=args)

        elif tool_name == "check_nw_presentation_exists":
            response = requests.get(f"{API_BASE_URL}/api/presentations/exists", params=args)

        elif tool_name == "check_task_status":
            response = requests.get(f"{API_BASE_URL}/api/tasks/{args['task_id']}")

        elif tool_name == "get_queue_status":
            response = requests.get(f"{API_BASE_URL}/api/presentations/concurrency-status")

        else:
            return {"error": f"Unknown tool: {tool_name}"}

        response.raise_for_status()
        return response.json()

    except Exception as e:
        return {"error": str(e)}


def chat_with_gemini(message: str, conversation_history: List[Dict]) -> tuple:
    """Send a message to Gemini and handle function calling"""

    conversation_history.append({
        "role": "user",
        "parts": [{"text": message}]
    })

    payload = {
        "contents": conversation_history,
        "tools": [{
            "function_declarations": TOOLS
        }],
        "systemInstruction": {
            "parts": [{
                "text": "You are a helpful assistant for the Creative API system. Use available tools when needed."
            }]
        }
    }

    response = requests.post(
        f"{GEMINI_ENDPOINT}?key={GEMINI_API_KEY}",
        json=payload,
        headers={"Content-Type": "application/json"}
    )

    response.raise_for_status()
    data = response.json()

    candidate = data["candidates"][0]

    # Check if Gemini wants to call a function
    if "functionCall" in candidate["content"]["parts"][0]:
        function_call = candidate["content"]["parts"][0]["functionCall"]

        print(f"\n🔧 Gemini is calling: {function_call['name']}")
        print(f"   Arguments: {json.dumps(function_call['args'], indent=2)}")

        # Execute the function
        tool_result = execute_tool(function_call['name'], function_call.get('args', {}))

        print(f"\n📊 Tool Result:")
        print(f"   {json.dumps(tool_result, indent=2)[:200]}...")

        # Add function call and response to history
        conversation_history.append({
            "role": "model",
            "parts": [{"functionCall": function_call}]
        })

        conversation_history.append({
            "role": "user",
            "parts": [{
                "functionResponse": {
                    "name": function_call['name'],
                    "response": tool_result
                }
            }]
        })

        # Get Gemini's natural language response
        follow_up = requests.post(
            f"{GEMINI_ENDPOINT}?key={GEMINI_API_KEY}",
            json={
                "contents": conversation_history,
                "tools": [{"function_declarations": TOOLS}]
            }
        )

        follow_up.raise_for_status()
        follow_data = follow_up.json()
        assistant_message = follow_data["candidates"][0]["content"]["parts"][0]["text"]

        conversation_history.append({
            "role": "model",
            "parts": [{"text": assistant_message}]
        })

        return assistant_message, conversation_history

    else:
        # Direct text response
        assistant_message = candidate["content"]["parts"][0]["text"]

        conversation_history.append({
            "role": "model",
            "parts": [{"text": assistant_message}]
        })

        return assistant_message, conversation_history


def test_basic_connectivity():
    """Test basic connectivity to FastAPI"""
    print("\n" + "="*60)
    print("🔍 Testing FastAPI Connectivity")
    print("="*60)

    try:
        response = requests.get(f"{API_BASE_URL}/health", timeout=5)
        if response.ok:
            print("✅ FastAPI is responding!")
            print(f"   Response: {response.json()}")
            return True
        else:
            print(f"❌ FastAPI returned error: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Cannot connect to FastAPI: {e}")
        print(f"   Make sure it's running at {API_BASE_URL}")
        return False


def run_test_conversations():
    """Run test conversations with the chatbot"""

    if not test_basic_connectivity():
        return

    print("\n" + "="*60)
    print("🤖 Starting Chatbot Test Conversations")
    print("="*60)

    test_messages = [
        "List the first 3 projects from the database",
        "How many tasks are currently in the queue?",
        "Show me NW presentations, limit to 2 results",
    ]

    conversation_history = []

    for i, message in enumerate(test_messages, 1):
        print(f"\n{'─'*60}")
        print(f"💬 Test {i}: {message}")
        print('─'*60)

        try:
            response, conversation_history = chat_with_gemini(message, conversation_history)
            print(f"\n🤖 Response:")
            print(f"   {response}\n")

        except Exception as e:
            print(f"\n❌ Error: {e}\n")
            break

    print("\n" + "="*60)
    print("✅ Test Complete!")
    print("="*60)


def interactive_mode():
    """Run in interactive mode"""
    print("\n" + "="*60)
    print("🎯 Interactive Chatbot Test Mode")
    print("="*60)
    print("\nType your messages below. Type 'quit' to exit.\n")

    if not test_basic_connectivity():
        return

    conversation_history = []

    while True:
        try:
            user_input = input("\n💬 You: ").strip()

            if user_input.lower() in ['quit', 'exit', 'q']:
                print("\n👋 Goodbye!")
                break

            if not user_input:
                continue

            response, conversation_history = chat_with_gemini(user_input, conversation_history)
            print(f"\n🤖 Bot: {response}")

        except KeyboardInterrupt:
            print("\n\n👋 Goodbye!")
            break

        except Exception as e:
            print(f"\n❌ Error: {e}")


def main():
    """Main entry point"""
    print("""
    ╔══════════════════════════════════════════════════════════╗
    ║     Creative API Chatbot Integration Test               ║
    ╚══════════════════════════════════════════════════════════╝
    """)

    print("Select mode:")
    print("  1. Run automated tests")
    print("  2. Interactive chat mode")
    print("  3. Test API connectivity only")

    choice = input("\nEnter choice (1-3): ").strip()

    if choice == "1":
        run_test_conversations()
    elif choice == "2":
        interactive_mode()
    elif choice == "3":
        test_basic_connectivity()
    else:
        print("Invalid choice!")


if __name__ == "__main__":
    main()
