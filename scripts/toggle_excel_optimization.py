"""Script to toggle between optimized and original Excel report generator."""

import sys
from pathlib import Path

# Get the project root
project_root = Path(__file__).parent.parent
orchestrator_file = project_root / "app" / "services" / "report_orchestrator_service.py"

# Import lines
OPTIMIZED_IMPORT = "from app.services.excel_report_generator_optimized import excel_report_generator_optimized as excel_report_generator"
ORIGINAL_IMPORT = "from app.services.excel_report_generator import excel_report_generator"

COMMENT_LINE = "# OPTIMIZATION: Using optimized Excel generator for better performance"
ORIGINAL_COMMENT = "# Original: from app.services.excel_report_generator import excel_report_generator"


def check_current_status():
    """Check which version is currently active."""
    content = orchestrator_file.read_text(encoding='utf-8')

    if OPTIMIZED_IMPORT in content:
        return "optimized"
    elif ORIGINAL_IMPORT in content and OPTIMIZED_IMPORT not in content:
        return "original"
    else:
        return "unknown"


def switch_to_optimized():
    """Switch to optimized version."""
    content = orchestrator_file.read_text(encoding='utf-8')

    if OPTIMIZED_IMPORT in content:
        print("✅ Already using OPTIMIZED version")
        return True

    # Replace original import with optimized
    new_content = content.replace(
        ORIGINAL_IMPORT,
        f"{COMMENT_LINE}\n{ORIGINAL_COMMENT}\n{OPTIMIZED_IMPORT}"
    )

    if new_content != content:
        orchestrator_file.write_text(new_content, encoding='utf-8')
        print("✅ Switched to OPTIMIZED version")
        print("   - Parallel execution of stored procedures")
        print("   - Increased timeouts (60s)")
        print("   - Detailed timing logs")
        return True
    else:
        print("❌ Failed to switch to optimized version")
        return False


def switch_to_original():
    """Switch to original version."""
    content = orchestrator_file.read_text(encoding='utf-8')

    if ORIGINAL_IMPORT in content and OPTIMIZED_IMPORT not in content:
        print("✅ Already using ORIGINAL version")
        return True

    # Replace optimized import with original
    lines = content.split('\n')
    new_lines = []
    skip_next = 0

    for i, line in enumerate(lines):
        if skip_next > 0:
            skip_next -= 1
            continue

        if COMMENT_LINE in line:
            # Skip this line and the next 2 lines (comment and import)
            skip_next = 2
            # Add back the original import
            new_lines.append(ORIGINAL_IMPORT)
        else:
            new_lines.append(line)

    new_content = '\n'.join(new_lines)

    if new_content != content:
        orchestrator_file.write_text(new_content, encoding='utf-8')
        print("✅ Switched to ORIGINAL version")
        print("   - Sequential execution of stored procedures")
        print("   - Original timeouts (30s)")
        return True
    else:
        print("❌ Failed to switch to original version")
        return False


def main():
    """Main function."""
    if len(sys.argv) < 2:
        current = check_current_status()
        print(f"📊 Current version: {current.upper()}")
        print()
        print("Usage:")
        print("  python toggle_excel_optimization.py optimized   # Switch to optimized version")
        print("  python toggle_excel_optimization.py original    # Switch to original version")
        print("  python toggle_excel_optimization.py status      # Check current status")
        sys.exit(1)

    command = sys.argv[1].lower()

    if command == "status":
        current = check_current_status()
        print(f"📊 Current version: {current.upper()}")
        if current == "optimized":
            print("   ⚡ Using parallel execution for faster reports")
        elif current == "original":
            print("   🐌 Using sequential execution")

    elif command == "optimized":
        success = switch_to_optimized()
        if success:
            print()
            print("🔄 Please restart your API server for changes to take effect")

    elif command == "original":
        success = switch_to_original()
        if success:
            print()
            print("🔄 Please restart your API server for changes to take effect")

    else:
        print(f"❌ Unknown command: {command}")
        print("   Valid commands: optimized, original, status")
        sys.exit(1)


if __name__ == "__main__":
    main()
