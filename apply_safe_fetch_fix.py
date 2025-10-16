"""Script to apply safe_fetch fix to all methods in nw_reports_service.py"""

import re
from pathlib import Path

# Read the file
file_path = Path("c:/inetpub/wwwroot/CreativePythonAPI/app/services/nw_reports_service.py")
content = file_path.read_text(encoding='utf-8')

# Methods to update (exclude those already updated)
methods_to_update = [
    ("get_roots_to_explore", "nw_dlRootsOrConceptsToExplore"),
    ("get_roots_to_avoid", "nw_dlRootsOrConceptsToAvoid"),
    ("get_open_notes", "nw_dlOpenNotes"),
    ("get_votes_by_groups", "nw_Votesbygroups"),
    ("get_voted_participants", "nw_VotedParticipants"),
]

for method_name, sp_name in methods_to_update:
    # Pattern to find the cursor.description line and replace with safe_fetch call
    pattern = r'(def ' + method_name + r'.*?with get_connection_scope.*?cursor\.execute\([^)]+\))\s+columns = \[column\[0\] for column in cursor\.description\]\s+rows = cursor\.fetchall\(\)'

    replacement = r'\1\n\n                columns, rows = _safe_fetch_sp_results(\n                    cursor, "' + sp_name + r'", presentation_id\n                )\n\n                if columns is None or rows is None:\n                    return []'

    content = re.sub(pattern, replacement, content, flags=re.DOTALL)

# Write back
file_path.write_text(content, encoding='utf-8')
print("Applied safe_fetch fixes to nw_reports_service.py")
