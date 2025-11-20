"""
Script to fix existing templates with incorrect paths
Updates TemplateFileName to include 'images/BackGrounds/' prefix
"""
import pyodbc

# Connection string
conn_str = "DRIVER={SQL Server};SERVER=192.168.0.85;DATABASE=BI_GUIDELINES;UID=sqlguide;PWD=sqlguidepwd;TrustServerCertificate=yes;Connection Timeout=30;Encrypt=no;"

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    
    print("=" * 80)
    print("Fixing templates with incorrect paths")
    print("=" * 80)
    
    # Find templates with just filename (no path)
    # Correct format: "images/BackGrounds/filename.ext"
    # Incorrect format: "filename.ext" or "slide_7.png"
    
    cursor.execute("""
        SELECT TemplateId, TemplateName, TemplateFileName
        FROM [dbo].[nw_Templates]
        WHERE TemplateFileName NOT LIKE 'images/%'
          AND TemplateFileName IS NOT NULL
          AND TemplateFileName <> ''
        ORDER BY TemplateId DESC
    """)
    
    templates_to_fix = cursor.fetchall()
    
    if not templates_to_fix:
        print("\n✓ No templates need fixing. All paths are correct!")
        print("=" * 80)
    else:
        print(f"\n⚠️  Found {len(templates_to_fix)} templates with incorrect paths:")
        print("-" * 80)
        print(f"{'ID':<6} {'Name':<30} {'Current Path':<40}")
        print("-" * 80)
        
        for row in templates_to_fix:
            print(f"{row[0]:<6} {row[1]:<30} {row[2]:<40}")
        
        print("\n" + "=" * 80)
        print("Fixing paths...")
        print("=" * 80)
        
        fixed_count = 0
        for row in templates_to_fix:
            template_id = row[0]
            template_name = row[1]
            current_path = row[2]
            
            # Build correct path
            correct_path = f"images/BackGrounds/{current_path}"
            
            # Update
            cursor.execute("""
                UPDATE [dbo].[nw_Templates]
                SET TemplateFileName = ?
                WHERE TemplateId = ?
            """, (correct_path, template_id))
            
            print(f"\n✓ Fixed ID {template_id} ({template_name}):")
            print(f"  Before: {current_path}")
            print(f"  After:  {correct_path}")
            
            fixed_count += 1
        
        conn.commit()
        
        print("\n" + "=" * 80)
        print(f"✓ Fixed {fixed_count} templates successfully!")
        print("=" * 80)
    
    cursor.close()
    conn.close()
    
except Exception as e:
    print(f"\n✗ Error: {e}")
    import traceback
    traceback.print_exc()
