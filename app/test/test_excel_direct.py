# """
# Simple test to verify Excel processing functionality without requiring a running server.
# """

# import sys
# import os
# import traceback

# # Add the project root to the Python path
# sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

# from app.models.excel_models import ExcelRowData
# from app.services.excel_service import excel_processing_service


# def create_simple_test_excel():
#     """Create test data in the expected format."""
#     return [
#         {
#             "SEQ": 1,
#             "Category": "Name Category",
#             "Name": "TestName1",
#             "Rationale": "Reason1",
#             "Katakana": "カタカナ1",
#         },
#         {
#             "SEQ": 2,
#             "Category": "Name Category",
#             "Name": "TestName2",
#             "Rationale": "Reason2",
#             "Group1": "GroupA",
#         },
#         {
#             "SEQ": 3,
#             "Category": "Name Category",
#             "Name": "TestName3",
#             "Rationale": "Reason3",
#             "Group1": "GroupA",
#         },
#         {"Category": "New Category"},
#         {
#             "SEQ": 4,
#             "Category": "New Category",
#             "Name": "NewName1",
#             "Rationale": "NewReason1",
#             "Group2": "GroupB",
#         },
#         {
#             "SEQ": 5,
#             "Category": "New Category",
#             "Name": "NewName2",
#             "Rationale": "NewReason2",
#             "Group2": "GroupB",
#         },
#     ]


# def test_excel_processing_direct():
#     """Test Excel processing service directly."""
#     print("🧪 Testing Excel Processing Service Directly...")
#     print("=" * 50)

#     try:
#         # Create test Excel file in memory
#         from openpyxl import Workbook
#         import io

#         workbook = Workbook()
#         worksheet = workbook.active

#         # Add headers
#         headers = ["SEQ", "Category", "Name", "Rationale", "Group1", "Group2", "Katakana"]
#         for col, header in enumerate(headers, 1):
#             worksheet.cell(row=1, column=col, value=header)

#         # Add test data
#         test_data = create_simple_test_excel()
#         for row_idx, row_data in enumerate(test_data, 2):
#             for col_idx, (key, value) in enumerate(row_data.items(), 1):
#                 worksheet.cell(row=row_idx, column=col_idx, value=value)

#         # Save to bytes
#         buffer = io.BytesIO()
#         workbook.save(buffer)
#         buffer.seek(0)
#         excel_bytes = buffer.getvalue()

#         print(f"📄 Created test Excel file ({len(excel_bytes)} bytes)")

#         # Test different processing modes
#         test_cases = [
#             {"is_phonetics": False, "has_groups": False, "name": "Basic Processing"},
#             {"is_phonetics": True, "has_groups": False, "name": "Phonetics Processing"},
#             {"is_phonetics": False, "has_groups": True, "name": "Groups Processing"},
#             {"is_phonetics": True, "has_groups": True, "name": "Combined Processing"},
#         ]

#         for test_case in test_cases:
#             print(f"\n🔬 Testing: {test_case['name']}")
#             print("-" * 30)

#             # Process Excel file
#             result = excel_processing_service.process_excel_file(
#                 file_content=excel_bytes,
#                 is_phonetics=test_case["is_phonetics"],
#                 has_groups=test_case["has_groups"]
#             )

#             print(f"✅ Processed {result.total_rows_processed} rows")
#             print(f"� Max item number: {result.lst_max_item_number}")
#             print(f"📋 Categories: {len(result.lst_categories)} items")
#             print(f"📋 Names: {len(result.lst_names)} items")
#             print(f"📋 Rationales: {len(result.lst_rationales)} items")

#             # Show sample data
#             if result.lst_names and len(result.lst_names) > 0:
#                 print("📝 Sample processed data:")
#                 for i in range(min(3, len(result.lst_names))):
#                     if result.lst_names[i]:
#                         print(f"  Row {i+1}: {result.lst_names[i]}")
#                         if result.lst_categories[i]:
#                             print(f"    Category: {result.lst_categories[i]}")
#                         if result.lst_rationales[i]:
#                             print(f"    Rationale: {result.lst_rationales[i]}")
#                         print()

#         print("✅ All processing modes tested successfully")
#         return True

#     except (ValueError, KeyError, AttributeError, ImportError) as e:
#         print(f"❌ Error during testing: {str(e)}")
#         traceback.print_exc()
#         return False


# if __name__ == "__main__":
#     print("🚀 Excel Processing Direct Test Suite")
#     print("=" * 50)

#     test_passed = test_excel_processing_direct()

#     print("=" * 50)
#     if test_passed:
#         print("🎉 Direct test completed successfully!")
#     else:
#         print("⚠️  Direct test failed. Check the output above.")
