# """
# Test script to verify Excel processing functionality.
# """

# import io

# import requests
# from openpyxl.workbook import Workbook


# def create_test_excel():
#     """Create a simple test Excel file."""
#     workbook = Workbook()
#     worksheet = workbook.active

#     # Headers
#     headers = ["SEQ", "Category", "Name", "Rationale", "Group1", "Group2", "Katakana"]
#     for col, header in enumerate(headers, 1):
#         worksheet.cell(row=1, column=col, value=header)

#     # Sample data
#     data = [
#         [1, "Name Category", "TestName1", "Reason1", "", "", "カタカナ1"],
#         [2, "Name Category", "TestName2", "Reason2", "GroupA", "", ""],
#         [3, "Name Category", "TestName3", "Reason3", "GroupA", "", ""],
#         ["", "New Category", "", "", "", "", ""],
#         [4, "New Category", "NewName1", "NewReason1", "", "GroupB", ""],
#         [5, "New Category", "NewName2", "NewReason2", "", "GroupB", "カタカナ2"],
#     ]

#     for row_idx, row_data in enumerate(data, 2):
#         for col_idx, value in enumerate(row_data, 1):
#             worksheet.cell(row=row_idx, column=col_idx, value=value)

#     # Save to bytes
#     buffer = io.BytesIO()
#     workbook.save(buffer)
#     buffer.seek(0)
#     return buffer.getvalue()


# def test_excel_processing():
#     """Test the Excel processing endpoint."""
#     print("🧪 Testing Excel Processing API...")

#     # Create test Excel file
#     excel_data = create_test_excel()
#     print(f"📄 Created test Excel file ({len(excel_data)} bytes)")

#     # Prepare request
#     url = "http://localhost:50100/api/excel/process-excel"
#     files = {
#         "file": (
#             "test_data.xlsx",
#             excel_data,
#             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         )
#     }
#     data = {
#         "is_phonetics": False,
#         "has_groups": False,
#     }

#     try:
#         print("🌐 Sending request to API...")
#         response = requests.post(url, files=files, data=data, timeout=30)

#         if response.status_code == 200:
#             result = response.json()
#             print(f"✅ SUCCESS! {result['message']}")
#             print(f"📋 Processing ID: {result.get('processing_id', 'N/A')}")

#             data = result["data"]
#             print(f"📊 Processed {data['total_rows_processed']} rows")
#             print(f"📊 Max item number: {data['lst_max_item_number']}")

#             print("\n📋 Processed Data Arrays:")
#             print(f"  Categories: {len(data['lst_categories'])} items")
#             print(f"  Names: {len(data['lst_names'])} items")
#             print(f"  Rationales: {len(data['lst_rationales'])} items")
#             print(f"  Notations: {len(data['lst_notations'])} items")
#             print(f"  Types: {len(data['lst_types'])} items")
#             print(f"  Kana: {len(data['lst_kana'])} items")
#             print(f"  Logos: {len(data['lst_logos'])} items")
#             print(f"  Name Sub Groups: {len(data['lst_name_sub_groups'])} items")

#             # Show first few items as example
#             if data['lst_names']:
#                 print("\n📝 Sample Data:")
#                 for i in range(min(3, len(data['lst_names']))):
#                     if data['lst_names'][i]:
#                         print(f"  Row {i+1}: {data['lst_names'][i]}")
#                         if data['lst_categories'][i]:
#                             print(f"    Category: {data['lst_categories'][i]}")
#                         if data['lst_rationales'][i]:
#                             print(f"    Rationale: {data['lst_rationales'][i]}")
#                         print()

#             return True
#         else:
#             print(f"❌ ERROR: HTTP {response.status_code}")
#             print(f"Response: {response.text}")
#             return False

#     except requests.exceptions.ConnectionError:
#         print("❌ ERROR: Could not connect to server. Is it running on port 50100?")
#         return False
#     except (requests.exceptions.RequestException, ValueError, KeyError) as e:
#         print(f"❌ ERROR: {str(e)}")
#         return False


# def test_info_endpoints():
#     """Test the info endpoints."""
#     print("\n🔍 Testing Info Endpoints...")

#     endpoints = [
#         ("http://localhost:50100/api/excel/slide-types", "Slide Types"),
#         ("http://localhost:50100/api/excel/excel-format", "Excel Format"),
#     ]

#     for url, name in endpoints:
#         try:
#             response = requests.get(url, timeout=10)
#             if response.status_code == 200:
#                 print(f"✅ {name}: OK")
#             else:
#                 print(f"❌ {name}: HTTP {response.status_code}")
#         except (requests.exceptions.RequestException, ValueError) as e:
#             print(f"❌ {name}: {str(e)}")


# if __name__ == "__main__":
#     print("🚀 Excel Processing API Test Suite")
#     print("=" * 50)

#     # Test main processing endpoint
#     success = test_excel_processing()

#     print("\n" + "=" * 50)
#     if success:
#         print("🎉 All tests completed successfully!")
#     else:
#         print("⚠️  Some tests failed. Check the output above.")
