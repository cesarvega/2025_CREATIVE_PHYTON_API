"""Test script to check nw_CombineNewNames stored procedure."""

from app.config.db import get_connection_scope
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)

def list_available_sps():
    """List all stored procedures containing 'nw_' in their name."""
    print("\n=== Listing NW Stored Procedures ===\n")

    try:
        with get_connection_scope(timeout=30) as cursor:
            cursor.execute("""
                SELECT ROUTINE_NAME
                FROM INFORMATION_SCHEMA.ROUTINES
                WHERE ROUTINE_TYPE = 'PROCEDURE'
                AND ROUTINE_SCHEMA = 'dbo'
                AND ROUTINE_NAME LIKE 'nw_%'
                ORDER BY ROUTINE_NAME
            """)

            sps = cursor.fetchall()
            print(f"Found {len(sps)} stored procedures:\n")
            for sp in sps:
                print(f"  - {sp[0]}")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

def test_newly_created_names(presentation_id: int = 8262):
    """Test the nw_CombineNewNames stored procedure."""

    print(f"\n=== Testing nw_CombineNewNames for presentation_id={presentation_id} ===\n")

    try:
        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[nw_CombineNewNames](?)}",
                (presentation_id,)
            )

            if cursor.description is None:
                print("ERROR: No result set returned (cursor.description is None)")
                print("Trying to fetch next result set...")

                # Try to get next result set
                if cursor.nextset():
                    if cursor.description is not None:
                        print("Found result set in nextset()!")
                        columns = [column[0] for column in cursor.description]
                        rows = cursor.fetchall()
                        print(f"Columns: {columns}")
                        print(f"Number of rows: {len(rows)}")

                        if rows:
                            print("\nFirst 5 rows:")
                            for i, row in enumerate(rows[:5], 1):
                                print(f"\nRow {i}:")
                                for col, val in zip(columns, row):
                                    print(f"  {col}: {val}")
                    else:
                        print("No result set in nextset() either")
                else:
                    print("No additional result sets")
                return

            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            print(f"Columns: {columns}")
            print(f"Number of rows: {len(rows)}")

            if rows:
                print("\nFirst 5 rows:")
                for i, row in enumerate(rows[:5], 1):
                    print(f"\nRow {i}:")
                    for col, val in zip(columns, row):
                        print(f"  {col}: {val}")
            else:
                print("\nNo rows returned from stored procedure")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

def test_dl_newly_created_names(presentation_id: int = 8262):
    """Test the nw_dlNewlyCreatedNames stored procedure."""

    print(f"\n=== Testing nw_dlNewlyCreatedNames for presentation_id={presentation_id} ===\n")

    try:
        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[nw_dlNewlyCreatedNames](?)}",
                (presentation_id,)
            )

            if cursor.description is None:
                print("ERROR: No result set returned (cursor.description is None)")
                return

            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            print(f"Columns: {columns}")
            print(f"Number of rows: {len(rows)}")

            if rows:
                print("\nFirst 10 rows:")
                for i, row in enumerate(rows[:10], 1):
                    print(f"\nRow {i}:")
                    for col, val in zip(columns, row):
                        print(f"  {col}: {val}")
            else:
                print("\nNo rows returned from stored procedure")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

def test_combine_new_names_with_nextset(presentation_id: int = 8262):
    """Test nw_CombineNewNames with multiple result sets."""

    print(f"\n=== Testing nw_CombineNewNames with all result sets for presentation_id={presentation_id} ===\n")

    try:
        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[nw_CombineNewNames](?)}",
                (presentation_id,)
            )

            result_set_num = 1
            while True:
                if cursor.description is not None:
                    columns = [column[0] for column in cursor.description]
                    rows = cursor.fetchall()

                    print(f"\n--- Result Set {result_set_num} ---")
                    print(f"Columns: {columns}")
                    print(f"Number of rows: {len(rows)}")

                    if rows:
                        print("\nFirst 5 rows:")
                        for i, row in enumerate(rows[:5], 1):
                            print(f"\nRow {i}:")
                            for col, val in zip(columns, row):
                                print(f"  {col}: {val}")

                    result_set_num += 1
                else:
                    print(f"\n--- Result Set {result_set_num} ---")
                    print("No columns (likely an UPDATE/INSERT/DELETE statement)")
                    result_set_num += 1

                # Try to move to next result set
                if not cursor.nextset():
                    break

            print(f"\nTotal result sets processed: {result_set_num - 1}")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # list_available_sps()
    # test_newly_created_names()
    # test_dl_newly_created_names()
    test_combine_new_names_with_nextset()
