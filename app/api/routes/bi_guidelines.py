"""Routes for BI Guidelines API.

This module provides endpoints to query and create BI Guideline presentations
from the BI_GUIDELINES database.
"""

import logging

import pyodbc
from fastapi import APIRouter, HTTPException
from fastapi.responses import JSONResponse

from app.config.settings import settings
from app.models.nw_master_request import PresentationData

router = APIRouter(prefix="/bi_guidelines", tags=["BI Guidelines"])
logger = logging.getLogger(__name__)


def get_db_connection():
    """Return a connection to the BI_GUIDELINES database."""
    connection_string = settings.sql_connection_string
    try:
        conn = pyodbc.connect(connection_string)  # pylint: disable=c-extension-no-member
        logger.info("Database connection established successfully.")
        return conn
    except pyodbc.Error as e:  # pylint: disable=c-extension-no-member
        logger.error("Database connection error: %s", e)
        return None


class PresentationNotFoundError(Exception):
    """Custom exception for when a presentation is not found."""


@router.post("/presentations/", status_code=201)
async def create_presentation(data: PresentationData):
    """
    Creates a new master presentation record and its associated detail records.

    This endpoint handles the creation of presentation records:
    1. Starts a transaction.
    2. Calls the `nw_InsertPresentationMaster_sep2025` stored procedure
       to create the master record.
    3. Gets the returned `PresentationId`.
    4. Iterates over the details and calls `nw_InsertPresentationDetail_copy`
       for each one.
    5. If everything is successful, it commits the transaction.
       If anything fails, it rolls back.
    """
    conn = None
    try:
        conn = get_db_connection()
        if not conn:
            raise HTTPException(
                status_code=500,
                detail="Could not connect to database."
            )

        cursor = conn.cursor()

        cursor.execute("SELECT DB_NAME(), @@SERVERNAME;")
        db_row = cursor.fetchone()
        print(f"Connected to DB: {db_row[0]} on Server: {db_row[1]}")

        # Start transaction
        cursor.execute(
            "SET NOCOUNT ON;"
        )  # Prevents empty result sets from 'X rows affected'

        # 1. Insert into the master table
        master_sql = """
            EXEC [dbo].[nw_InsertPresentationMaster_sep2025]
                @Project=?, @DisplayName=?, @MainPptFileName=?, @NameCandidateFileName=?,
                @NameCandidateBGType=?, @NameCandidateBGName=?, @NameCandidateStartingSlide=?,
                @PresentationType=?, @UploadedBy=?, @BSRDisplayName=?,
                @isParticipantsVote=?, @isWideScreenPPT=?, @isAWSLinkReq=?;
        """
        master_params = (
            data.project,
            data.display_name,
            data.powerpoint_file,
            data.excel_file,
            data.background_type,
            data.background_name,
            data.page_number,
            data.presentation_type,
            data.user_name,
            data.bsr_display_name,
            data.participant_vote,
            data.is_wide_ppt,
            data.is_aws_email,
        )

        print("Executing master stored procedure...")
        cursor.execute(master_sql, master_params)

        # Force moving through possible result sets until the final SELECT is found
        presentation_id = None
        while True:
            try:
                row = cursor.fetchone()
                if row:
                    presentation_id = row[0]
                    print(f"PresentationId found: {presentation_id}")
                    break
            except pyodbc.Error:  # pylint: disable=c-extension-no-member
                pass
            if not cursor.nextset():
                break

        if not presentation_id:
            raise PresentationNotFoundError(
                "Could not get PresentationId from the master record."
            )

        # 2. Insert into the details table
        detail_sql = """
            EXEC [dbo].[nw_InsertPresentationDetail_copy]
                @PresentationId=?, @SlideNumber=?, @SlideType=?, @SlideBGFileName=?,
                @SlideDescription=?, @NameGroup=?, @NameCategory=?, @Name=?,
                @NameRationale=?, @NameNotation=?, @KanaNames=?, @NameLogo=?,
                @TemplateId=?, @NameSubGroup=?;
        """
        for item in data.details:
            detail_params = (
                presentation_id,
                item.slide_number,
                item.slide_type,
                item.slide_bg_file_name,
                item.slide_description,
                item.group_name,
                item.category,
                item.name,
                item.rationale,
                item.notation,
                item.kana,
                item.logo_filename,
                item.template_id,
                item.name_sub_group,
            )
            cursor.execute(detail_sql, detail_params)

        # Commit the transaction
        conn.commit()
        print("Changes committed to database.")

        return {
            "message": "Presentation created successfully.",
            "presentation_id": presentation_id,
        }

    except pyodbc.Error as e:  # pylint: disable=c-extension-no-member
        logger.exception("Transaction failed: %s", e)
        if conn:
            conn.rollback()
        raise HTTPException(status_code=500, detail=f"Transaction error: {e}") from e

    except PresentationNotFoundError as e:
        logger.error("Presentation creation failed: %s", e)
        if conn:
            conn.rollback()
        raise HTTPException(status_code=500, detail=str(e)) from e

    finally:
        if conn:
            conn.close()


@router.get("/presentations/{presentation_id}")
async def get_presentation(presentation_id: int):
    """
    Retrieves a presentation by its ID, including master and detail records.
    """
    conn = get_db_connection()
    if not conn:
        return JSONResponse(
            status_code=500, content={"error": "Could not connect to database."}
        )
    try:
        cursor = conn.cursor()

        # Get master record
        master_query = """
        SELECT [PresentationId], [Project], [DisplayName], [MainPptFileName],
               [NameCandidateFileName], [NameCandidateBGType], [NameCandidateBGName],
               [NameCandidateStartingSlide], [UploadedBy], [UploadedDate],
               [PresentationStatus], [PresentationOpenDate], [NotesExplore],
               [NotesAvoid], [LastUpdateDate], [PresentationType], [BSRDisplayName],
               [isParticipantsVote], [show_EngKat_in_groups], [isWideScreenPPT],
               [isAWSLinkReq], [isWide]
        FROM [BI_GUIDELINES].[dbo].[nw_Master]
        WHERE PresentationId = ?
        """
        cursor.execute(master_query, (presentation_id,))
        master_columns = [column[0] for column in cursor.description]
        master_row = cursor.fetchone()
        if not master_row:
            return JSONResponse(
                status_code=404, content={"error": "Presentation not found."}
            )
        master_data = dict(zip(master_columns, master_row))

        # Get detail records
        detail_query = """
        SELECT [PresentationId], [SlideNumber], [SlideType], [SlideBGFileName],
               [SlideDescription], [NameGroup], [NameCategory], [Name],
               [NameRationale], [NameNotation], [KanaNames], [NameLogo],
               [TemplateId], [NameSubGroup]
        FROM [BI_GUIDELINES].[dbo].[nw_Details]
        WHERE PresentationId = ?
        ORDER BY SlideNumber
        """
        cursor.execute(detail_query, (presentation_id,))
        detail_columns = [column[0] for column in cursor.description]
        detail_rows = cursor.fetchall()
        detail_data = [dict(zip(detail_columns, row)) for row in detail_rows]

        cursor.close()
        conn.close()

        return {"master": master_data, "details": detail_data}

    except pyodbc.Error as e:  # pylint: disable=c-extension-no-member
        logger.error("Database query failed: %s", e)
        return JSONResponse(status_code=500, content={"error": str(e)})


# --- Data Models (Pydantic) ---
# Define the structure of the JSON that the API expects to receive.
