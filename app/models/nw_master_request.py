from pydantic import BaseModel
from typing import List, Optional

class DetailItem(BaseModel):
    """Represents a single record in the nw_Details table."""
    slide_number: int
    slide_type: str
    slide_bg_file_name: Optional[str] = None
    slide_description: Optional[str] = None
    group_name: Optional[str] = None
    category: Optional[str] = None
    name: Optional[str] = None
    rationale: Optional[str] = None
    notation: Optional[str] = None
    kana: Optional[str] = None
    logo_filename: Optional[str] = None
    template_id: int
    name_sub_group: Optional[str] = None

class PresentationData(BaseModel):
    """Represents the complete payload for creating a new presentation."""
    project: str
    display_name: str
    powerpoint_file: str
    excel_file: str
    background_type: str
    background_name: str
    page_number: int
    presentation_type: str
    user_name: str
    bsr_display_name: str
    mobile_link_bsr: Optional[str] = None
    participant_vote: int
    is_wide_ppt: int
    is_aws_email: int
    details: List[DetailItem]