"""Service for Global Analytics Statistics.

This service handles the Statistics tab functionality, which generates analytics
reports for ALL projects (not individual project reports).

IMPORTANT NOTES:
- This is DIFFERENT from nw_reports_service.py (which handles individual project reports)
- The original SPs (NW_ProjectAnalytics, etc.) depend on external databases (DayMaster, BNRS)
  which are not available, so we use alternative queries directly from BI_GUIDELINES
- We calculate percentages and aggregations in Python instead of relying on the SPs
"""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from app.config.db import get_connection_scope
from app.models.analytics_models import (
    BSRProjectAnalytics,
    BSRRegionAnalytics,
    NWProjectAnalytics,
    NWRegionAnalytics,
)
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class AnalyticsService:
    """Service for retrieving global analytics statistics for ALL projects."""

    def _calculate_percentages(
        self,
        positive: int,
        neutral: int,
        negative: int
    ) -> Tuple[str, str, str, str]:
        """Calculate percentage strings for vote counts.

        Args:
            positive: Number of positive votes
            neutral: Number of neutral votes
            negative: Number of negative votes

        Returns:
            Tuple of (percent_retained, percent_positive, percent_neutral, percent_negative)
            All as formatted percentage strings (e.g., "85.50%")
        """
        total = positive + neutral + negative

        if total == 0:
            return "0.00%", "0.00%", "0.00%", "0.00%"

        # Calculate percentages
        percent_retained = ((positive + neutral) / total) * 100
        percent_positive = (positive / total) * 100
        percent_neutral = (neutral / total) * 100
        percent_negative = (negative / total) * 100

        # Format to 2 decimal places with % symbol
        return (
            f"{percent_retained:.2f}%",
            f"{percent_positive:.2f}%",
            f"{percent_neutral:.2f}%",
            f"{percent_negative:.2f}%",
        )

    def _get_new_names_created(self, display_name: str) -> int:
        """Get the number of newly created names for a project.

        This requires calling two stored procedures in sequence:
        1. nw_GetPresentationId to get the presentation ID from display name
        2. nw_wdGetResults with SummaryType="ByTheNumbers" to get the count

        NOTE: Many projects may not have this data available, so we gracefully
        return 0 for any errors.

        Args:
            display_name: Display name of the presentation

        Returns:
            Number of newly created names, or 0 if not found
        """
        # Since this is optional data and many SPs may not return results,
        # we just return 0 for now to avoid slowing down the main query
        # TODO: Implement this once we have a more reliable way to get the data
        return 0

    async def get_nw_projects_analytics(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None
    ) -> List[NWProjectAnalytics]:
        """Get analytics for ALL NW projects.

        This method queries the database for all NW projects and calculates
        their statistics (vote percentages, new names created, etc.).

        Since the original SP (NW_ProjectAnalytics) depends on external databases
        that are not available, we use an alternative query directly from BI_GUIDELINES.

        Args:
            start_date: Optional start date for filtering (YYYY-MM-DD)
            end_date: Optional end date for filtering (YYYY-MM-DD)

        Returns:
            List of NWProjectAnalytics objects
        """
        logger.info(
            "Fetching analytics for ALL NW projects (start_date=%s, end_date=%s)",
            start_date,
            end_date
        )

        try:
            with get_connection_scope(timeout=60) as cursor:
                # Alternative query that doesn't depend on external databases
                query = """
                SELECT
                    mas.Project,
                    mas.DisplayName,
                    mas.UploadedBy,
                    mas.PresentationId,
                    mas.PresentationStatus,
                    COUNT(CASE WHEN det.NameRanking = 'positive' THEN 1 END) AS PositiveCount,
                    COUNT(CASE WHEN det.NameRanking = 'negative' THEN 1 END) AS NegativeCount,
                    COUNT(CASE WHEN det.NameRanking = 'neutral' THEN 1 END) AS NeutralCount
                FROM [BI_GUIDELINES].[dbo].nw_master mas
                LEFT JOIN [BI_GUIDELINES].[dbo].nw_Details det ON mas.PresentationId = det.PresentationId
                WHERE mas.PresentationStatus = 'OPEN'
                """

                # Add date filters if provided
                params = []
                if start_date:
                    query += " AND mas.UploadedDate >= ?"
                    params.append(start_date)
                if end_date:
                    query += " AND mas.UploadedDate <= ?"
                    params.append(end_date)

                query += """
                GROUP BY mas.Project, mas.DisplayName, mas.UploadedBy, mas.PresentationId, mas.PresentationStatus
                ORDER BY mas.UploadedBy, mas.Project
                """

                cursor.execute(query, params)
                rows = cursor.fetchall()
                columns = [column[0] for column in cursor.description]

                logger.info("Retrieved %d NW project records from database", len(rows))

                analytics_list = []
                for row in rows:
                    row_dict = dict(zip(columns, row))

                    project = row_dict.get("Project", "")
                    display_name = row_dict.get("DisplayName", "")
                    uploaded_by = row_dict.get("UploadedBy", "")
                    status = row_dict.get("PresentationStatus", "")

                    # Get vote counts
                    positive = row_dict.get("PositiveCount", 0) or 0
                    negative = row_dict.get("NegativeCount", 0) or 0
                    neutral = row_dict.get("NeutralCount", 0) or 0

                    # Skip if no votes at all (based on VB.NET logic)
                    # VB.NET: if Positive + Neutral == Positive + Neutral + Neg: continue
                    # This means: skip if negative is 0 and positive+neutral is 0
                    total_votes = positive + neutral + negative
                    if total_votes == 0:
                        logger.debug("Skipping project %s - no votes", display_name)
                        continue

                    # Calculate percentages
                    (
                        percent_retained,
                        percent_positive,
                        percent_neutral,
                        percent_negative,
                    ) = self._calculate_percentages(positive, neutral, negative)

                    # Get new names created (requires calling additional SPs)
                    # Note: This can be slow for many projects, consider caching or parallel execution
                    new_names_created = self._get_new_names_created(display_name)

                    # Determine active status
                    active = "Yes" if status == "OPEN" else "No"

                    analytics = NWProjectAnalytics(
                        region="",  # Not available without external DB
                        active_lead=uploaded_by,  # Use UploadedBy as fallback
                        project_name=project,
                        display_name=display_name,
                        client_name="",  # Not available without external DB (DayMaster)
                        percent_retained=percent_retained,
                        percent_positive=percent_positive,
                        percent_neutral=percent_neutral,
                        percent_negative=percent_negative,
                        names_created=new_names_created,
                    )

                    analytics_list.append(analytics)

                logger.info(
                    "Processed %d NW project analytics (filtered from %d total)",
                    len(analytics_list),
                    len(rows)
                )

                return analytics_list

        except Exception as e:
            logger.error(
                "Error fetching NW projects analytics: %s",
                str(e),
                exc_info=True
            )
            raise

    async def get_nw_regions_analytics(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None
    ) -> List[NWRegionAnalytics]:
        """Get aggregated analytics by region for NW projects.

        Since we don't have access to region data from external databases,
        this method aggregates by UploadedBy (active lead) instead.

        Args:
            start_date: Optional start date for filtering (YYYY-MM-DD)
            end_date: Optional end date for filtering (YYYY-MM-DD)

        Returns:
            List of NWRegionAnalytics objects
        """
        logger.info(
            "Fetching region-specific analytics for NW projects (start_date=%s, end_date=%s)",
            start_date,
            end_date
        )

        try:
            # First get all projects analytics with date filters
            projects = await self.get_nw_projects_analytics(start_date, end_date)

            # Aggregate by active_lead (since we don't have region data)
            aggregates: Dict[str, Dict] = {}

            for project in projects:
                lead = project.active_lead or "Unknown"

                if lead not in aggregates:
                    aggregates[lead] = {
                        "projects_count": 0,
                        "total_positive": 0,
                        "total_neutral": 0,
                        "total_negative": 0,
                        "total_names": 0,
                    }

                agg = aggregates[lead]
                agg["projects_count"] += 1
                agg["total_names"] += project.names_created

                # Parse percentages back to counts for aggregation
                # This is an approximation since we don't have the raw counts
                # In production, you'd want to store the raw counts
                agg["total_positive"] += float(project.percent_positive.rstrip('%'))
                agg["total_neutral"] += float(project.percent_neutral.rstrip('%'))
                agg["total_negative"] += float(project.percent_negative.rstrip('%'))

            # Convert aggregates to NWRegionAnalytics objects
            region_analytics = []
            for lead, agg in aggregates.items():
                projects_count = agg["projects_count"]

                # Calculate average percentages
                if projects_count > 0:
                    avg_positive = agg["total_positive"] / projects_count
                    avg_neutral = agg["total_neutral"] / projects_count
                    avg_negative = agg["total_negative"] / projects_count
                    avg_retained = avg_positive + avg_neutral
                    avg_names = agg["total_names"] / projects_count
                else:
                    avg_positive = avg_neutral = avg_negative = avg_retained = avg_names = 0

                region = NWRegionAnalytics(
                    region=lead,  # Using lead as "region" since region data not available
                    active_lead=lead,
                    projects_to_date=projects_count,
                    percent_retained=f"{avg_retained:.2f}%",
                    percent_positive=f"{avg_positive:.2f}%",
                    percent_neutral=f"{avg_neutral:.2f}%",
                    percent_negative=f"{avg_negative:.2f}%",
                    avg_names_per_project=avg_names,
                )

                region_analytics.append(region)

            logger.info("Processed %d region aggregates", len(region_analytics))
            return region_analytics

        except Exception as e:
            logger.error(
                "Error fetching NW regions analytics: %s",
                str(e),
                exc_info=True
            )
            raise

    async def get_bsr_projects_analytics(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None
    ) -> List[BSRProjectAnalytics]:
        """Get analytics for ALL BSR projects.

        This method queries the database for all BSR projects and calculates
        their statistics (PC count, mobile count, totals).

        Since the original SP depends on external databases, we use an
        alternative query from BI_GUIDELINES tables.

        Args:
            start_date: Optional start date for filtering (YYYY-MM-DD)
            end_date: Optional end date for filtering (YYYY-MM-DD)

        Returns:
            List of BSRProjectAnalytics objects
        """
        logger.info(
            "Fetching analytics for ALL BSR projects (start_date=%s, end_date=%s)",
            start_date,
            end_date
        )

        try:
            with get_connection_scope(timeout=60) as cursor:
                # Alternative query for BSR analytics
                # IsMobile is in bsr_ProjectConcepts, Source is in BSR_CONCEPTNAMES
                query = """
                SELECT
                    bm.Project,
                    bm.DisplayName,
                    bm.UploadedBy,
                    bm.PresentationId,
                    bm.PresentationStatus,
                    COUNT(CASE WHEN pc.IsMobile = 1 AND cn.Source <> 'Moderator' THEN 1 END) AS MobileCount,
                    COUNT(CASE WHEN pc.IsMobile = 0 OR (pc.IsMobile = 1 AND cn.Source = 'Moderator') THEN 1 END) AS PCCount
                FROM [BI_GUIDELINES].[dbo].bsr_Master bm
                LEFT JOIN [BI_GUIDELINES].[dbo].bsr_ProjectConcepts pc ON bm.PresentationId = pc.projectid
                LEFT JOIN [BI_GUIDELINES].[dbo].BSR_CONCEPTNAMES cn ON pc.conceptid = cn.ConceptId AND pc.projectid = cn.ProjectId
                WHERE bm.PresentationStatus = 'OPEN'
                """

                # Add date filters if provided
                params = []
                if start_date:
                    query += " AND bm.UploadedDate >= ?"
                    params.append(start_date)
                if end_date:
                    query += " AND bm.UploadedDate <= ?"
                    params.append(end_date)

                query += """
                GROUP BY bm.Project, bm.DisplayName, bm.UploadedBy, bm.PresentationId, bm.PresentationStatus
                HAVING COUNT(cn.NameId) > 0
                ORDER BY bm.UploadedBy, bm.Project
                """

                cursor.execute(query, params)
                rows = cursor.fetchall()
                columns = [column[0] for column in cursor.description]

                logger.info("Retrieved %d BSR project records from database", len(rows))

                analytics_list = []
                for row in rows:
                    row_dict = dict(zip(columns, row))

                    project = row_dict.get("Project", "")
                    display_name = row_dict.get("DisplayName", "")
                    uploaded_by = row_dict.get("UploadedBy", "")
                    status = row_dict.get("PresentationStatus", "")

                    # Get counts
                    mobile_count = row_dict.get("MobileCount", 0) or 0
                    pc_count = row_dict.get("PCCount", 0) or 0
                    total = mobile_count + pc_count

                    # Skip if no data
                    if total == 0:
                        logger.debug("Skipping BSR project %s - no data", display_name)
                        continue

                    analytics = BSRProjectAnalytics(
                        region="",  # Not available without external DB
                        active_lead=uploaded_by,  # Use UploadedBy as fallback
                        project_name=project,
                        display_name=display_name,
                        client_name="",  # Not available without external DB (DayMaster)
                        names_created_web=pc_count,
                        names_created_mobile=mobile_count,
                        total_names_created=total,
                    )

                    analytics_list.append(analytics)

                logger.info(
                    "Processed %d BSR project analytics",
                    len(analytics_list)
                )

                return analytics_list

        except Exception as e:
            logger.error(
                "Error fetching BSR projects analytics: %s",
                str(e),
                exc_info=True
            )
            raise

    async def get_bsr_regions_analytics(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None
    ) -> List[BSRRegionAnalytics]:
        """Get aggregated analytics by region for BSR projects.

        Since we don't have access to region data from external databases,
        this method aggregates by UploadedBy (active lead) instead.

        Args:
            start_date: Optional start date for filtering (YYYY-MM-DD)
            end_date: Optional end date for filtering (YYYY-MM-DD)

        Returns:
            List of BSRRegionAnalytics objects
        """
        logger.info(
            "Fetching region-specific analytics for BSR projects (start_date=%s, end_date=%s)",
            start_date,
            end_date
        )

        try:
            # First get all projects analytics with date filters
            projects = await self.get_bsr_projects_analytics(start_date, end_date)

            # Aggregate by active_lead
            aggregates: Dict[str, Dict] = {}

            for project in projects:
                lead = project.active_lead or "Unknown"

                if lead not in aggregates:
                    aggregates[lead] = {
                        "projects_count": 0,
                        "total_pc_count": 0,
                        "total_mobile_count": 0,
                    }

                agg = aggregates[lead]
                agg["projects_count"] += 1
                agg["total_pc_count"] += project.names_created_web
                agg["total_mobile_count"] += project.names_created_mobile

            # Convert aggregates to BSRRegionAnalytics objects
            region_analytics = []
            for lead, agg in aggregates.items():
                total = agg["total_pc_count"] + agg["total_mobile_count"]

                region = BSRRegionAnalytics(
                    region=lead,  # Using lead as "region"
                    active_lead=lead,
                    projects_to_date=agg["projects_count"],
                    pc_count=agg["total_pc_count"],
                    mobile_count=agg["total_mobile_count"],
                    total=total,
                )

                region_analytics.append(region)

            logger.info("Processed %d BSR region aggregates", len(region_analytics))
            return region_analytics

        except Exception as e:
            logger.error(
                "Error fetching BSR regions analytics: %s",
                str(e),
                exc_info=True
            )
            raise


# Global service instance
analytics_service = AnalyticsService()
