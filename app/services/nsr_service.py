"""Service for NSR (Name Selection Research) configuration management."""

from __future__ import annotations

from typing import List, Dict, Optional

from app.config.db import get_connection_scope
from app.models.nsr_models import NSRRuleConfig
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class NSRService:
    """Service for managing NSR rule configurations."""

    def get_project_config(self, project_name: str) -> List[NSRRuleConfig]:
        """
        Get all rule configurations for an NSR project.

        Args:
            project_name: Display name of the NSR project

        Returns:
            List of NSRRuleConfig objects with rule_id and is_on values
        """
        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nsr_getProjectConfig](?)}",
                    (project_name,)
                )

                columns = [c[0] for c in cursor.description] if cursor.description else []
                rows = cursor.fetchall()

                rules: List[NSRRuleConfig] = []
                for row in rows:
                    row_dict = dict(zip(columns, row))
                    rules.append(
                        NSRRuleConfig(
                            rule_id=row_dict.get("ruleid") or row_dict.get("rule_id") or row_dict.get("RuleId") or 0,
                            is_on=row_dict.get("isOn") or row_dict.get("ison") or row_dict.get("is_on") or row_dict.get("IsOn") or 0,
                        )
                    )

                logger.info("Retrieved %d NSR rules for project '%s'", len(rules), project_name)
                return rules

        except Exception as e:
            logger.error(
                "Error fetching NSR config for project '%s': %s",
                project_name,
                str(e),
                exc_info=True
            )
            return []

    def check_rule_exists(self, project_name: str, rule_id: int) -> bool:
        """
        Check if a rule entry exists for a project.

        Args:
            project_name: Display name of the NSR project
            rule_id: Rule ID to check (101-120)

        Returns:
            True if rule exists (needs UPDATE), False if not (needs INSERT)
        """
        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nsr_CheckTableEntry](?, ?)}",
                    (project_name, rule_id)
                )
                result = cursor.fetchone()

                if result:
                    exists = result[0] == 1
                    logger.info(
                        "Rule %d for project '%s': %s",
                        rule_id,
                        project_name,
                        "EXISTS" if exists else "NOT EXISTS"
                    )
                    return exists

                return False

        except Exception as e:
            logger.error(
                "Error checking NSR rule existence (project='%s', rule=%d): %s",
                project_name,
                rule_id,
                str(e),
                exc_info=True
            )
            return False

    def insert_rule(self, project_name: str, rule_id: int, is_on: int) -> bool:
        """
        Insert a new NSR rule configuration.

        Args:
            project_name: Display name of the NSR project
            rule_id: Rule ID (101-120)
            is_on: Rule state (0=OFF, 1=ON)

        Returns:
            True if successful, False otherwise
        """
        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nsr_InsertPresentationDetail](?, ?, ?)}",
                    (project_name, rule_id, is_on)
                )
                # Commit is handled by the context manager in db.py

                logger.info(
                    "Inserted NSR rule: project='%s', rule_id=%d, is_on=%d",
                    project_name,
                    rule_id,
                    is_on
                )
                return True

        except Exception as e:
            logger.error(
                "Error inserting NSR rule (project='%s', rule=%d, is_on=%d): %s",
                project_name,
                rule_id,
                is_on,
                str(e),
                exc_info=True
            )
            return False

    def update_rule(self, project_name: str, rule_id: int, is_on: int) -> bool:
        """
        Update an existing NSR rule configuration.

        Args:
            project_name: Display name of the NSR project
            rule_id: Rule ID (101-120)
            is_on: Rule state (0=OFF, 1=ON)

        Returns:
            True if successful, False otherwise
        """
        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nsr_UpdatePresentationDetail](?, ?, ?)}",
                    (project_name, rule_id, is_on)
                )
                # Commit is handled by the context manager in db.py

                logger.info(
                    "Updated NSR rule: project='%s', rule_id=%d, is_on=%d",
                    project_name,
                    rule_id,
                    is_on
                )
                return True

        except Exception as e:
            logger.error(
                "Error updating NSR rule (project='%s', rule=%d, is_on=%d): %s",
                project_name,
                rule_id,
                is_on,
                str(e),
                exc_info=True
            )
            return False

    def upsert_rule(self, project_name: str, rule_id: int, is_on: int) -> tuple[bool, str]:
        """
        Insert or update an NSR rule configuration (upsert operation).

        Args:
            project_name: Display name of the NSR project
            rule_id: Rule ID (101-120)
            is_on: Rule state (0=OFF, 1=ON)

        Returns:
            Tuple of (success: bool, operation: str) where operation is 'insert' or 'update'
        """
        # Check if rule exists
        exists = self.check_rule_exists(project_name, rule_id)

        if exists:
            # Update existing rule
            success = self.update_rule(project_name, rule_id, is_on)
            return success, "update"
        else:
            # Insert new rule
            success = self.insert_rule(project_name, rule_id, is_on)
            return success, "insert"

    def get_active_nsr_projects(self) -> List[str]:
        """
        Get list of active NSR project display names.

        Returns:
            List of NSR project display names
        """
        try:
            with get_connection_scope(timeout=30) as cursor:
                # Try using stored procedure first
                try:
                    cursor.execute("{CALL [BI_GUIDELINES].[dbo].[nsr_ActivePresentations]}")
                except Exception:
                    # Fallback to direct query if SP doesn't exist
                    cursor.execute(
                        """
                        SELECT DISTINCT displayname
                        FROM [BI_GUIDELINES].[dbo].[bsr_Master]
                        WHERE PresentationType IN ('NSR', 'NSR-Japan')
                          AND status = 'ACTIVE'
                        ORDER BY displayname ASC
                        """
                    )

                rows = cursor.fetchall()
                projects = [row[0] for row in rows if row and row[0]]

                logger.info("Retrieved %d active NSR projects", len(projects))
                return projects

        except Exception as e:
            logger.error(
                "Error fetching active NSR projects: %s",
                str(e),
                exc_info=True
            )
            return []

    def initialize_all_rules(self, project_name: str) -> tuple[bool, int]:
        """
        Initialize all NSR rules (101-120, except 107) for a project in ON state.

        Args:
            project_name: Display name of the NSR project

        Returns:
            Tuple of (success: bool, rules_created: int)
        """
        # All rule IDs except 107 (which doesn't exist)
        all_rule_ids = [101, 102, 103, 104, 105, 106, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120]
        rules_created = 0

        try:
            for rule_id in all_rule_ids:
                # Insert each rule as ON (1)
                success = self.insert_rule(project_name, rule_id, 1)
                if success:
                    rules_created += 1
                else:
                    logger.warning(
                        "Failed to initialize rule %d for project '%s'",
                        rule_id,
                        project_name
                    )

            logger.info(
                "Initialized %d/%d NSR rules for project '%s'",
                rules_created,
                len(all_rule_ids),
                project_name
            )
            return True, rules_created

        except Exception as e:
            logger.error(
                "Error initializing NSR rules for project '%s': %s",
                project_name,
                str(e),
                exc_info=True
            )
            return False, rules_created


# Global instance
nsr_service = NSRService()
