"""
Manage SharePoint document libraries and their permissions through Entra ID security groups.

SharePointLibraryManager creates document libraries, gives each its own permissions, grants access
through Entra ID security groups it creates and keeps their membership in sync, and keeps the site's
navigation and sharing settings in line. Every operation takes a pandas DataFrame describing the
desired state and returns an action-log DataFrame describing what it did (or, in a dry run, would do).
"""
from ._common import (
    ACTION_LOG_COLUMNS,
    ALL_GROUP_SUFFIXES,
    AUTO_GROUP_SUFFIXES,
    FATAL_ACTIONS,
    MANUAL_GROUP_SUFFIXES,
    VALID_ROLES,
    WARNING_ACTIONS,
)
from ._groups import GroupResolutionError, UserResolutionError, build_group_name, sanitise_for_group_name
from .manager import SharePointLibraryManager

__all__ = [
    "SharePointLibraryManager",
    "UserResolutionError",
    "GroupResolutionError",
    "build_group_name",
    "sanitise_for_group_name",
    "ACTION_LOG_COLUMNS",
    "ALL_GROUP_SUFFIXES",
    "AUTO_GROUP_SUFFIXES",
    "MANUAL_GROUP_SUFFIXES",
    "VALID_ROLES",
    "FATAL_ACTIONS",
    "WARNING_ACTIONS",
]
