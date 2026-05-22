# =============================================================================
#  version.py  —  Single source of truth for BreakDown versioning
#  Auto-updated by build.bat via bump_version.py
# =============================================================================

__version__      = "1.2.5"        # MAJOR.MINOR.PATCH
__build__        = "20260513"      # YYYYMMDD  — updated by build.bat
__release_name__ = "Initial Release"
__channel__      = "stable"        # stable | beta | dev
__author__       = "BreakDown Team"
__app_name__     = "BreakDown"


# ---------------------------------------------------------------------------
#  Helpers
# ---------------------------------------------------------------------------

def get_version_string() -> str:
    """Returns e.g.  'BreakDown v1.0.0 (build 20250317)'"""
    return f"{__app_name__} v{__version__} (build {__build__})"


def get_full_version_info() -> dict:
    """Returns all version metadata as a dict."""
    return {
        "app":     __app_name__,
        "version": __version__,
        "build":   __build__,
        "release": __release_name__,
        "channel": __channel__,
        "author":  __author__,
    }


def parse_version(version_str: str) -> tuple:
    """Parse '1.2.3' → (1, 2, 3)"""
    try:
        parts = version_str.strip().split(".")
        return tuple(int(p) for p in parts)
    except (ValueError, AttributeError):
        return (0, 0, 0)


def is_newer(other_version: str) -> bool:
    """Return True if other_version is newer than current __version__."""
    return parse_version(other_version) > parse_version(__version__)
