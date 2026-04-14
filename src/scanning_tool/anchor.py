"""Backward-compatible anchor module forwarding to the core anchor package."""

from scanning_tool.core.anchor import AnchorRegionTracker
from scanning_tool.core.auto_alignment import perform_auto_alignment

__all__ = ["AnchorRegionTracker", "perform_auto_alignment"]
