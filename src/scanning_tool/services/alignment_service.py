"""Auto-alignment background service."""

from scanning_tool.services.base_service import BaseService

class AlignmentService(BaseService):
    """Manages continuous alignment calculations for screen anchor points."""
    
    def _on_start(self) -> None:
        self.logger.info("Starting auto-alignment service.")
        # Hook into core/auto_alignment.py functions
        from scanning_tool.core.auto_alignment import perform_auto_alignment
        
    def _on_stop(self) -> None:
        self.logger.info("Stopping auto-alignment tracking.")
