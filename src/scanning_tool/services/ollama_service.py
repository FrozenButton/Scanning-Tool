"""Ollama lifecycle management service."""

from scanning_tool.services.base_service import BaseService

class OllamaService(BaseService):
    """Manages the Ollama process and model availability."""
    
    def _on_start(self) -> None:
        self.logger.info("Initializing Ollama backend...")
        # Simulation of integration with scanning_tool.ollama.installer
        from scanning_tool.ollama.installer import installer
        # Assuming installer setup logic continues here...
        
    def _on_stop(self) -> None:
        self.logger.info("Shutting down Ollama integration.")
