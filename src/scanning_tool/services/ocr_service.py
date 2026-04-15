"""OCR Service implementation."""

from typing import Any
from scanning_tool.interfaces.services import IOcrService

class OcrService(IOcrService):
    """Concrete implementation of OCR capabilities."""
    
    def extract_text(self, image: Any) -> str:
        """Extract text from the given image."""
        # Simulated extraction routing to existing ollama functions
        from scanning_tool.ocr import ocr_with_ollama
        return ocr_with_ollama(image)
