"""Secure configuration loader infrastructure."""

import os
from typing import Dict, Any

class ConfigLoader:
    """Handles parsing and validating env vars, yaml, or json configs."""
    
    @staticmethod
    def load(file_path: str) -> Dict[str, Any]:
        """Loads configuration from environment or file."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Configuration file {file_path} not found.")
        
        # In a real scenario, we would parse JSON/YAML here
        return {"loaded_from": file_path}
