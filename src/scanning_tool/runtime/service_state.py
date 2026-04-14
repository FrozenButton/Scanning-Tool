from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Optional, Pattern
import re
import subprocess

@dataclass
class ServiceState:
    ollama_client: Any = None
    ollama_client_host: str = ""
    ollama_server_process: Optional[subprocess.Popen] = field(default=None, repr=False)
    code_re: Pattern[str] = field(default_factory=lambda: re.compile(
        r"(?:[A-Za-z]?-?\d[\d,\.]{1,10}|\d{2,10})",
        re.IGNORECASE,
    ))
    host_scheme_re: Pattern[str] = field(default_factory=lambda: re.compile(r"^[a-zA-Z][a-zA-Z0-9+.-]*://"))
    rock_data: Dict[str, Any] = field(default_factory=dict)
    deposit_tables: Dict[str, Any] = field(default_factory=dict)
    gui_status_callback: Optional[Callable[[str], None]] = None
