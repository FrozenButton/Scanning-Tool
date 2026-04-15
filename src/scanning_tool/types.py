"""Shared type definitions across the scanning tool ecosystem."""

from typing import Any, Callable, Dict, List, Optional, Tuple, Union

# Common structural types
ImageArray = Any  # Typically a numpy array
Coordinate = Tuple[int, int]
BoundingBox = Tuple[int, int, int, int]
ConfigDict = Dict[str, Any]

# UI related types
EventHandler = Callable[[Any], None]
