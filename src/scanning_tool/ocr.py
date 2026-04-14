"""OCR via Ollama vision models."""

import io
import logging
from typing import Optional
import ollama

from PIL import Image

from scanning_tool.ollama_service import get_ollama_client, get_ollama_model

logger = logging.getLogger("scanning_tool")


def ocr_with_ollama(pil_img: Image.Image, model: Optional[str] = None) -> str:
    """Send an image to Ollama for OCR and return the extracted text."""
    if model is None:
        model = get_ollama_model()
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    img_bytes = buf.getvalue()
    client: ollama.Client = get_ollama_client()
    # Model-specific prompt selection
    prompt = "Extract the numeric code shown in this image. Only return the code, no extra words."
    if model:
        m = model.lower()
        if m.startswith("moondream"):
            prompt = (
                "Only output the numbers you see in this image. Do not describe the image. "
                "If there are no numbers, output nothing."
            )
        elif m.startswith("granite"):
            prompt = "Read all numbers in this image. Only return the numbers."
        elif m.startswith("deepseek-ocr"):
            prompt = "Read all text in this image. Only return the numbers."
        elif m.startswith("smolvlm"):
            prompt = "Only output the numbers you see in this image. Do not describe the image. If there are no numbers, output nothing."
        elif m.startswith("bakllava"):
            prompt = "Extract all numbers from this image. Only output the numbers."
        elif m.startswith("llava"):
            prompt = "What numbers are visible in this image? Only output the numbers."
    logger.info(f"Using OCR prompt for model '{model}': {prompt}")
    try:
        response: ollama.ChatResponse = client.chat(
            model=model,
            messages=[{
                "role": "user",
                "content": prompt,
                "images": [img_bytes],
            }],
        )
        return response["message"]["content"].strip()
    except Exception as e:
        logger.error(f"Ollama OCR error: {e}")
        return ""
