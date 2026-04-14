import os

from scanning_tool.ollama.host import get_ollama_host, get_ollama_model, is_local_ollama_host, sanitize_ollama_host


def test_sanitize_ollama_host_adds_http_when_missing():
    assert sanitize_ollama_host("127.0.0.1:11434") == "http://127.0.0.1:11434"
    assert sanitize_ollama_host("http://127.0.0.1:11434") == "http://127.0.0.1:11434"
    assert sanitize_ollama_host("") == ""


def test_is_local_ollama_host_identifies_local_hosts():
    assert is_local_ollama_host("localhost") is True
    assert is_local_ollama_host("127.0.0.1") is True
    assert is_local_ollama_host("0.0.0.0") is True
    assert is_local_ollama_host("example.com") is False


def test_get_ollama_host_prefers_env_over_config(monkeypatch):
    monkeypatch.setenv("OLLAMA_HOST", "example.com")
    assert get_ollama_host() == "http://example.com"


def test_get_ollama_model_prefers_env_over_config(monkeypatch):
    monkeypatch.setenv("OLLAMA_MODEL", "test-model")
    assert get_ollama_model() == "test-model"
    monkeypatch.delenv("OLLAMA_MODEL", raising=False)
