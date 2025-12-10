"""
OpenAI Client - Wrapper for OpenAI API interactions.
"""

import os
from typing import List, Dict, Any, Optional
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()


class OpenAIClient:
    """Handles interactions with OpenAI API."""

    def __init__(self, api_key: Optional[str] = None,
                 model: str = "gpt-4-turbo-preview",
                 base_url: Optional[str] = None):
        """
        Initialize the OpenAI client.

        Args:
            api_key: OpenAI API key (defaults to OPENAI_API_KEY env var)
            model: Model to use for completions
            base_url: Optional base URL for API (for OpenAI-compatible APIs)
        """
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        self.model = model or os.getenv("OPENAI_MODEL", "gpt-4-turbo-preview")

        if not self.api_key:
            raise ValueError(
                "OpenAI API key must be provided or set in OPENAI_API_KEY environment variable"
            )

        client_kwargs = {"api_key": self.api_key}
        if base_url:
            client_kwargs["base_url"] = base_url
        elif os.getenv("OPENAI_BASE_URL"):
            client_kwargs["base_url"] = os.getenv("OPENAI_BASE_URL")

        self.client = OpenAI(**client_kwargs)

    def chat_completion(self, messages: List[Dict[str, str]],
                       temperature: float = 0.7,
                       max_tokens: Optional[int] = None,
                       response_format: Optional[Dict] = None) -> str:
        """
        Get a chat completion from OpenAI.

        Args:
            messages: List of message dictionaries with 'role' and 'content'
            temperature: Sampling temperature (0-2)
            max_tokens: Maximum tokens in response
            response_format: Optional response format specification

        Returns:
            The assistant's response content
        """
        kwargs = {
            "model": self.model,
            "messages": messages,
            "temperature": temperature,
        }

        if max_tokens:
            kwargs["max_tokens"] = max_tokens

        if response_format:
            kwargs["response_format"] = response_format

        response = self.client.chat.completions.create(**kwargs)
        return response.choices[0].message.content

    def generate_json(self, messages: List[Dict[str, str]],
                     temperature: float = 0.7) -> str:
        """
        Generate a JSON response from OpenAI.

        Args:
            messages: List of message dictionaries
            temperature: Sampling temperature

        Returns:
            JSON string response
        """
        return self.chat_completion(
            messages=messages,
            temperature=temperature,
            response_format={"type": "json_object"}
        )

    def create_system_message(self, content: str) -> Dict[str, str]:
        """Create a system message."""
        return {"role": "system", "content": content}

    def create_user_message(self, content: str) -> Dict[str, str]:
        """Create a user message."""
        return {"role": "user", "content": content}

    def create_assistant_message(self, content: str) -> Dict[str, str]:
        """Create an assistant message."""
        return {"role": "assistant", "content": content}
