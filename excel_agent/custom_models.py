"""Custom model wrappers for GLM and other OpenAI-compatible APIs."""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from typing import Any, Literal, overload

from pydantic_ai.messages import ModelMessage, ModelResponse, TextPart
from pydantic_ai.models import Model, ModelRequestParameters
from pydantic_ai.models.openai import OpenAIChatModel, OpenAIChatModelSettings
from pydantic_ai.settings import ModelSettings

try:
    from openai import AsyncStream, NOT_GIVEN, APIStatusError
    from openai.types import chat
    from openai.types.chat import ChatCompletionChunk
except ImportError:
    raise ImportError("openai package is required for custom model support")

logger = logging.getLogger(__name__)


@dataclass(init=False)
class CustomOpenAIModel(OpenAIChatModel):
    """
    Custom OpenAI-compatible model for providers like GLM, DeepSeek, etc.
    
    This model allows you to use any OpenAI-compatible API endpoint
    with custom base URLs and authentication.
    
    Usage:
        model = CustomOpenAIModel(
            model_name='glm-4.6',
            base_url='https://api.z.ai/api/paas/v4',
            api_key='your-api-key'
        )
        agent = Agent(model, ...)
    """

    def __init__(
        self,
        model_name: str,
        *,
        base_url: str,
        api_key: str,
        profile=None,
        settings: ModelSettings | None = None,
    ):
        """Initialize a custom OpenAI-compatible model.
        
        Args:
            model_name: The name of the model (e.g., 'glm-4.6')
            base_url: Custom API base URL
            api_key: API key for the service
            profile: Optional model profile
            settings: Optional default model settings
        """
        # Temporarily set environment variables for the parent class
        original_openai_key = os.environ.get("OPENAI_API_KEY")
        original_base_url = os.environ.get("OPENAI_BASE_URL")
        
        try:
            os.environ["OPENAI_API_KEY"] = api_key
            os.environ["OPENAI_BASE_URL"] = base_url
            
            # Initialize parent with environment variables
            super().__init__(
                model_name, 
                provider="openai",  # Use openai provider for compatibility
                profile=profile, 
                settings=settings
            )
            
        finally:
            # Restore original environment variables
            if original_openai_key is not None:
                os.environ["OPENAI_API_KEY"] = original_openai_key
            else:
                os.environ.pop("OPENAI_API_KEY", None)
                
            if original_base_url is not None:
                os.environ["OPENAI_BASE_URL"] = original_base_url
            else:
                os.environ.pop("OPENAI_BASE_URL", None)

    @overload
    async def _completions_create(
        self,
        messages: list[ModelMessage],
        stream: Literal[True],
        model_settings: OpenAIChatModelSettings,
        model_request_parameters: ModelRequestParameters,
    ) -> AsyncStream[ChatCompletionChunk]: ...

    @overload
    async def _completions_create(
        self,
        messages: list[ModelMessage],
        stream: Literal[False],
        model_settings: OpenAIChatModelSettings,
        model_request_parameters: ModelRequestParameters,
    ) -> chat.ChatCompletion: ...

    async def _completions_create(
        self,
        messages: list[ModelMessage],
        stream: bool,
        model_settings: OpenAIChatModelSettings,
        model_request_parameters: ModelRequestParameters,
    ) -> chat.ChatCompletion | AsyncStream[ChatCompletionChunk]:
        """Create completion using custom API endpoint."""
        from pydantic_ai.profiles.openai import OpenAIModelProfile
        
        # Map messages to OpenAI format
        openai_messages = await self._map_messages(messages)
        
        # Get model profile
        model_profile = self.profile
        if isinstance(model_profile, OpenAIModelProfile):
            supports_tool_choice_required = (
                model_profile.openai_supports_tool_choice_required
            )
            unsupported_settings = (
                model_profile.openai_unsupported_model_settings
            )
        else:
            supports_tool_choice_required = True
            unsupported_settings = set()

        # Filter out unsupported settings
        filtered_settings = {
            k: v for k, v in model_settings.items() 
            if k not in unsupported_settings
        }
        
        # Create the completion request
        response = await self.client.chat.completions.create(
            model=self.model_name,
            messages=openai_messages,
            stream=stream,
            tools=filtered_settings.get("tools", NOT_GIVEN),
            tool_choice=filtered_settings.get("tool_choice", NOT_GIVEN),
            max_tokens=filtered_settings.get("max_tokens", NOT_GIVEN),
            temperature=filtered_settings.get("temperature", NOT_GIVEN),
            top_p=filtered_settings.get("top_p", NOT_GIVEN),
            stop=filtered_settings.get("stop", NOT_GIVEN),
            presence_penalty=filtered_settings.get("presence_penalty", NOT_GIVEN),
            frequency_penalty=filtered_settings.get("frequency_penalty", NOT_GIVEN),
            seed=filtered_settings.get("seed", NOT_GIVEN),
            response_format=filtered_settings.get("response_format", NOT_GIVEN),
            user=filtered_settings.get("user", NOT_GIVEN),
        )

        return response


def create_custom_model(
    model_name: str, 
    base_url: str, 
    api_key: str,
    **kwargs
) -> CustomOpenAIModel:
    """Create a custom OpenAI-compatible model.
    
    Args:
        model_name: Model name (e.g., 'glm-4.6')
        base_url: API base URL (e.g., 'https://api.z.ai/api/paas/v4')
        api_key: API key for authentication
        **kwargs: Additional model settings
        
    Returns:
        CustomOpenAIModel instance
        
    Example:
        model = create_custom_model(
            model_name='glm-4.6',
            base_url='https://api.z.ai/api/paas/v4',
            api_key='18f65090a96a425898a8398a5c4518ce.DDtUvTTnUmK020Wx'
        )
    """
    return CustomOpenAIModel(
        model_name=model_name,
        base_url=base_url,
        api_key=api_key,
        **kwargs
    )


def create_glm_model(api_key: str) -> CustomOpenAIModel:
    """Create a GLM model instance with default settings.
    
    Args:
        api_key: Your GLM API key
        
    Returns:
        CustomOpenAIModel configured for GLM
    """
    return create_custom_model(
        model_name='glm-4.6',
        base_url='https://api.z.ai/api/paas/v4',
        api_key=api_key
    )


# Model factory function for easy integration
def create_model_from_string(model_string: str, **kwargs) -> Model:
    """Create a model from a string specification.
    
    Supports formats:
    - 'glm' -> GLM model (requires api_key in kwargs)
    - 'custom:model_name:base_url' -> Custom model (requires api_key in kwargs)
    - 'openai/gpt-4' -> Standard OpenAI model
    - 'openrouter:model_name' -> OpenRouter model
    
    Args:
        model_string: Model specification string
        **kwargs: Additional parameters (api_key, etc.)
        
    Returns:
        Model instance
    """
    if model_string == "glm":
        from .reasoning_models import create_openrouter_model
        return create_glm_model(kwargs.get("api_key"))
    
    elif model_string.startswith("custom:"):
        # Format: custom:model_name:base_url
        parts = model_string.split(":", 2)
        if len(parts) != 3:
            raise ValueError(f"Invalid custom model format: {model_string}")
        
        _, model_name, base_url = parts
        return create_custom_model(
            model_name=model_name,
            base_url=base_url,
            api_key=kwargs.get("api_key"),
            **{k: v for k, v in kwargs.items() if k != "api_key"}
        )
    
    elif model_string.startswith("openrouter:"):
        from .reasoning_models import create_openrouter_model
        return create_openrouter_model(
            model_string[len("openrouter:"):], 
            **kwargs
        )
    
    else:
        # Default to standard model creation
        from pydantic_ai.models import KnownModelName
        return Model(model_string, **kwargs)