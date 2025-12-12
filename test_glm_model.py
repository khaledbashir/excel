#!/usr/bin/env python3
"""
Test script for using GLM model with Excel Agent.
This shows how to configure and use your custom GLM-4.6 model.
"""

import asyncio
import sys
from pathlib import Path

# Add the project root to Python path
sys.path.insert(0, str(Path(__file__).parent))

from excel_agent.custom_models import create_glm_model, create_custom_model
from excel_agent.config import ExperimentConfig

async def test_glm_model():
    """Test the GLM model configuration."""
    
    print("🤖 Testing GLM Model Configuration")
    print("=" * 50)
    
    # Your GLM API configuration
    GLM_API_KEY = "18f65090a96a425898a8398a5c4518ce.DDtUvTTnUmK020Wx"
    GLM_BASE_URL = "https://api.z.ai/api/paas/v4"
    GLM_MODEL = "glm-4.6"
    
    print(f"✓ Model: {GLM_MODEL}")
    print(f"✓ Base URL: {GLM_BASE_URL}")
    print(f"✓ API Key: {GLM_API_KEY[:20]}...")
    
    try:
        # Create GLM model instance
        glm_model = create_glm_model(GLM_API_KEY)
        print("✓ GLM model created successfully")
        
        # Test model configuration
        config = ExperimentConfig(model='glm')  # We'll update this to use custom model
        print(f"✓ Config created with model: {config.model}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error creating GLM model: {e}")
        return False

def show_available_models():
    """Show all available model options."""
    print("\n🎯 Available Model Options:")
    print("=" * 50)
    
    models = [
        {
            "name": "GLM-4.6 (Your Custom Model)",
            "config": "model='glm'",
            "api_key": "GLM_API_KEY",
            "description": "Your custom GLM model with reasoning capabilities"
        },
        {
            "name": "OpenAI GPT-4",
            "config": "model='openai/gpt-4'",
            "api_key": "OPENAI_API_KEY", 
            "description": "OpenAI's GPT-4 model"
        },
        {
            "name": "OpenAI GPT-4 Turbo",
            "config": "model='openai/gpt-4-turbo'",
            "api_key": "OPENAI_API_KEY",
            "description": "Faster version of GPT-4"
        },
        {
            "name": "Anthropic Claude 3 Sonnet",
            "config": "model='anthropic/claude-3-sonnet'",
            "api_key": "ANTHROPIC_API_KEY",
            "description": "Anthropic's Claude 3 Sonnet"
        },
        {
            "name": "OpenRouter Models",
            "config": "model='openrouter:google/gemini-3-pro-preview'",
            "api_key": "OPENROUTER_API_KEY",
            "description": "Access to various models via OpenRouter"
        },
        {
            "name": "Custom OpenAI-Compatible",
            "config": "model='custom:model-name:base-url'",
            "api_key": "CUSTOM_API_KEY",
            "description": "Any OpenAI-compatible API"
        }
    ]
    
    for i, model in enumerate(models, 1):
        print(f"\n{i}. {model['name']}")
        print(f"   Config: {model['config']}")
        print(f"   API Key: {model['api_key']}")
        print(f"   Description: {model['description']}")

def show_glm_setup():
    """Show specific setup instructions for GLM."""
    print("\n🔧 GLM Model Setup Instructions:")
    print("=" * 50)
    
    print("\n1. Update your .env file:")
    print("   cd /root/excel/demo/ExcelAssistant")
    print("   nano .env")
    print("\n   Add these lines:")
    print("   EXCEL_ASSISTANT_MODEL=glm")
    print("   GLM_API_KEY=18f65090a96a425898a8398a5c4518ce.DDtUvTTnUmK020Wx")
    
    print("\n2. Or use directly in Python:")
    print("""
from excel_agent.custom_models import create_glm_model
from excel_agent.agent_runner import ExcelAgentRunner, TaskInput
from excel_agent.config import ExperimentConfig

# Create GLM model
glm_model = create_glm_model("18f65090a96a425898a8398a5c4518ce.DDtUvTTnUmK020Wx")

# Create config with custom model
config = ExperimentConfig(model='glm')  # Will use GLM model

# Create agent runner
runner = ExcelAgentRunner(
    config=config,
    mcp_server_url="http://127.0.0.1:8765/sse"
)

# Run task
task_input = TaskInput(
    instruction="Add a new row with data: Alice, 28, Paris",
    input_file="test_input.xlsx",
    output_file="test_output.xlsx"
)

result = await runner.run_excel_agent(task_input)
""")
    
    print("\n3. For Web Interface:")
    print("   Stop the current backend server (Ctrl+C)")
    print("   Update .env with GLM settings")
    print("   Restart: python server.py")

async def main():
    """Run the GLM model test."""
    success = await test_glm_model()
    show_available_models()
    show_glm_setup()
    
    if success:
        print(f"\n🎉 GLM model setup is ready!")
        print(f"📝 Check the instructions above to configure your web interface")
    else:
        print(f"\n❌ GLM model setup failed. Check the errors above.")
    
    return success

if __name__ == "__main__":
    success = asyncio.run(main())
    sys.exit(0 if success else 1)