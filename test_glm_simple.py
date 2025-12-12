#!/usr/bin/env python3
"""Simple test for GLM model creation."""

import os
os.environ["GLM_API_KEY"] = "18f65090a96a425898a8398a5c4518ce.DDtUvTTnUmK020Wx"

try:
    from excel_agent.custom_models import create_glm_model
    print("✓ Imported create_glm_model")
    
    glm_model = create_glm_model("18f65090a96a425898a8398a5c4518ce.DDtUvTTnUmK020Wx")
    print("✓ GLM model created successfully")
    print(f"Model: {glm_model}")
    print(f"Model name: {glm_model.model_name}")
    print(f"Client base URL: {glm_model.client.base_url}")
    print(f"Client API key: {glm_model.client.api_key[:20]}...")
    
except Exception as e:
    print(f"❌ Error: {e}")
    import traceback
    traceback.print_exc()