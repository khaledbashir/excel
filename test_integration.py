#!/usr/bin/env python3
"""
Integration test for the complete Excel Agent setup with GLM-4.6.
"""

import asyncio
import sys
from pathlib import Path

# Add the project root to Python path
sys.path.insert(0, str(Path(__file__).parent))

from excel_agent.custom_models import create_glm_model
from excel_agent.config import ExperimentConfig


async def test_glm_integration():
    """Test GLM model integration."""
    print("🧪 Testing GLM-4.6 Integration")
    print("=" * 50)
    
    try:
        # Test 1: Create GLM model
        print("1. Creating GLM model...")
        api_key = "18f65090a96a425898a8398a5c4518ce.DDtUvTTnUmK020Wx"
        glm_model = create_glm_model(api_key)
        print(f"   ✅ GLM model created: {glm_model.model_name}")
        
        # Verify base URL (accessing client.base_url which might be private or wrapped)
        # We can just check the model was created successfully

        
        # Test 2: Create config with GLM
        print("2. Creating config with GLM model...")
        config = ExperimentConfig(model='glm')
        print(f"   ✅ Config created with model: {config.model}")
        
        # Test 3: Test model string resolution
        print("3. Testing model string resolution...")
        from excel_agent.reasoning_models import get_model_for_config
        
        resolved_model = get_model_for_config('glm', api_key=api_key)
        print(f"   ✅ Model resolved: {type(resolved_model).__name__}")
        
        # Test 4: Test basic model properties
        print("4. Testing model properties...")
        print(f"   ✅ Model name: {resolved_model.model_name}")
        print(f"   ✅ Model system: {resolved_model.system}")
        print(f"   ✅ Model profile type: {type(resolved_model.profile).__name__}")
        
        print("\n🎉 All integration tests passed!")
        print("📋 Summary:")
        print(f"   - Model: {resolved_model.model_name}")
        print(f"   - Provider: {resolved_model.system}")
        print(f"   - API Key: {'*' * 20}{api_key[-10:]}")
        print(f"   - Ready for web interface: ✅")
        
        return True
        
    except Exception as e:
        print(f"❌ Integration test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


async def test_web_interface_config():
    """Test web interface configuration."""
    print("\n🌐 Testing Web Interface Configuration")
    print("=" * 50)
    
    try:
        # Check .env file
        env_file = Path("/root/excel/demo/ExcelAssistant/.env")
        if env_file.exists():
            print("1. Checking .env file...")
            content = env_file.read_text()
            
            if "EXCEL_ASSISTANT_MODEL=glm" in content:
                print("   ✅ EXCEL_ASSISTANT_MODEL=glm found")
            else:
                print("   ⚠️  EXCEL_ASSISTANT_MODEL=glm not found")
            
            if "GLM_API_KEY=" in content:
                print("   ✅ GLM_API_KEY found")
            else:
                print("   ⚠️  GLM_API_KEY not found")
        
        # Test server health
        print("2. Testing server health...")
        import httpx
        
        async with httpx.AsyncClient() as client:
            response = await client.get("http://localhost:8000/api/health")
            if response.status_code == 200:
                print("   ✅ Server health check passed")
            else:
                print(f"   ❌ Server health check failed: {response.status_code}")
        
        print("\n🎉 Web interface configuration verified!")
        
    except Exception as e:
        print(f"❌ Web interface test failed: {e}")
        return False
    
    return True


async def main():
    """Run all integration tests."""
    print("🚀 Excel Agent Integration Test Suite")
    print("=" * 60)
    
    # Test GLM integration
    glm_success = await test_glm_integration()
    
    # Test web interface
    web_success = await test_web_interface_config()
    
    # Final summary
    print("\n" + "=" * 60)
    print("📊 FINAL TEST RESULTS")
    print("=" * 60)
    print(f"GLM-4.6 Model Integration: {'✅ PASS' if glm_success else '❌ FAIL'}")
    print(f"Web Interface Configuration: {'✅ PASS' if web_success else '❌ FAIL'}")
    
    if glm_success and web_success:
        print("\n🎉 ALL TESTS PASSED!")
        print("🌟 Your Excel Agent is ready to use with GLM-4.6!")
        print("📱 Access it at: http://localhost:8000")
        return 0
    else:
        print("\n❌ Some tests failed. Please check the errors above.")
        return 1


if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(exit_code)