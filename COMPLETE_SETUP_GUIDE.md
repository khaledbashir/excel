# Excel Agent - Complete Setup Guide

## 🎯 Overview
The SylvianAI Excel Agent is now fully deployed with:
- ✅ Core Excel MCP Server (port 8765)
- ✅ Web Interface with GLM-4.6 model (port 8000)
- ✅ Custom GLM-4.6 model integration
- ✅ All dependencies installed and tested

## 🚀 Quick Start

### 1. Access the Web Interface
Open your browser and go to: **http://localhost:8000**

The web interface provides:
- File upload for Excel files
- Natural language instructions
- Real-time processing with GLM-4.6
- Downloadable results

### 2. Current Configuration
- **Model**: GLM-4.6 (your custom model)
- **API Endpoint**: https://api.z.ai/api/paas/v4
- **API Key**: Configured and working
- **Web Server**: http://localhost:8000
- **MCP Server**: http://127.0.0.1:8765/sse

## 🔧 Technical Details

### Server Status
```bash
# MCP Server (Excel Tools)
Running on: http://127.0.0.1:8765/sse
Status: ✅ Active

# Web Interface (Frontend + Backend)
Running on: http://localhost:8000
Status: ✅ Active
```

### Model Configuration
The system is configured to use your GLM-4.6 model:

```env
EXCEL_ASSISTANT_MODEL=glm
GLM_API_KEY=18f65090a96a425898a8398a5c4518ce.DDtUvTTnUmK020Wx
```

### File Structure
```
/root/excel/
├── excel_env/                    # Python virtual environment
├── excel_agent/                  # Core agent code
│   ├── custom_models.py         # GLM-4.6 model wrapper
│   └── reasoning_models.py      # Model routing logic
├── demo/ExcelAssistant/         # Web interface
│   ├── dist/                   # Built frontend
│   ├── server.py               # FastAPI backend
│   └── .env                    # Configuration
├── test_setup.py               # Core functionality tests
├── test_glm_model.py           # GLM model tests
└── COMPLETE_SETUP_GUIDE.md     # This guide
```

## 🎮 How to Use

### Web Interface (Recommended)
1. Open **http://localhost:8000** in your browser
2. Upload an Excel file
3. Type your instruction in natural language
4. Click "Process" and wait for results
5. Download the modified file

### Example Instructions
- "Add a new row with data: Alice, 28, Paris"
- "Calculate the sum of column B and put it in cell B10"
- "Create a chart from the data in range A1:C10"
- "Filter the data to show only rows where column A > 100"

### API Usage
You can also use the API directly:

```python
import asyncio
from excel_agent.agent_runner import ExcelAgentRunner, TaskInput
from excel_agent.config import ExperimentConfig

async def process_excel():
    config = ExperimentConfig(model='glm')
    runner = ExcelAgentRunner(
        config=config,
        mcp_server_url="http://127.0.0.1:8765/sse"
    )
    
    task_input = TaskInput(
        instruction="Add a summary row at the bottom",
        input_file="input.xlsx",
        output_file="output.xlsx"
    )
    
    result = await runner.run_excel_agent(task_input)
    return result

# Run the task
result = asyncio.run(process_excel())
```

## 🛠️ Available Models

### Currently Active
- **GLM-4.6**: Your custom model (configured and working)

### Alternative Options
You can switch to other models by updating the `.env` file:

```env
# OpenAI GPT-4
EXCEL_ASSISTANT_MODEL=openai/gpt-4
OPENAI_API_KEY=your_openai_key

# Anthropic Claude
EXCEL_ASSISTANT_MODEL=anthropic/claude-3-sonnet
ANTHROPIC_API_KEY=your_anthropic_key

# OpenRouter (various models)
EXCEL_ASSISTANT_MODEL=openrouter:google/gemini-3-pro-preview
OPENROUTER_API_KEY=your_openrouter_key
```

## 🔍 Testing

### Run Core Tests
```bash
cd /root/excel
source excel_env/bin/activate
python test_setup.py
```

### Test GLM Model
```bash
cd /root/excel
source excel_env/bin/activate
python test_glm_model.py
```

### Test Web Interface
```bash
curl http://localhost:8000/api/health
# Expected: {"status":"ok"}
```

## 🚨 Troubleshooting

### If Web Interface is Down
```bash
# Restart the web server
cd /root/excel/demo/ExcelAssistant
source ../../excel_env/bin/activate
python server.py
```

### If MCP Server is Down
```bash
# Restart the MCP server
cd /root/excel
source excel_env/bin/activate
python -m excel_agent.mcp_server --port 8765
```

### Port Conflicts
If you get port conflicts:
```bash
# Kill processes on ports 8000 or 8765
lsof -ti:8000 | xargs kill -9
lsof -ti:8765 | xargs kill -9
```

## 📝 Environment Variables

Key environment variables in `/root/excel/demo/ExcelAssistant/.env`:

```env
# Model Configuration
EXCEL_ASSISTANT_MODEL=glm
GLM_API_KEY=18f65090a96a425898a8398a5c4518ce.DDtUvTTnUmK020Wx

# Server Configuration
MCP_SERVER_URL=http://127.0.0.1:8765/sse
HOST=0.0.0.0
PORT=8000
```

## 🎉 Success!

Your Excel Agent is now fully operational with:
- ✅ GLM-4.6 model integration
- ✅ Web interface at http://localhost:8000
- ✅ Excel processing capabilities
- ✅ All tests passing

You can now start using the system to process Excel files with natural language instructions powered by your GLM-4.6 model!