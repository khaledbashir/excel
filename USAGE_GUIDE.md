# Excel Agent Usage Guide

## 🎉 Deployment Complete!

The SylvianAI Excel Agent has been successfully deployed and tested. Here's how to use it:

## 📋 Prerequisites

1. **Virtual Environment**: The `excel_env` is set up and activated
2. **Dependencies**: All required packages are installed
3. **Test Files**: Sample Excel files are created (`test_input.xlsx`, `test_output.xlsx`)

## 🚀 Quick Start

### 1. Start the MCP Server

Open a terminal and run:

```bash
cd /root/excel
source excel_env/bin/activate
python -c "import asyncio; from excel_mcp.excel_server import mcp; asyncio.run(mcp.run_sse_async(host='127.0.0.1', port=8765))"
```

The server will start on `http://127.0.0.1:8765/sse`

### 2. Use the Excel Agent

Create a Python script with your Excel task:

```python
import asyncio
from excel_agent.agent_runner import ExcelAgentRunner, TaskInput
from excel_agent.config import ExperimentConfig

async def run_excel_task():
    # Configure the agent (choose your model)
    config = ExperimentConfig(model='openai/gpt-4')  # or 'anthropic/claude-3-sonnet'
    
    # Create the agent runner
    runner = ExcelAgentRunner(
        config=config,
        mcp_server_url="http://127.0.0.1:8765/sse",
    )
    
    # Define your task
    task_input = TaskInput(
        instruction="Add a new row with data: Alice, 28, Paris",
        input_file="test_input.xlsx",
        output_file="test_output.xlsx",
    )
    
    # Run the agent
    agent_response = await runner.run_excel_agent(task_input)
    print("Task completed!")
    print(f"Response: {agent_response}")

# Run the task
asyncio.run(run_excel_task())
```

## 🔧 Configuration Options

### Available Models

The agent supports various models through different providers:

```python
# OpenAI models
config = ExperimentConfig(model='openai/gpt-4')
config = ExperimentConfig(model='openai/gpt-4-turbo')

# Anthropic models  
config = ExperimentConfig(model='anthropic/claude-3-sonnet')
config = ExperimentConfig(model='anthropic/claude-3-opus')

# OpenRouter models (with reasoning support)
config = ExperimentConfig(model='openrouter:google/gemini-3-pro-preview')
config = ExperimentConfig(model='openrouter:openai/gpt-5.1')

# Local models (if configured)
config = ExperimentConfig(model='ollama/llama2')
```

### Environment Variables

Set your API keys as environment variables:

```bash
export OPENAI_API_KEY="your-openai-key"
export ANTHROPIC_API_KEY="your-anthropic-key"  
export OPENROUTER_API_KEY="your-openrouter-key"
```

## 📊 Available Excel Operations

The Excel Agent provides ~30 tools for spreadsheet operations:

### Data Operations
- **Read cells**: Get data from specific ranges
- **Write cells**: Set values in ranges
- **Search**: Find specific values in the spreadsheet
- **AutoFill**: Automatically fill series

### Sheet Management
- **Create sheets**: Add new worksheets
- **Rename sheets**: Change sheet names
- **Delete sheets**: Remove worksheets
- **Display sheets**: Show sheet information

### Row/Column Operations
- **Insert/Delete Rows**: Add or remove rows
- **Insert/Delete Columns**: Add or remove columns

### Formatting
- **Style ranges**: Apply fonts, colors, borders
- **Conditional formatting**: Add rules-based formatting
- **Data validation**: Set input restrictions

### Advanced Features
- **Formulas**: Write and calculate Excel formulas
- **Merge/Unmerge**: Combine or split cells
- **Batch operations**: Apply changes to multiple ranges

## 🎯 Example Tasks

### Example 1: Data Entry
```python
task_input = TaskInput(
    instruction="Add 5 new employees to the table with names, ages, and cities",
    input_file="employees.xlsx",
    output_file="employees_updated.xlsx",
)
```

### Example 2: Data Analysis
```python
task_input = TaskInput(
    instruction="Calculate the sum of column B and add a total row at the bottom",
    input_file="sales_data.xlsx", 
    output_file="sales_with_total.xlsx",
)
```

### Example 3: Formatting
```python
task_input = TaskInput(
    instruction="Format the header row in bold with a blue background and add borders to all data",
    input_file="report.xlsx",
    output_file="report_formatted.xlsx",
)
```

### Example 4: Formula Creation
```python
task_input = TaskInput(
    instruction="Add a column D with formulas to calculate 10% commission from column C values",
    input_file="sales.xlsx",
    output_file="sales_with_commission.xlsx",
)
```

## 🌐 Demo Applications

### Excel Assistant (Web App)
```bash
cd demo/ExcelAssistant
source ../../excel_env/bin/activate
python app.py
```

### Slack Excel Bot
```bash
cd demo/SlackExcelWorkflow
source ../../excel_env/bin/activate
python bot.py
```

## 🛠️ Advanced Usage

### Custom System Prompts
```python
config = ExperimentConfig(
    model='openai/gpt-4',
    system_prompt="You are an expert financial analyst specializing in Excel spreadsheets..."
)
```

### Timeout Configuration
```python
config = ExperimentConfig(
    model='openai/gpt-4',
    timeout=300.0,  # 5 minutes total timeout
    timeout_request=90.0,  # 90 seconds per API request
    timeout_tool=30.0,  # 30 seconds per tool execution
)
```

## 🔍 Testing and Validation

The deployment includes comprehensive tests:

```bash
source excel_env/bin/activate
python test_setup.py
```

## 📚 Additional Resources

- **Repository**: https://github.com/SylvianAI/sv-excel-agent
- **Documentation**: Check the `demo/` folders for complete examples
- **Evaluation**: Run benchmarks in the `evals/` folder

## 🚨 Important Notes

1. **API Keys Required**: You need valid API keys for the model providers
2. **File Paths**: Use absolute paths for Excel files in production
3. **Memory Usage**: Large Excel files may require increased memory limits
4. **Concurrent Requests**: Start one MCP server per concurrent agent instance

## 🎞️ Next Steps

1. Set up your preferred API provider and key
2. Test with a simple Excel task
3. Explore the demo applications
4. Integrate into your workflow or application

The Excel Agent is now ready to automate your spreadsheet tasks! 🚀