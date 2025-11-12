# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Robovan Business Model Analysis** project that quantitatively compares the economic viability of autonomous freight vehicles (Robovan) against traditional human-driven electric light trucks in warehousing logistics scenarios. The tool generates comprehensive Excel models for investment decision-making and business planning.

## Development Commands

### Setup and Installation
```bash
# Install dependencies
pip install -r requirements.txt
```

### Running the Analysis
```bash
# Generate the Excel model
python generate_model.py
```

The output Excel file will be generated at: `output/robovan_analysis.xlsx`

## Architecture and Code Structure

### Core Components

**Main Application**: [`generate_model.py`](generate_model.py) (1,371 lines)
- **Primary Class**: `RobovanAnalyzer` (lines 19-1363)
- **Entry Point**: `main()` function at line 1364
- **Architecture Pattern**: Object-oriented with clear separation of concerns

**Configuration System**: [`assumptions.json`](assumptions.json) (158 lines)
- Structured configuration with multiple predefined scenarios
- All business parameters are externally configurable
- No hardcoded values in the calculation logic

### Key Methods Structure

**TCO Calculation Methods**:
- `calculate_robovan_tco()` ([lines 34-142](generate_model.py#L34-L142)): Autonomous vehicle TCO computation
- `calculate_traditional_tco()` ([lines 143-253](generate_model.py#L143-L253)): Traditional vehicle TCO computation

**Excel Sheet Generation Methods**:
- `create_sheet_parameters()` ([lines 254-523](generate_model.py#L254-L523)): Sheet 1 - Configurable parameters
- `create_sheet_robovan()` ([lines 524-733](generate_model.py#L524-L733)): Sheet 2 - Robovan analysis
- `create_sheet_traditional()` ([lines 734-942](generate_model.py#L734-L942)): Sheet 3 - Traditional vehicle analysis
- `create_sheet_comparison()` ([lines 943-1326](generate_model.py#L943-L1326)): Sheet 4 - Comparative analysis

**Core Utilities**:
- `load_config()` ([lines 29-32](generate_model.py#L29-L32)): JSON configuration loading
- `generate()` ([lines 1327-1362](generate_model.py#L1327-L1362)): Main orchestration method

### Configuration Architecture

The [`assumptions.json`](assumptions.json) file contains:

**Main Categories**:
- **无人车参数**: Robovan vehicle specifications, costs, operational parameters
- **传统车辆参数**: Traditional vehicle specifications for comparison
- **业务场景参数**: Route characteristics, service pricing
- **市场定价参数**: External market rates (distinct from internal costs)
- **场景参数**: 5 predefined scenarios (baseline, mature, scaled, conservative, optimistic)

**Key Parameter Groups**:
- Investment costs (vehicle purchase, lifecycle, residual value)
- Operational costs (labor, software, insurance, maintenance)
- Energy parameters (power consumption, electricity pricing)
- Time parameters (operating hours, days per year, speed)
- Business metrics (route distances, load capacity)

## Business Logic

### Core Financial Models
- **TCO (Total Cost of Ownership)**: 5-year lifecycle analysis
- **Unit Cost Analysis**: Per km, per hour, per trip, per item calculations
- **Revenue Modeling**: Based on external market pricing parameters
- **Profitability Analysis**: Break-even periods, ROI calculations

### Calculation Approach
- **Cost-Revenue Separation**: Clear distinction between internal unit costs and external market pricing
- **Comparative Analysis**: Robovan vs traditional vehicle across all metrics
- **Scenario Testing**: Multiple predefined scenarios for sensitivity analysis

## Technology Stack

- **Python 3.8+**: Main programming language
- **openpyxl**: Excel file generation and formatting
- **pandas**: Data processing and calculations
- **numpy**: Numerical computations
- **No external APIs**: Pure offline computation

## Output Structure

Generated Excel file contains 4 worksheets:
1. **参数设置** - All configurable parameters with visual formatting
2. **Robovan模型** - Detailed autonomous vehicle analysis
3. **传统货运** - Traditional vehicle analysis
4. **对比分析** - Direct comparison and competitive advantage quantification

## Key Development Notes

### Parameter Modification
All business parameters are in [`assumptions.json`](assumptions.json). To modify assumptions:
- Edit the JSON file directly
- Use predefined scenarios or create custom scenarios
- Regenerate the Excel model to see changes

### Code Quality
- Clean object-oriented design with single main class
- Comprehensive error handling for file operations
- Professional Excel formatting with color coding
- Detailed function documentation

### Current Limitations
- No unit tests (testing opportunity)
- Basic print statements instead of structured logging
- Manual execution via command line only
- No configuration schema validation