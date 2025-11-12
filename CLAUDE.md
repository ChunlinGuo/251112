# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Robovan Business Model Analysis v3.0** project that provides **dual-perspective** quantitative analysis (operator + manufacturer) of autonomous freight vehicles (Robovan) compared to traditional human-driven electric light trucks. The tool generates comprehensive Excel models with scale economics, break-even analysis, and TCO comparison for both buyers and sellers.

## Development Commands

### Setup and Installation
```bash
# Install dependencies
pip install -r requirements.txt
```

### Running the Analysis
```bash
# Generate v3.0 Excel model (recommended)
python generate_model_v3.py

# Generate v2.0 Excel model (legacy)
python generate_model.py
```

Output files:
- v3.0: `output/robovan_analysis_v3.xlsx` (15KB, 6 sheets)
- v2.0: `output/robovan_analysis.xlsx` (legacy, 4 sheets)

## Architecture and Code Structure

### Core Components

**Main Application v3.0**: [`generate_model_v3.py`](generate_model_v3.py) (1,281 lines)
- **Primary Class**: `RobovanAnalyzer_v3`
- **Architecture Pattern**: Object-oriented with dual-perspective analysis
- **Key Features**:
  - Operator perspective TCO (6-year for robovan, 7-year for traditional)
  - Manufacturer profitability analysis with scale economics
  - Break-even analysis (R&D cost allocation)
  - Dynamic formula-driven Excel model

**Legacy Application v2.0**: [`generate_model.py`](generate_model.py) (1,371 lines)
- Single-perspective operator analysis only
- 4-sheet Excel output

**Configuration System**: [`assumptions.json`](assumptions.json) (v2.0, 235 lines)
- All business parameters externally configurable
- **v3.0 additions**:
  - Manufacturer cost parameters (R&D, BOM, production)
  - Remote monitoring parameters (scale effects)
  - One-time purchase mode parameters
- No hardcoded values in calculation logic

### Key Methods Structure (v3.0)

**Excel Sheet Generation Methods** (generate_model_v3.py):
- `create_sheet_parameters()`: Sheet 1 - 6-area parameter configuration (A-F zones)
- `create_sheet_robovan_subscription()`: Sheet 2 - Operator subscription mode (6-year TCO with calculation notes)
- `create_sheet_traditional()`: Sheet 3 - Traditional vehicle (7-year TCO with calculation notes)
- `create_sheet_operator_comparison()`: Sheet 4 - Operator perspective comparison
- `create_sheet_manufacturer()`: Sheet 5 - Manufacturer profitability with break-even analysis ⭐
- `create_sheet_comprehensive()`: Sheet 6 - Comprehensive dual-perspective summary ⭐

**Scale Economics Methods** (v3.0 new):
- R&D cost allocation: `Total R&D = Fixed (380M) + Marginal (3K) × Scale`
- Monitoring cost: `Total = Fixed (5M/yr) + (Scale ÷ Ratio) × Personnel Salary`
- Break-even calculation: `Units = Total R&D ÷ Unit Margin Profit`

**Core Utilities**:
- `load_config()`: JSON configuration loading
- `generate()`: Main orchestration, generates all 6 sheets

### Configuration Architecture

The [`assumptions.json`](assumptions.json) file contains:

**Main Categories**:
- **无人车参数**: Robovan specifications, costs, operational parameters (6-year lifecycle)
- **人驾电动轻卡**: Traditional vehicle specifications (7-year lifecycle)
- **业务场景参数**: Route characteristics, service pricing
- **市场定价参数**: External market rates (distinct from internal costs)
- **厂商成本参数** ⭐ (v3.0): Manufacturer cost structure with scale economics
- **远程监控参数** ⭐ (v3.0): Remote monitoring costs with personnel scaling
- **一次性买断模式** ⭐ (v3.0): One-time purchase alternative to subscription
- **典型情形参数**: 5 predefined scenarios (baseline, mature, scaled, conservative, optimistic)

**Key Parameter Groups**:
- Investment costs (vehicle purchase, lifecycle, residual value)
- Operational costs (labor, software, insurance, maintenance)
- Energy parameters (power consumption, electricity pricing)
- Time parameters (operating hours, days per year, speed)
- Business metrics (route distances, load capacity)
- **Scale economics** ⭐ (v3.0): Sales scale, R&D allocation, monitoring ratios

## Business Logic

### Core Financial Models

**Operator Perspective**:
- **TCO (Total Cost of Ownership)**: 6-year for robovan, 7-year for traditional
- **Unit Cost Analysis**: Per km, per hour, per trip, per item calculations
- **Subscription vs One-time Purchase**: Two business models comparison

**Manufacturer Perspective** ⭐ (v3.0):
- **Single Vehicle Cost**: R&D allocation + BOM + Production + Sales
- **Annual Recurring Cost**: Monitoring center + Software ops + Technical support
- **6-Year Profit Analysis**: Revenue (vehicle + subscription) - Total costs
- **Break-even Analysis**: Minimum units needed to cover R&D investment
- **Scale Economics**: R&D allocation and monitoring cost reduction with scale

### Calculation Approach
- **Dual-Perspective Analysis** ⭐ (v3.0): Both buyer and seller viewpoints
- **Cost-Revenue Separation**: Clear distinction between internal costs and external pricing
- **Scale-Driven Modeling**: Cost reduction with increasing production volume
- **Formula-Driven Excel**: All calculations via cell references to parameter sheet
- **Scenario Testing**: Multiple predefined scenarios for sensitivity analysis

## Technology Stack

- **Python 3.8+**: Main programming language
- **openpyxl**: Excel file generation and formatting
- **pandas**: Data processing and calculations
- **numpy**: Numerical computations
- **No external APIs**: Pure offline computation

## Output Structure

**v3.0 Excel Model** contains 6 worksheets:
1. **参数配置** - 6-area parameter configuration (A-F zones) with editable cells
2. **运营商-订阅模式** - Robovan 6-year TCO with calculation notes
3. **运营商-人驾轻卡** - Traditional vehicle 7-year TCO with calculation notes
4. **运营商对比分析** - Operator perspective comparison
5. **厂商盈利分析** ⭐ - Manufacturer profitability and break-even analysis
6. **综合对比** ⭐ - Dual-perspective comprehensive summary

**v2.0 Excel Model** (legacy) contains 4 worksheets:
1. **参数设置** - Parameter configuration
2. **Robovan模型** - Autonomous vehicle analysis
3. **传统货运** - Traditional vehicle analysis
4. **对比分析** - Comparison analysis

## Key Development Notes

### v3.0 New Features

**Scale Economics Formulas**:
```
R&D Cost Allocation:
  Total R&D = 380M (fixed) + 3K (marginal) × Scale
  Per Vehicle = Total R&D ÷ Scale

Monitoring Cost:
  Total = 5M/yr (fixed) + (Scale ÷ Ratio) × 120K (salary)
  Per Vehicle = Total ÷ Scale

Break-even Units:
  Units = Total R&D ÷ Unit Margin Profit (excluding R&D allocation)
```

**Example**: With 1,000 units and 20:1 ratio:
- R&D per vehicle: 383M ÷ 1,000 = 383K/unit
- Monitoring per vehicle: 11M ÷ 1,000 = 11K/unit/year
- Scale to 5,000: R&D drops to 79K (-79%), Monitoring to 7K (-36%)

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

## Version History

### v3.0 (2025-01-12) - Dual-Perspective Scale Model
- ✅ Manufacturer profitability analysis (Sheet 5)
- ✅ Scale economics (R&D allocation + monitoring costs)
- ✅ Break-even analysis
- ✅ One-time purchase mode parameters
- ✅ Calculation notes columns in all TCO sheets
- ✅ Lifecycle differentiation: 6yr (robovan) vs 7yr (traditional)
- ✅ Comprehensive dual-perspective summary (Sheet 6)

### v2.0 (2025-01-11) - Formula-Driven Dynamic Model
- ✅ Fully formula-driven Excel (no static values)
- ✅ Parameter sheet linkage to all sheets
- ✅ Visual formatting (yellow inputs / blue formulas)

### v1.0 (2025-01-10) - Initial Release
- ✅ Basic TCO analysis
- ✅ Operator perspective only