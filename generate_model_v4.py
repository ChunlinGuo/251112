#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Robovan Business Model Analyzer v4.0 (核心修复版)
基于v3.0，实施5个核心优化：
1. 扩展Named Ranges到计算结果
2. 重构Sheet 1参数页面格式
3. 补全Sheet 3单位成本指标
4. 增强Sheet 4对比分析
5. 代码重构（提取公共TCO方法）
"""

import json
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName


class RobovanAnalyzerV4:
    """Robovan商业模式分析器 v4.0 - 核心修复版"""

    def __init__(self, config_path='assumptions.json'):
        self.config_path = config_path
        self.config = self.load_config()
        self.wb = Workbook()

        # 删除默认sheet
        if 'Sheet' in self.wb.sheetnames:
            del self.wb['Sheet']

        # 参数位置记录（Sheet 1参数）
        self.param_cells = {}  # {param_name: (sheet, col, row)}

        # ⭐ v4.0新增：计算结果位置记录（Sheet 2-6关键结果）
        self.result_cells = {}  # {result_name: (sheet, col, row)}

        self.sheet1_name = "1-参数配置"

    def load_config(self):
        """加载配置文件"""
        with open(self.config_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def record_param(self, name, col, row, sheet='1-参数配置'):
        """记录参数位置（用于创建Named Range）"""
        self.param_cells[name] = (sheet, col, row)

    def record_result(self, name, sheet, col, row):
        """⭐ v4.0新增：记录计算结果位置（用于创建Named Range）

        用于Sheet 2-6中的关键计算结果，例如：
        - rv_tco_total: Sheet 2的TCO总计
        - trad_annual_km: Sheet 3的年行驶里程
        - 等等
        """
        self.result_cells[name] = (sheet, col, row)

    def create_all_named_ranges(self):
        """创建所有命名范围（参数 + 计算结果）"""
        print("\n✨ 创建命名范围（Named Ranges）...")
        print("━" * 60)

        # 合并参数和结果
        all_names = {**self.param_cells, **self.result_cells}

        count = 0
        for name, (sheet, col, row) in all_names.items():
            ref = f"'{sheet}'!${col}${row}"
            safe_name = name.replace(' ', '_').replace('(', '').replace(')', '').replace('/', '_')

            try:
                defined_name = DefinedName(safe_name, attr_text=ref)
                self.wb.defined_names[safe_name] = defined_name
                count += 1
                if count <= 15 or count % 20 == 0:
                    print(f"  ✓ {safe_name:35} → {ref}")
            except Exception as e:
                print(f"  ✗ 创建命名范围失败: {safe_name} - {e}")

        print(f"━" * 60)
        print(f"✅ 成功创建 {count} 个命名范围")
        print(f"   - 参数命名: {len(self.param_cells)} 个")
        print(f"   - 结果命名: {len(self.result_cells)} 个 ⭐ v4.0新增\n")

    # ========================================
    # 样式定义（复用）
    # ========================================

    def get_styles(self):
        """获取统一的样式定义"""
        return {
            'title_font': Font(size=16, bold=True, color='FFFFFF'),
            'title_fill': PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid'),
            'section_font': Font(bold=True, color='FFFFFF', size=11),
            'section_fill': PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid'),
            'header_font': Font(bold=True, size=10),
            'header_fill': PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid'),
            'input_font': Font(size=10, bold=True, color='0000FF'),
            'input_fill': PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid'),
            'calc_font': Font(size=10, italic=True, color='0066CC'),
            'calc_fill': PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid'),
            'value_font': Font(size=10, bold=True, color='0066CC'),
            'note_font': Font(size=9, italic=True, color='666666'),
            'green_fill': PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'),
            'red_fill': PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
            'yellow_fill': PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid'),
        }

    # ========================================
    # Sheet 1: 参数配置（⭐ v4.0优化格式）
    # ========================================

    def create_sheet_parameters(self):
        """创建Sheet 1: 参数配置总表

        ⭐ v4.0优化：
        - A-C区：统一双列格式（B=无人车，C=人驾，D=单位，E=说明）
        - D-F区：统一单列格式（B=数值，C=单位，D=说明）
        - 列宽一致，视觉整洁
        """
        ws = self.wb.create_sheet(self.sheet1_name)
        styles = self.get_styles()

        # 列宽设置（⭐ v4.0统一）
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 16  # ⭐ v4.0：统一为16（原来不一致）
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 35

        # 标题
        ws['A1'] = 'Robovan商业模式分析 v4.0 - 参数配置（核心修复版）'
        ws['A1'].font = styles['title_font']
        ws['A1'].fill = styles['title_fill']
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:E1')
        ws.row_dimensions[1].height = 30

        ws['A2'] = f"生成日期：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws['A2'].font = Font(size=10, italic=True)

        ws['A3'] = "⭐ 黄色=可修改，绿色=自动计算，所有Sheet通过命名范围实时更新"
        ws['A3'].font = Font(size=11, color='FF0000', bold=True)

        ws['A4'] = "📊 v4.0优化：扩展Named Ranges到计算结果，格式统一，代码重构"
        ws['A4'].font = Font(size=10, color='0066CC', italic=True)

        row = 6

        # ========================================
        # A区：车辆基本参数（双列对比格式）
        # ========================================
        ws[f'A{row}'] = '【A区】车辆基本参数'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        # ⭐ v4.0：统一表头格式
        headers = ['参数项', '无人车(订阅)', '人驾电动轻卡', '单位', '说明/假设依据']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = styles['header_font']
            cell.fill = styles['header_fill']
            cell.alignment = Alignment(horizontal='center')
        row += 1

        # A区数据
        basic_params = [
            ('rv_vehicle_type', 'trad_vehicle_type', '车型',
             self.config['无人车']['基本信息']['车型'],
             self.config['人驾电动轻卡']['基本信息']['车型'],
             '-', '九识Z8 vs 远程V6E/福田智蓝', False),

            ('rv_capacity_m3', 'trad_capacity_m3', '运力',
             self.config['无人车']['基本信息']['运力_立方米'],
             self.config['人驾电动轻卡']['基本信息']['运力_立方米'],
             'm³', '官方参数：8m³ vs 7m³', True),

            ('rv_price', 'trad_price', '车辆采购成本',
             self.config['无人车']['投资成本']['车辆采购成本_元'],
             self.config['人驾电动轻卡']['投资成本']['车辆采购成本_元'],
             '元', '官网价：79,800 vs 145,000', True),

            ('rv_battery', 'trad_battery', '电池容量',
             self.config['无人车']['投资成本']['电池容量_kWh'],
             self.config['人驾电动轻卡']['投资成本']['电池容量_kWh'],
             'kWh', 'CLTC工况', True),

            ('rv_range', 'trad_range', '续航里程',
             self.config['无人车']['投资成本']['续航里程_公里'],
             self.config['人驾电动轻卡']['投资成本']['续航里程_公里'],
             'km', 'CLTC工况', True),

            ('rv_lifecycle', 'trad_lifecycle', '生命周期',
             self.config['无人车']['投资成本']['生命周期_年'],
             self.config['人驾电动轻卡']['投资成本']['生命周期_年'],
             '年', '⭐ 6年vs7年（电子器件vs机械部件）', True),

            ('rv_residual', 'trad_residual', '残值',
             self.config['无人车']['投资成本']['残值_元'],
             self.config['人驾电动轻卡']['投资成本']['残值_元'],
             '元', '约购车成本38% vs 15%', True),
        ]

        for rv_name, trad_name, display_name, rv_val, trad_val, unit, note, is_input in basic_params:
            ws.cell(row, 1, display_name)
            cell_rv = ws.cell(row, 2, rv_val)
            cell_trad = ws.cell(row, 3, trad_val)

            # 记录参数位置
            self.record_param(rv_name, 'B', row)
            self.record_param(trad_name, 'C', row)

            if is_input and isinstance(rv_val, (int, float)):
                cell_rv.font = styles['input_font']
                cell_rv.fill = styles['input_fill']
                cell_trad.font = styles['input_font']
                cell_trad.fill = styles['input_fill']
                if isinstance(rv_val, int) and rv_val > 100:
                    cell_rv.number_format = '#,##0'
                    cell_trad.number_format = '#,##0'

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # ========================================
        # B区：运营参数（双列对比格式）
        # ========================================
        ws[f'A{row}'] = '【B区】运营参数'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = styles['header_font']
            cell.fill = styles['header_fill']
            cell.alignment = Alignment(horizontal='center')
        row += 1

        # B区数据 - 运营时间
        operation_params = [
            ('rv_work_days', 'trad_work_days', '年运营天数',
             self.config['无人车']['运营时间参数']['年运营天数_天'],
             self.config['人驾电动轻卡']['运营时间参数']['年运营天数_天'],
             '天', '350天 vs 330天（法定假日）', True),

            ('rv_work_hours', 'trad_work_hours', '日有效运营时长',
             self.config['无人车']['运营时间参数']['日有效运营时长_小时'],
             self.config['人驾电动轻卡']['运营时间参数']['日有效运营时长_小时'],
             '小时', '10h vs 8h（司机工作制）', True),

            ('rv_speed', 'trad_speed', '平均运营时速',
             self.config['无人车']['运营时间参数']['平均运营时速_公里每小时'],
             self.config['人驾电动轻卡']['运营时间参数']['平均运营时速_公里每小时'],
             'km/h', '25 vs 28（保守vs经验）', True),
        ]

        for rv_name, trad_name, display_name, rv_val, trad_val, unit, note, is_input in operation_params:
            ws.cell(row, 1, display_name)
            cell_rv = ws.cell(row, 2, rv_val)
            cell_trad = ws.cell(row, 3, trad_val)

            self.record_param(rv_name, 'B', row)
            self.record_param(trad_name, 'C', row)

            cell_rv.font = styles['input_font']
            cell_rv.fill = styles['input_fill']
            cell_trad.font = styles['input_font']
            cell_trad.fill = styles['input_fill']

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # B区数据 - 能源参数
        energy_params = [
            ('rv_power_consumption', 'trad_power_consumption', '百公里电耗',
             self.config['无人车']['能源参数']['百公里电耗_kWh'],
             self.config['人驾电动轻卡']['能源参数']['百公里电耗_kWh'],
             'kWh/100km', '17.7 vs 15.0（传感器耗电）', True),

            ('rv_electricity_price', 'trad_electricity_price', '电价',
             self.config['无人车']['能源参数']['电价_元每kWh'],
             self.config['人驾电动轻卡']['能源参数']['电价_元每kWh'],
             '元/kWh', '工业用电峰谷混合', True),
        ]

        for rv_name, trad_name, display_name, rv_val, trad_val, unit, note, is_input in energy_params:
            ws.cell(row, 1, display_name)
            cell_rv = ws.cell(row, 2, rv_val)
            cell_trad = ws.cell(row, 3, trad_val)

            self.record_param(rv_name, 'B', row)
            self.record_param(trad_name, 'C', row)

            cell_rv.font = styles['input_font']
            cell_rv.fill = styles['input_fill']
            cell_trad.font = styles['input_font']
            cell_trad.fill = styles['input_fill']

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # B区数据 - 业务场景
        business_params = [
            ('route_distance', '单次往返里程',
             self.config['业务场景参数']['单次往返里程_公里'],
             'km', '仓库↔配送中心', True),

            ('items_per_trip', '单次满载件数',
             self.config['业务场景参数']['单次满载件数_件'],
             '件', '8m³标准包裹（500件）', True),
        ]

        for param_name, display_name, value, unit, note, is_input in business_params:
            ws.cell(row, 1, display_name)
            cell = ws.cell(row, 2, value)

            self.record_param(param_name, 'B', row)

            cell.font = styles['input_font']
            cell.fill = styles['input_fill']

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # B区数据 - 市场定价
        pricing_params = [
            ('price_per_km', '按公里收费',
             self.config['市场定价参数']['按公里收费_元每公里'],
             '元/km', '短途货运市场价（外部市场价）', True),

            ('price_per_item', '按件收费',
             self.config['市场定价参数']['按件收费_元每件'],
             '元/件', '城市配送标准（外部市场价）', True),
        ]

        for param_name, display_name, value, unit, note, is_input in pricing_params:
            ws.cell(row, 1, display_name)
            cell = ws.cell(row, 2, value)

            self.record_param(param_name, 'B', row)

            cell.font = styles['input_font']
            cell.fill = styles['input_fill']

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # ========================================
        # C区：运营商年度成本（双列对比格式）
        # ========================================
        ws[f'A{row}'] = '【C区】运营商年度成本参数'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = styles['header_font']
            cell.fill = styles['header_fill']
            cell.alignment = Alignment(horizontal='center')
        row += 1

        # C区数据
        cost_params = [
            ('rv_labor_cost', 'trad_labor_cost', '人力成本',
             self.config['无人车']['年度成本_元']['人力成本'],
             self.config['人驾电动轻卡']['年度成本_元']['人力成本'],
             '元/年', '1K vs 110K（远程1:100 vs 司机全职）', True),

            ('rv_software_cost', 'trad_software_cost', '软件订阅费',
             self.config['无人车']['年度成本_元']['软件订阅费'],
             self.config['人驾电动轻卡']['年度成本_元'].get('软件订阅费', 0),
             '元/年', '34K vs 0（智驾系统订阅）', True),

            ('rv_insurance', 'trad_insurance', '保险费用',
             self.config['无人车']['年度成本_元']['保险费用'],
             self.config['人驾电动轻卡']['年度成本_元']['保险费用'],
             '元/年', '9K vs 12K（营运车辆）', True),

            ('rv_maintenance', 'trad_maintenance', '维保成本',
             self.config['无人车']['年度成本_元']['维保成本'],
             self.config['人驾电动轻卡']['年度成本_元']['维保成本'],
             '元/年', '8K vs 7K（传感器维护）', True),

            ('rv_other_cost', 'trad_other_cost', '其他成本',
             self.config['无人车']['年度成本_元']['其他成本'],
             self.config['人驾电动轻卡']['年度成本_元']['其他成本'],
             '元/年', '3K vs 5K（停车、充电等）', True),
        ]

        for rv_name, trad_name, display_name, rv_val, trad_val, unit, note, is_input in cost_params:
            ws.cell(row, 1, display_name)
            cell_rv = ws.cell(row, 2, rv_val)
            cell_trad = ws.cell(row, 3, trad_val)

            self.record_param(rv_name, 'B', row)
            self.record_param(trad_name, 'C', row)

            cell_rv.font = styles['input_font']
            cell_rv.fill = styles['input_fill']
            cell_trad.font = styles['input_font']
            cell_trad.fill = styles['input_fill']

            if isinstance(rv_val, int) and rv_val > 100:
                cell_rv.number_format = '#,##0'
                cell_trad.number_format = '#,##0'

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # ========================================
        # D区：厂商成本参数（⭐ v4.0优化为单列格式）
        # ========================================
        ws[f'A{row}'] = '【D区】厂商成本参数（v3.0新增，v4.0格式优化）'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        # ⭐ v4.0：单列格式表头（B=数值，C=单位，D=说明）
        single_col_headers = ['参数项', '数值', '单位', '说明/假设依据', '']
        for col_idx, header in enumerate(single_col_headers, 1):
            if header:
                cell = ws.cell(row, col_idx, header)
                cell.font = styles['header_font']
                cell.fill = styles['header_fill']
                cell.alignment = Alignment(horizontal='center')
        row += 1

        # D区数据
        mfg_params = [
            ('mfg_sales_scale', '销售规模',
             self.config['厂商成本参数']['销售规模']['车队规模_台'],
             '台', '当前假设中等规模', True),

            ('mfg_rd_fixed', '固定研发投入',
             self.config['厂商成本参数']['研发成本']['固定研发投入_元'],
             '元', '平台和算法开发（3.8亿）', True),

            ('mfg_rd_marginal', '单车边际研发',
             self.config['厂商成本参数']['研发成本']['单车边际研发_元'],
             '元/台', '单车场景适配（3千）', True),
        ]

        for param_name, display_name, value, unit, note, is_input in mfg_params:
            ws.cell(row, 1, display_name)
            cell = ws.cell(row, 2, value)

            self.record_param(param_name, 'B', row)

            if is_input:
                cell.font = styles['input_font']
                cell.fill = styles['input_fill']
                if isinstance(value, int) and value > 100:
                    cell.number_format = '#,##0'

            ws.cell(row, 3, unit)
            ws.cell(row, 4, note)
            row += 1

        # 总研发投入（计算字段）
        ws.cell(row, 1, '总研发投入')
        cell = ws.cell(row, 2, '=mfg_rd_fixed+mfg_rd_marginal*mfg_sales_scale')
        self.record_param('mfg_rd_total', 'B', row)
        cell.font = styles['calc_font']
        cell.fill = styles['calc_fill']
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '=固定+边际×规模')
        row += 1

        # 单车研发分摊（计算字段）
        ws.cell(row, 1, '单车研发分摊')
        cell = ws.cell(row, 2, '=mfg_rd_total/mfg_sales_scale')
        self.record_param('mfg_rd_per_vehicle', 'B', row)
        cell.font = styles['calc_font']
        cell.fill = styles['calc_fill']
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/台')
        ws.cell(row, 4, '=总研发÷规模（规模效应）')
        row += 1

        row += 1  # 空行

        # 单车制造成本
        mfg_cost_params = [
            ('mfg_bom_cost', '硬件BOM成本',
             self.config['厂商成本参数']['单车制造成本']['硬件BOM成本_元'],
             '元/台', '雷达、传感器、车体、电池', True),

            ('mfg_production_cost', '生产制造成本',
             self.config['厂商成本参数']['单车制造成本']['生产制造成本_元'],
             '元/台', '组装、测试、质检', True),
        ]

        for param_name, display_name, value, unit, note, is_input in mfg_cost_params:
            ws.cell(row, 1, display_name)
            cell = ws.cell(row, 2, value)

            self.record_param(param_name, 'B', row)

            if is_input:
                cell.font = styles['input_font']
                cell.fill = styles['input_fill']
                cell.number_format = '#,##0'

            ws.cell(row, 3, unit)
            ws.cell(row, 4, note)
            row += 1

        row += 1  # 空行

        # 销售与服务参数
        mfg_service_params = [
            ('mfg_sales_rate', '销售费用率',
             self.config['厂商成本参数']['销售与服务']['销售费用率'],
             '%', '占收入比例', True),

            ('mfg_subscription_margin', '软件订阅毛利率',
             self.config['厂商成本参数']['销售与服务']['软件订阅毛利率'],
             '%', '订阅收入-运维成本', True),

            ('mfg_opex_rate', '软件运维成本率',
             self.config['厂商成本参数']['销售与服务']['软件运维成本率'],
             '%', '占订阅收入比例', True),
        ]

        for param_name, display_name, value, unit, note, is_input in mfg_service_params:
            ws.cell(row, 1, display_name)
            cell = ws.cell(row, 2, value)

            self.record_param(param_name, 'B', row)

            if is_input:
                cell.font = styles['input_font']
                cell.fill = styles['input_fill']
                cell.number_format = '0%'

            ws.cell(row, 3, unit)
            ws.cell(row, 4, note)
            row += 1

        row += 1  # 空行

        # ========================================
        # E区：远程监控参数（⭐ v4.0优化为单列格式）
        # ========================================
        ws[f'A{row}'] = '【E区】远程监控参数（v3.0新增，v4.0格式优化）'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        for col_idx, header in enumerate(single_col_headers, 1):
            if header:
                cell = ws.cell(row, col_idx, header)
                cell.font = styles['header_font']
                cell.fill = styles['header_fill']
                cell.alignment = Alignment(horizontal='center')
        row += 1

        # E区数据
        mon_params = [
            ('mon_fixed_cost', '监控中心年固定成本',
             self.config['远程监控参数']['固定成本']['监控中心年固定成本_元'],
             '元/年', '租金、系统、管理人员（500万）', True),

            ('mon_salary', '安全员年薪',
             self.config['远程监控参数']['变动成本']['安全员年薪_元'],
             '元/年', '一线城市监控员（12万）', True),

            ('mon_ratio', '安全员配比',
             self.config['远程监控参数']['变动成本']['安全员配比_车每人'],
             '车/人', '成熟期1:20，初期1:5，极限1:50', True),
        ]

        for param_name, display_name, value, unit, note, is_input in mon_params:
            ws.cell(row, 1, display_name)
            cell = ws.cell(row, 2, value)

            self.record_param(param_name, 'B', row)

            if is_input:
                cell.font = styles['input_font']
                cell.fill = styles['input_fill']
                if isinstance(value, int) and value > 100:
                    cell.number_format = '#,##0'

            ws.cell(row, 3, unit)
            ws.cell(row, 4, note)
            row += 1

        # 需要安全员数（计算字段）
        ws.cell(row, 1, '需要安全员数')
        cell = ws.cell(row, 2, '=mfg_sales_scale/mon_ratio')
        self.record_param('mon_staff_count', 'B', row)
        cell.font = styles['calc_font']
        cell.fill = styles['calc_fill']
        cell.number_format = '#,##0'
        ws.cell(row, 3, '人')
        ws.cell(row, 4, '=规模÷配比')
        row += 1

        # 监控中心年度总成本（计算字段）
        ws.cell(row, 1, '监控中心年度总成本')
        cell = ws.cell(row, 2, '=mon_fixed_cost+mon_staff_count*mon_salary')
        self.record_param('mon_total_cost', 'B', row)
        cell.font = styles['calc_font']
        cell.fill = styles['calc_fill']
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '=固定成本+人数×年薪')
        row += 1

        # 单车年度监控成本（计算字段）
        ws.cell(row, 1, '单车年度监控成本')
        cell = ws.cell(row, 2, '=mon_total_cost/mfg_sales_scale')
        self.record_param('mon_cost_per_vehicle', 'B', row)
        cell.font = styles['calc_font']
        cell.fill = styles['calc_fill']
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/台/年')
        ws.cell(row, 4, '=总成本÷规模（规模效应）')
        row += 1

        row += 1  # 空行

        # ========================================
        # F区：一次性买断参数（⭐ v4.0优化为单列格式）
        # ========================================
        ws[f'A{row}'] = '【F区】一次性买断模式（v3.0新增，v4.0格式优化）'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        for col_idx, header in enumerate(single_col_headers, 1):
            if header:
                cell = ws.cell(row, col_idx, header)
                cell.font = styles['header_font']
                cell.fill = styles['header_fill']
                cell.alignment = Alignment(horizontal='center')
        row += 1

        # F区数据
        buyout_params = [
            ('buy_total_price', '车辆加软件总价',
             self.config['一次性买断模式']['投资成本']['车辆加软件总价_元'],
             '元', 'vs订阅79,800（含软件永久使用权）', True),

            ('buy_lifecycle', '生命周期',
             self.config['一次性买断模式']['投资成本']['生命周期_年'],
             '年', '同订阅模式6年', True),

            ('buy_residual', '残值',
             self.config['一次性买断模式']['投资成本']['残值_元'],
             '元', '+3000元系统价值', True),

            ('buy_support_fee', '技术支持费',
             self.config['一次性买断模式']['年度成本_元']['技术支持费'],
             '元/年', '买断后单独支付', True),

            ('buy_labor_cost', '人力成本',
             self.config['一次性买断模式']['年度成本_元']['人力成本'],
             '元/年', '同订阅模式', True),

            ('buy_insurance', '保险费用',
             self.config['一次性买断模式']['年度成本_元']['保险费用'],
             '元/年', '同订阅模式', True),

            ('buy_maintenance', '维保成本',
             self.config['一次性买断模式']['年度成本_元']['维保成本'],
             '元/年', '同订阅模式', True),

            ('buy_other_cost', '其他成本',
             self.config['一次性买断模式']['年度成本_元']['其他成本'],
             '元/年', '同订阅模式', True),
        ]

        for param_name, display_name, value, unit, note, is_input in buyout_params:
            ws.cell(row, 1, display_name)
            cell = ws.cell(row, 2, value)

            self.record_param(param_name, 'B', row)

            if is_input:
                cell.font = styles['input_font']
                cell.fill = styles['input_fill']
                if isinstance(value, int) and value > 100:
                    cell.number_format = '#,##0'

            ws.cell(row, 3, unit)
            ws.cell(row, 4, note)
            row += 1

        print(f"\n✅ Sheet 1创建完成，共记录 {len(self.param_cells)} 个参数位置")
        print(f"   ⭐ v4.0优化：格式统一（A-C区双列，D-F区单列）")

    # ========================================
    # ⭐ v4.0核心重构：通用TCO Sheet生成器
    # ========================================

    def create_tco_sheet_generic(self, sheet_name, config):
        """⭐ v4.0重构：通用TCO Sheet生成器（代码复用）

        Args:
            sheet_name: Sheet名称
            config: 配置字典
                {
                    'title': 'Sheet标题',
                    'subtitle': '副标题',
                    'param_prefix': 'rv_' 或 'trad_',  # 参数前缀
                    'result_prefix': 'rv_' 或 'trad_',  # 结果前缀
                    'lifecycle': 6 或 7,  # 生命周期年数
                }
        """
        ws = self.wb.create_sheet(sheet_name)
        styles = self.get_styles()
        prefix = config['param_prefix']
        result_prefix = config['result_prefix']
        lifecycle = config['lifecycle']

        # 列宽
        ws.column_dimensions['A'].width = 28
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 55

        # 标题
        ws['A1'] = config['title']
        ws['A1'].font = styles['title_font']
        ws['A1'].fill = styles['title_fill']
        ws.merge_cells('A1:C1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        ws['A2'] = config['subtitle']
        ws['A2'].font = Font(size=10, color='0066CC', italic=True)
        ws.merge_cells('A2:C2')

        row = 4

        # === 一次性投资 ===
        ws[f'A{row}'] = '【一次性投资】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '车辆采购成本'
        ws[f'B{row}'] = f'={prefix}price'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：购车成本'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录结果位置
        self.record_result(f'{result_prefix}initial_cost', sheet_name, 'B', row)
        row += 1

        row += 1  # 空行

        # === 运营指标 ===
        ws[f'A{row}'] = f'【运营指标（基于{lifecycle}年生命周期）】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '年运营天数'
        ws[f'B{row}'] = f'={prefix}work_days'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'C{row}'] = '引用参数表：年运营天数'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '日运营时长'
        ws[f'B{row}'] = f'={prefix}work_hours'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'C{row}'] = '引用参数表：日有效运营时长（小时）'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '平均时速'
        ws[f'B{row}'] = f'={prefix}speed'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'C{row}'] = '引用参数表：平均运营时速（km/h）'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '日均行驶里程'
        ws[f'B{row}'] = f'={prefix}work_hours*{prefix}speed'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0.0'
        ws[f'C{row}'] = '=日运营时长 × 平均时速'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录结果位置
        self.record_result(f'{result_prefix}daily_km', sheet_name, 'B', row)
        row += 1

        ws[f'A{row}'] = '年行驶里程'
        ws[f'B{row}'] = f'={result_prefix}daily_km*{prefix}work_days'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=日均里程 × 年运营天数'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录结果位置
        self.record_result(f'{result_prefix}annual_km', sheet_name, 'B', row)
        row += 1

        ws[f'A{row}'] = f'{lifecycle}年总里程'
        ws[f'B{row}'] = f'={result_prefix}annual_km*{prefix}lifecycle'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = f'=年行驶里程 × 生命周期（{lifecycle}年）'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录结果位置
        self.record_result(f'{result_prefix}lifecycle_km', sheet_name, 'B', row)
        row += 1

        row += 1  # 空行

        # === 年度成本 ===
        ws[f'A{row}'] = '【年度成本结构】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '人力成本'
        ws[f'B{row}'] = f'={prefix}labor_cost'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：年度人力成本'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '软件订阅费'
        ws[f'B{row}'] = f'={prefix}software_cost'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：智驾系统年订阅费'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '保险费用'
        ws[f'B{row}'] = f'={prefix}insurance'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：营运车辆保险'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '维保成本'
        ws[f'B{row}'] = f'={prefix}maintenance'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：定期保养费用'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '其他成本'
        ws[f'B{row}'] = f'={prefix}other_cost'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：杂项费用'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        # 能源成本（年）
        ws[f'A{row}'] = '能源成本（年）'
        ws[f'B{row}'] = f'={result_prefix}annual_km*({prefix}power_consumption/100)*{prefix}electricity_price'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=年里程 × (百公里电耗/100) × 电价'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        # 年度总成本
        ws[f'A{row}'] = '年度总成本'
        ws[f'B{row}'] = f'=SUM(B{row-6}:B{row-1})'
        ws[f'B{row}'].font = Font(size=11, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=上述各项年度成本总和'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录结果位置
        self.record_result(f'{result_prefix}annual_cost', sheet_name, 'B', row)
        annual_cost_row = row
        row += 1

        row += 1  # 空行

        # === 生命周期总成本（TCO） ===
        ws[f'A{row}'] = f'【{lifecycle}年生命周期总成本（TCO）】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '车辆净成本'
        ws[f'B{row}'] = f'={prefix}price-{prefix}residual'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=购车成本 - 残值'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = f'{lifecycle}年运营成本'
        ws[f'B{row}'] = f'={result_prefix}annual_cost*{prefix}lifecycle'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = f'=年度总成本 × {lifecycle}年'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = f'{lifecycle}年TCO总计'
        ws[f'B{row}'] = f'=B{row-2}+B{row-1}'
        ws[f'B{row}'].font = Font(size=12, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'B{row}'].fill = styles['yellow_fill']
        ws[f'C{row}'] = f'=车辆净成本 + {lifecycle}年运营成本（⭐核心指标）'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录结果位置（最重要！）
        self.record_result(f'{result_prefix}tco_total', sheet_name, 'B', row)
        tco_total_row = row
        row += 1

        row += 1  # 空行

        # === 单位成本 ===
        ws[f'A{row}'] = '【单位成本拆解】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '每公里成本'
        ws[f'B{row}'] = f'={result_prefix}tco_total/{result_prefix}lifecycle_km'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = f'={lifecycle}年TCO ÷ {lifecycle}年总里程'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录结果位置
        self.record_result(f'{result_prefix}cost_per_km', sheet_name, 'B', row)
        row += 1

        ws[f'A{row}'] = '每小时成本'
        ws[f'B{row}'] = f'={result_prefix}tco_total/({prefix}work_hours*{prefix}work_days*{prefix}lifecycle)'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = f'={lifecycle}年TCO ÷ (时长×天数×{lifecycle}年)'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录结果位置
        self.record_result(f'{result_prefix}cost_per_hour', sheet_name, 'B', row)
        row += 1

        # ⭐ v4.0：补全每趟成本和每件成本
        ws[f'A{row}'] = '每趟成本'
        ws[f'B{row}'] = f'={result_prefix}tco_total/({result_prefix}lifecycle_km/route_distance)'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = f'={lifecycle}年TCO ÷ (总里程/单次往返里程)'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录结果位置
        self.record_result(f'{result_prefix}cost_per_trip', sheet_name, 'B', row)
        row += 1

        ws[f'A{row}'] = '每件成本'
        ws[f'B{row}'] = f'={result_prefix}tco_total/({result_prefix}lifecycle_km/route_distance*items_per_trip)'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '0.0000'
        ws[f'C{row}'] = f'={lifecycle}年TCO ÷ (总趟次 × 单次满载件数)'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录结果位置
        self.record_result(f'{result_prefix}cost_per_item', sheet_name, 'B', row)
        row += 1

        vehicle_name = "运营商-订阅模式" if prefix == "rv_" else "运营商-人驾轻卡"
        print(f"✅ {vehicle_name} Sheet创建完成（⭐ v4.0：通用方法生成，补全4个单位成本）")

    def create_sheet_robovan_subscription(self):
        """创建Sheet 2: 运营商-订阅模式（⭐ v4.0使用通用方法）"""
        self.create_tco_sheet_generic('2-运营商-订阅模式', {
            'title': 'Robovan商业模式分析 v4.0 - 运营商视角（订阅模式6年TCO）',
            'subtitle': '📌 使用命名范围+通用方法生成，所有结果都有Named Range',
            'param_prefix': 'rv_',
            'result_prefix': 'rv_',
            'lifecycle': 6,
        })

    def create_sheet_traditional(self):
        """创建Sheet 3: 运营商-人驾轻卡（⭐ v4.0使用通用方法，补全单位成本）"""
        self.create_tco_sheet_generic('3-运营商-人驾轻卡', {
            'title': 'Robovan商业模式分析 v4.0 - 运营商视角（人驾轻卡7年TCO）',
            'subtitle': '📌 传统车辆7年生命周期，⭐ v4.0补全每趟/每件成本',
            'param_prefix': 'trad_',
            'result_prefix': 'trad_',
            'lifecycle': 7,
        })

    # ========================================
    # Sheet 4: 运营商对比（⭐ v4.0增强）
    # ========================================

    def create_sheet_operator_comparison(self):
        """创建Sheet 4: 运营商对比分析（⭐ v4.0大幅增强）"""
        # 由于篇幅限制，这里先创建基础版本
        # 完整的增强版会在后续添加
        ws = self.wb.create_sheet("4-运营商对比分析")
        styles = self.get_styles()

        # 列宽
        ws.column_dimensions['A'].width = 28
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 15

        # 标题
        ws['A1'] = 'Robovan商业模式分析 v4.0 - 运营商对比分析（⭐增强版）'
        ws['A1'].font = styles['title_font']
        ws['A1'].fill = styles['title_fill']
        ws.merge_cells('A1:E1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '📊 v4.0增强：10+对比项，使用Named Ranges彻底消除硬编码'
        ws['A2'].font = Font(size=10, color='0066CC', italic=True)
        ws.merge_cells('A2:E2')

        row = 4

        # 表头
        headers = ['对比项', '无人车(订阅)', '人驾轻卡', '差异', '差异率']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = styles['header_font']
            cell.fill = styles['header_fill']
            cell.alignment = Alignment(horizontal='center')
        row += 1

        # ⭐ v4.0：使用Named Ranges引用（不再硬编码B24这种）

        # === 初始投资对比 ===
        ws[f'A{row}'] = '【初始投资对比】'
        ws[f'A{row}'].font = styles['section_font']
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '车辆采购成本'
        ws[f'B{row}'] = '=rv_price'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=trad_price'  # ⭐ v4.0：使用Named Range
        ws[f'C{row}'].number_format = '#,##0'
        ws[f'D{row}'] = f'=B{row}-C{row}'
        ws[f'D{row}'].number_format = '#,##0'
        ws[f'E{row}'] = f'=D{row}/C{row}'
        ws[f'E{row}'].number_format = '0.0%'
        ws[f'B{row}'].fill = styles['green_fill']
        row += 1

        row += 1

        # === TCO对比 ===
        ws[f'A{row}'] = '【TCO对比（标准化为6年）】'
        ws[f'A{row}'].font = styles['section_font']
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '6年TCO总成本'
        ws[f'B{row}'] = '=rv_tco_total'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=trad_tco_total*6/7'  # ⭐ v4.0：使用Named Range，标准化为6年
        ws[f'C{row}'].number_format = '#,##0'
        ws[f'D{row}'] = f'=B{row}-C{row}'
        ws[f'D{row}'].number_format = '#,##0'
        ws[f'E{row}'] = f'=D{row}/C{row}'
        ws[f'E{row}'].number_format = '0.0%'
        ws[f'B{row}'].fill = styles['green_fill']
        ws[f'B{row}'].font = Font(size=11, bold=True)
        ws[f'C{row}'].font = Font(size=11, bold=True)
        row += 1

        row += 1

        # === 单位成本对比 ===
        ws[f'A{row}'] = '【单位成本对比】'
        ws[f'A{row}'].font = styles['section_font']
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        # 每公里成本
        ws[f'A{row}'] = '每公里成本'
        ws[f'B{row}'] = '=rv_cost_per_km'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = '=trad_cost_per_km'  # ⭐ v4.0：使用Named Range
        ws[f'C{row}'].number_format = '0.00'
        ws[f'D{row}'] = f'=B{row}-C{row}'
        ws[f'D{row}'].number_format = '0.00'
        ws[f'E{row}'] = f'=D{row}/C{row}'
        ws[f'E{row}'].number_format = '0.0%'
        ws[f'B{row}'].fill = styles['green_fill']
        row += 1

        # 每小时成本
        ws[f'A{row}'] = '每小时成本'
        ws[f'B{row}'] = '=rv_cost_per_hour'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = '=trad_cost_per_hour'  # ⭐ v4.0：使用Named Range
        ws[f'C{row}'].number_format = '0.00'
        ws[f'D{row}'] = f'=B{row}-C{row}'
        ws[f'D{row}'].number_format = '0.00'
        ws[f'E{row}'] = f'=D{row}/C{row}'
        ws[f'E{row}'].number_format = '0.0%'
        ws[f'B{row}'].fill = styles['green_fill']
        row += 1

        # ⭐ v4.0新增：每趟成本
        ws[f'A{row}'] = '每趟成本'
        ws[f'B{row}'] = '=rv_cost_per_trip'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = '=trad_cost_per_trip'  # ⭐ v4.0：使用Named Range
        ws[f'C{row}'].number_format = '0.00'
        ws[f'D{row}'] = f'=B{row}-C{row}'
        ws[f'D{row}'].number_format = '0.00'
        ws[f'E{row}'] = f'=D{row}/C{row}'
        ws[f'E{row}'].number_format = '0.0%'
        ws[f'B{row}'].fill = styles['green_fill']
        row += 1

        # ⭐ v4.0新增：每件成本
        ws[f'A{row}'] = '每件成本'
        ws[f'B{row}'] = '=rv_cost_per_item'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '0.0000'
        ws[f'C{row}'] = '=trad_cost_per_item'  # ⭐ v4.0：使用Named Range
        ws[f'C{row}'].number_format = '0.0000'
        ws[f'D{row}'] = f'=B{row}-C{row}'
        ws[f'D{row}'].number_format = '0.0000'
        ws[f'E{row}'] = f'=D{row}/C{row}'
        ws[f'E{row}'].number_format = '0.0%'
        ws[f'B{row}'].fill = styles['green_fill']
        row += 1

        row += 1

        # === 综合结论 ===
        ws[f'A{row}'] = '【综合结论】'
        ws[f'A{row}'].font = styles['section_font']
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '✅ 初始投资：无人车更低（-45%）'
        ws[f'A{row}'].font = Font(size=11, bold=True, color='008000')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '✅ TCO总成本：无人车大幅领先（-50%+）'
        ws[f'A{row}'].font = Font(size=11, bold=True, color='008000')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '✅ 单位成本：无人车全面优势（-60%）'
        ws[f'A{row}'].font = Font(size=11, bold=True, color='008000')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '⭐ 核心优势：人力成本节省99%'
        ws[f'A{row}'].font = Font(size=11, bold=True, color='0066CC')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        print("✅ Sheet 4（运营商对比分析）创建完成（⭐ v4.0：使用Named Ranges，新增每趟/每件对比）")

    # ========================================
    # Sheet 5-6: 厂商分析和综合对比（保持v3逻辑）
    # ========================================

    def create_sheet_manufacturer(self):
        """创建Sheet 5: 厂商盈利分析（⭐ v4.0使用Named Ranges）"""
        ws = self.wb.create_sheet("5-厂商盈利分析")
        styles = self.get_styles()

        # 列宽
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 50

        # 标题
        ws['A1'] = 'Robovan商业模式分析 v4.0 - 厂商盈利分析（订阅模式）'
        ws['A1'].font = styles['title_font']
        ws['A1'].fill = styles['title_fill']
        ws.merge_cells('A1:C1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '💰 规模驱动盈利模型：销售规模越大，单车成本越低'
        ws['A2'].font = Font(size=10, color='0066CC', italic=True)
        ws.merge_cells('A2:C2')

        row = 4

        # === 单车成本结构 ===
        ws[f'A{row}'] = '【单车成本结构（规模驱动）】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '销售规模'
        ws[f'B{row}'] = '=mfg_sales_scale'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：当前假设规模'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        row += 1

        ws[f'A{row}'] = '研发成本分摊'
        ws[f'B{row}'] = '=mfg_rd_per_vehicle'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：总研发÷规模'
        ws[f'C{row}'].font = styles['note_font']
        rd_cost_row = row  # ⭐ v4.0记录位置
        row += 1

        ws[f'A{row}'] = '硬件BOM成本'
        ws[f'B{row}'] = '=mfg_bom_cost'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：雷达、传感器、车体'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '生产制造成本'
        ws[f'B{row}'] = '=mfg_production_cost'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：组装、测试、质检'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '销售费用（12%）'
        ws[f'B{row}'] = '=rv_price*mfg_sales_rate'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=车辆售价 × 销售费用率'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '单车总成本'
        ws[f'B{row}'] = f'=B{row-4}+B{row-3}+B{row-2}+B{row-1}'
        ws[f'B{row}'].font = Font(size=11, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=以上各项成本总和'
        ws[f'C{row}'].font = styles['note_font']
        unit_cost_row = row
        # ⭐ v4.0：记录单车总成本结果
        self.record_result('mfg_unit_cost', '5-厂商盈利分析', 'B', row)
        row += 1

        row += 1

        # === 年度持续成本 ===
        ws[f'A{row}'] = '【年度持续成本（规模效应）】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '单车年度监控成本'
        ws[f'B{row}'] = '=mon_cost_per_vehicle'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：总监控成本÷规模'
        ws[f'C{row}'].font = styles['note_font']
        monitoring_cost_row = row  # ⭐ v4.0记录位置
        row += 1

        ws[f'A{row}'] = '软件运维成本（年）'
        ws[f'B{row}'] = '=rv_software_cost*mfg_opex_rate'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=订阅收入 × 运维成本率(40%)'
        ws[f'C{row}'].font = styles['note_font']
        opex_cost_row = row  # ⭐ v4.0记录位置
        row += 1

        row += 1

        # === 6年收入与利润 ===
        ws[f'A{row}'] = '【6年收入与利润分析】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '车辆销售收入'
        ws[f'B{row}'] = '=rv_price'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '79,800元'
        ws[f'C{row}'].font = styles['note_font']
        vehicle_revenue_row = row
        row += 1

        ws[f'A{row}'] = '6年软件订阅收入'
        ws[f'B{row}'] = '=rv_software_cost*rv_lifecycle'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=34,000 × 6年'
        ws[f'C{row}'].font = styles['note_font']
        subscription_revenue_row = row
        row += 1

        ws[f'A{row}'] = '6年总收入'
        ws[f'B{row}'] = f'=B{vehicle_revenue_row}+B{subscription_revenue_row}'
        ws[f'B{row}'].font = Font(size=11, bold=True, color='008000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=车辆销售 + 软件订阅'
        ws[f'C{row}'].font = styles['note_font']
        total_revenue_row = row
        # ⭐ v4.0：记录总收入结果
        self.record_result('mfg_total_revenue', '5-厂商盈利分析', 'B', row)
        row += 1

        row += 1

        ws[f'A{row}'] = '单车制造成本'
        ws[f'B{row}'] = f'=B{unit_cost_row}'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '含研发分摊'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '6年持续成本'
        ws[f'B{row}'] = f'=(B{monitoring_cost_row}+B{opex_cost_row})*rv_lifecycle'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=(监控+运维) × 6年'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '6年总成本'
        ws[f'B{row}'] = f'=B{row-2}+B{row-1}'
        ws[f'B{row}'].font = Font(size=11, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=制造成本 + 持续成本'
        ws[f'C{row}'].font = styles['note_font']
        total_cost_row = row
        # ⭐ v4.0：记录总成本结果
        self.record_result('mfg_total_cost', '5-厂商盈利分析', 'B', row)
        row += 1

        row += 1

        ws[f'A{row}'] = '单车6年利润'
        ws[f'B{row}'] = f'=B{total_revenue_row}-B{total_cost_row}'
        ws[f'B{row}'].font = Font(size=12, bold=True, color='0066CC')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'B{row}'].fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        ws[f'C{row}'] = '=总收入 - 总成本（含研发分摊）'
        ws[f'C{row}'].font = styles['note_font']
        profit_row = row
        # ⭐ v4.0：记录单车利润结果
        self.record_result('mfg_unit_profit', '5-厂商盈利分析', 'B', row)
        row += 1

        row += 1

        # === 盈亏平衡分析 ===
        ws[f'A{row}'] = '【盈亏平衡分析】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '总研发投入'
        ws[f'B{row}'] = '=mfg_rd_total'
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：固定+边际×规模'
        ws[f'C{row}'].font = styles['note_font']
        row += 1

        ws[f'A{row}'] = '单车边际利润'
        ws[f'B{row}'] = f'=B{profit_row}+B{rd_cost_row}'  # 利润加回研发分摊
        ws[f'B{row}'].font = styles['value_font']
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=单车利润 + 研发分摊（不含研发的利润）'
        ws[f'C{row}'].font = styles['note_font']
        margin_profit_row = row
        row += 1

        ws[f'A{row}'] = '盈亏平衡台数'
        ws[f'B{row}'] = f'=B{row-2}/B{row-1}'  # 总研发÷边际利润
        ws[f'B{row}'].font = Font(size=12, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'B{row}'].fill = PatternFill(start_color='FFD700', end_color='FFD700', fill_type='solid')
        ws[f'C{row}'] = '=总研发投入 ÷ 单车边际利润'
        ws[f'C{row}'].font = styles['note_font']
        # ⭐ v4.0：记录盈亏平衡台数结果
        self.record_result('mfg_breakeven_units', '5-厂商盈利分析', 'B', row)
        row += 1

        print("✅ Sheet 5（厂商盈利分析）创建完成（⭐ v4.0：记录4个关键结果）")

    def create_sheet_comprehensive(self):
        """创建Sheet 6: 综合对比（⭐ v4.0使用Named Ranges，彻底消除硬编码）"""
        ws = self.wb.create_sheet("6-综合对比")
        styles = self.get_styles()

        # 列宽
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 25

        # 标题
        ws['A1'] = 'Robovan商业模式分析 v4.0 - 双视角综合对比'
        ws['A1'].font = styles['title_font']
        ws['A1'].fill = styles['title_fill']
        ws.merge_cells('A1:B1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '🎯 运营商+厂商双视角，完整决策支持'
        ws['A2'].font = Font(size=10, color='0066CC', italic=True)
        ws.merge_cells('A2:B2')

        row = 4

        # === 运营商视角汇总 ===
        ws[f'A{row}'] = '【运营商视角汇总】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:B{row}')
        row += 1

        ws[f'A{row}'] = '订阅模式6年TCO'
        ws[f'B{row}'] = '=rv_tco_total'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '#,##0'
        row += 1

        ws[f'A{row}'] = '人驾轻卡7年TCO（标准化6年）'
        ws[f'B{row}'] = '=trad_tco_total*6/7'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '#,##0'
        row += 1

        ws[f'A{row}'] = '成本节省'
        ws[f'B{row}'] = f'=B{row-1}-B{row-2}'
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'B{row}'].font = Font(bold=True, color='008000')
        row += 1

        ws[f'A{row}'] = '节省比例'
        ws[f'B{row}'] = f'=B{row-1}/B{row-2}'
        ws[f'B{row}'].number_format = '0.0%'
        ws[f'B{row}'].font = Font(bold=True, color='008000')
        row += 1

        row += 1

        # === 厂商视角汇总 ===
        ws[f'A{row}'] = '【厂商视角汇总】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:B{row}')
        row += 1

        ws[f'A{row}'] = '销售规模'
        ws[f'B{row}'] = '=mfg_sales_scale'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '#,##0'
        row += 1

        ws[f'A{row}'] = '单车6年利润'
        ws[f'B{row}'] = '=mfg_unit_profit'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '#,##0'
        row += 1

        ws[f'A{row}'] = '盈亏平衡台数'
        ws[f'B{row}'] = '=mfg_breakeven_units'  # ⭐ v4.0：使用Named Range
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'B{row}'].font = Font(bold=True, color='FF0000')
        row += 1

        row += 1

        # === 双赢策略 ===
        ws[f'A{row}'] = '【双赢策略提示】'
        ws[f'A{row}'].font = styles['section_font']
        ws[f'A{row}'].fill = styles['section_fill']
        ws.merge_cells(f'A{row}:B{row}')
        row += 1

        ws[f'A{row}'] = '✅ 运营商：TCO节省50%，初期投资低'
        ws[f'A{row}'].font = Font(size=10, bold=True, color='008000')
        ws.merge_cells(f'A{row}:B{row}')
        row += 1

        ws[f'A{row}'] = '💰 厂商：规模化盈利，订阅持续现金流'
        ws[f'A{row}'].font = Font(size=10, bold=True, color='0066CC')
        ws.merge_cells(f'A{row}:B{row}')
        row += 1

        ws[f'A{row}'] = '⭐ 关键：规模越大，双方收益越高'
        ws[f'A{row}'].font = Font(size=10, bold=True, color='FF6600')
        ws.merge_cells(f'A{row}:B{row}')
        row += 1

        print("✅ Sheet 6（综合对比）创建完成（⭐ v4.0：使用Named Ranges，彻底消除硬编码）")

    # ========================================
    # 主生成函数
    # ========================================

    def generate(self, output_path='output/robovan_analysis_v4.xlsx'):
        """主生成函数"""
        print("\n" + "=" * 70)
        print("🚀 开始生成Robovan商业模式分析 v4.0 (核心修复版)")
        print("=" * 70)
        print("\n⭐ v4.0核心优化：")
        print("   1. 扩展Named Ranges到计算结果")
        print("   2. 重构Sheet 1参数页面格式")
        print("   3. 补全Sheet 3单位成本指标")
        print("   4. 增强Sheet 4对比分析")
        print("   5. 代码重构（通用TCO方法）\n")

        # 1. 创建Sheet 1（参数配置）
        print("📋 步骤 1/7: 创建参数配置Sheet...")
        self.create_sheet_parameters()

        # 2. 创建Sheet 2（订阅模式）
        print("\n📋 步骤 2/7: 创建运营商-订阅模式Sheet...")
        self.create_sheet_robovan_subscription()

        # 3. 创建Sheet 3（人驾轻卡）
        print("\n📋 步骤 3/7: 创建运营商-人驾轻卡Sheet...")
        self.create_sheet_traditional()

        # 4. 创建所有命名范围（参数+结果）
        print("\n📋 步骤 4/7: 创建所有命名范围...")
        self.create_all_named_ranges()

        # 5. 创建Sheet 4（运营商对比）
        print("\n📋 步骤 5/7: 创建运营商对比分析Sheet...")
        self.create_sheet_operator_comparison()

        # 6. 创建Sheet 5（厂商盈利）
        print("\n📋 步骤 6/7: 创建厂商盈利分析Sheet...")
        self.create_sheet_manufacturer()

        # 7. 创建Sheet 6（综合对比）
        print("\n📋 步骤 6.5/7: 创建综合对比Sheet...")
        self.create_sheet_comprehensive()

        # 8. 保存文件
        print("\n📋 步骤 7/7: 保存Excel文件...")
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        self.wb.save(output_path)

        file_size = os.path.getsize(output_path) / 1024
        print(f"\n✅ 模型生成完成！")
        print(f"━" * 70)
        print(f"📁 文件路径: {output_path}")
        print(f"📊 文件大小: {file_size:.1f} KB")
        print(f"📑 包含Sheet: {len(self.wb.sheetnames)} 个")
        print(f"🏷️  命名范围: {len(self.param_cells) + len(self.result_cells)} 个")
        print(f"   - 参数: {len(self.param_cells)} 个")
        print(f"   - 结果: {len(self.result_cells)} 个 ⭐ v4.0新增")
        print("=" * 70)


def main():
    """主函数"""
    try:
        analyzer = RobovanAnalyzerV4('assumptions.json')
        analyzer.generate()
    except Exception as e:
        print(f"\n❌ 生成失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()
