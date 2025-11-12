#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Robovan Business Model Analyzer v3.0 (Named Ranges版)
生成动态Excel模型，使用命名范围避免公式引用错位
双视角分析：运营商+厂商
"""

import json
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName


class RobovanAnalyzerV3:
    """Robovan商业模式分析器 v3.0 - 双视角动态模型（Named Ranges版）"""

    def __init__(self, config_path='assumptions.json'):
        self.config_path = config_path
        self.config = self.load_config()
        self.wb = Workbook()
        # 删除默认sheet
        if 'Sheet' in self.wb.sheetnames:
            del self.wb['Sheet']

        # 参数位置动态记录（在create_sheet_parameters中填充）
        self.param_cells = {}  # {param_name: (sheet_name, col, row)}
        self.sheet1_name = "1-参数配置"

    def load_config(self):
        """加载配置文件"""
        with open(self.config_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def record_param(self, name, col, row, sheet='1-参数配置'):
        """记录参数位置

        Args:
            name: 参数名称（用于Named Range）
            col: 列（如'B'）
            row: 行号
            sheet: Sheet名称
        """
        self.param_cells[name] = (sheet, col, row)

    def create_named_ranges(self):
        """创建所有参数的命名范围

        在Sheet 1创建完成后调用，为所有参数创建Named Ranges
        """
        print("\n✨ 创建命名范围（Named Ranges）...")
        print("━" * 50)

        count = 0
        for name, (sheet, col, row) in self.param_cells.items():
            # 构建引用字符串
            ref = f"'{sheet}'!${col}${row}"

            # 创建命名范围
            # 注意：名称中不能有空格和特殊字符，用下划线替换
            safe_name = name.replace(' ', '_').replace('(', '').replace(')', '').replace('/', '_')

            try:
                defined_name = DefinedName(safe_name, attr_text=ref)
                self.wb.defined_names[safe_name] = defined_name
                count += 1
                if count <= 10 or count % 10 == 0:  # 只打印前10个和每10个
                    print(f"  ✓ {safe_name:30} → {ref}")
            except Exception as e:
                print(f"  ✗ 创建命名范围失败: {safe_name} - {e}")

        print(f"━" * 50)
        print(f"✅ 成功创建 {count} 个命名范围\n")

    def get_named_ref(self, param_name):
        """获取参数的命名引用（用于公式）

        Args:
            param_name: 参数名称

        Returns:
            命名范围引用字符串，如 "rv_price"
        """
        safe_name = param_name.replace(' ', '_').replace('(', '').replace(')', '').replace('/', '_')
        return safe_name

    def create_sheet_parameters(self):
        """创建Sheet 1: 参数配置总表（6个区块，动态记录位置）"""
        ws = self.wb.create_sheet(self.sheet1_name)

        # === 样式定义 ===
        title_font = Font(size=16, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        section_font = Font(bold=True, color='FFFFFF', size=11)
        header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        header_font = Font(bold=True, size=10)
        input_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
        input_font = Font(size=10, bold=True, color='0000FF')
        calc_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
        calc_font = Font(size=10, italic=True, color='0066CC')

        # === 列宽设置 ===
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 30
        ws.column_dimensions['F'].width = 25

        # === 标题 ===
        ws['A1'] = 'Robovan商业模式分析 v3.0 - 参数配置（双视角+命名范围）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:F1')
        ws.row_dimensions[1].height = 30

        ws['A2'] = f"生成日期：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws['A2'].font = Font(size=10, italic=True)

        ws['A3'] = "⭐ 黄色单元格可修改，绿色单元格自动计算，其他Sheet通过命名范围实时更新！"
        ws['A3'].font = Font(size=11, color='FF0000', bold=True)

        ws['A4'] = "📊 v3.0新增：使用Named Ranges避免公式引用错位"
        ws['A4'].font = Font(size=10, color='0066CC', italic=True)

        row = 6

        # ========================================
        # A区：车辆基本参数
        # ========================================
        ws[f'A{row}'] = '【A区】车辆基本参数'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        # 表头
        headers = ['参数项', '无人车(订阅)', '人驾电动轻卡', '单位', '说明', '备注']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        # A区数据
        basic_params = [
            ('rv_vehicle_type', 'trad_vehicle_type', '车型',
             self.config['无人车']['基本信息']['车型'],
             self.config['人驾电动轻卡']['基本信息']['车型'],
             '-', '市场主流车型', False),

            ('rv_capacity_m3', 'trad_capacity_m3', '运力',
             self.config['无人车']['基本信息']['运力_立方米'],
             self.config['人驾电动轻卡']['基本信息']['运力_立方米'],
             'm³', '官方参数', True),

            ('rv_price', 'trad_price', '车辆采购成本',
             self.config['无人车']['投资成本']['车辆采购成本_元'],
             self.config['人驾电动轻卡']['投资成本']['车辆采购成本_元'],
             '元', '官网价格', True),

            ('rv_battery', 'trad_battery', '电池容量',
             self.config['无人车']['投资成本']['电池容量_kWh'],
             self.config['人驾电动轻卡']['投资成本']['电池容量_kWh'],
             'kWh', '官方参数', True),

            ('rv_range', 'trad_range', '续航里程',
             self.config['无人车']['投资成本']['续航里程_公里'],
             self.config['人驾电动轻卡']['投资成本']['续航里程_公里'],
             'km', 'CLTC工况', True),

            ('rv_lifecycle', 'trad_lifecycle', '生命周期',
             self.config['无人车']['投资成本']['生命周期_年'],
             self.config['人驾电动轻卡']['投资成本']['生命周期_年'],
             '年', '无人车6年/人驾7年', True),

            ('rv_residual', 'trad_residual', '残值',
             self.config['无人车']['投资成本']['残值_元'],
             self.config['人驾电动轻卡']['投资成本']['残值_元'],
             '元', '约购车成本38%/15%', True),
        ]

        for rv_name, trad_name, display_name, rv_val, trad_val, unit, note, is_input in basic_params:
            ws.cell(row, 1, display_name)
            cell_rv = ws.cell(row, 2, rv_val)
            cell_trad = ws.cell(row, 3, trad_val)

            # 记录参数位置
            self.record_param(rv_name, 'B', row)
            self.record_param(trad_name, 'C', row)

            if is_input and isinstance(rv_val, (int, float)):
                cell_rv.font = input_font
                cell_rv.fill = input_fill
                cell_trad.font = input_font
                cell_trad.fill = input_fill
                if isinstance(rv_val, int) and rv_val > 100:
                    cell_rv.number_format = '#,##0'
                    cell_trad.number_format = '#,##0'

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # ========================================
        # B区：运营参数
        # ========================================
        ws[f'A{row}'] = '【B区】运营参数'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        # B区数据 - 运营时间参数
        operation_params = [
            ('rv_work_days', 'trad_work_days', '年运营天数',
             self.config['无人车']['运营时间参数']['年运营天数_天'],
             self.config['人驾电动轻卡']['运营时间参数']['年运营天数_天'],
             '天', '无人车350天/人驾330天'),

            ('rv_work_hours', 'trad_work_hours', '日有效运营时长',
             self.config['无人车']['运营时间参数']['日有效运营时长_小时'],
             self.config['人驾电动轻卡']['运营时间参数']['日有效运营时长_小时'],
             '小时', '无人车10h/人驾8h'),

            ('rv_speed', 'trad_speed', '平均运营时速',
             self.config['无人车']['运营时间参数']['平均运营时速_公里每小时'],
             self.config['人驾电动轻卡']['运营时间参数']['平均运营时速_公里每小时'],
             'km/h', '无人车25/人驾28'),
        ]

        for rv_name, trad_name, display_name, rv_val, trad_val, unit, note in operation_params:
            ws.cell(row, 1, display_name)
            cell_rv = ws.cell(row, 2, rv_val)
            cell_trad = ws.cell(row, 3, trad_val)

            # 记录参数位置
            self.record_param(rv_name, 'B', row)
            self.record_param(trad_name, 'C', row)

            cell_rv.font = input_font
            cell_rv.fill = input_fill
            cell_trad.font = input_font
            cell_trad.fill = input_fill

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # B区数据 - 能源参数
        energy_params = [
            ('rv_power_consumption', 'trad_power_consumption', '百公里电耗',
             self.config['无人车']['能源参数']['百公里电耗_kWh'],
             self.config['人驾电动轻卡']['能源参数']['百公里电耗_kWh'],
             'kWh/100km', '传感器耗电'),

            ('rv_electricity_price', 'trad_electricity_price', '电价',
             self.config['无人车']['能源参数']['电价_元每kWh'],
             self.config['人驾电动轻卡']['能源参数']['电价_元每kWh'],
             '元/kWh', '工业用电'),
        ]

        for rv_name, trad_name, display_name, rv_val, trad_val, unit, note in energy_params:
            ws.cell(row, 1, display_name)
            cell_rv = ws.cell(row, 2, rv_val)
            cell_trad = ws.cell(row, 3, trad_val)

            # 记录参数位置
            self.record_param(rv_name, 'B', row)
            self.record_param(trad_name, 'C', row)

            cell_rv.font = input_font
            cell_rv.fill = input_fill
            cell_trad.font = input_font
            cell_trad.fill = input_fill

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # B区数据 - 业务场景参数
        business_params = [
            ('route_distance', None, '单次往返里程',
             self.config['业务场景参数']['单次往返里程_公里'],
             None, 'km', '仓库到配送中心'),

            ('items_per_trip', None, '单次满载件数',
             self.config['业务场景参数']['单次满载件数_件'],
             None, '件', '8m³标准包裹'),
        ]

        for rv_name, trad_name, display_name, rv_val, trad_val, unit, note in business_params:
            ws.cell(row, 1, display_name)
            cell_rv = ws.cell(row, 2, rv_val)

            # 记录参数位置
            self.record_param(rv_name, 'B', row)

            cell_rv.font = input_font
            cell_rv.fill = input_fill

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # B区数据 - 市场定价参数
        pricing_params = [
            ('price_per_km', None, '按公里收费',
             self.config['市场定价参数']['按公里收费_元每公里'],
             None, '元/km', '短途货运市场价'),

            ('price_per_item', None, '按件收费',
             self.config['市场定价参数']['按件收费_元每件'],
             None, '元/件', '城市配送标准'),
        ]

        for rv_name, trad_name, display_name, rv_val, trad_val, unit, note in pricing_params:
            ws.cell(row, 1, display_name)
            cell_rv = ws.cell(row, 2, rv_val)

            # 记录参数位置
            self.record_param(rv_name, 'B', row)

            cell_rv.font = input_font
            cell_rv.fill = input_fill

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # ========================================
        # C区：运营商年度成本参数
        # ========================================
        ws[f'A{row}'] = '【C区】运营商年度成本参数'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        # C区数据
        cost_params = [
            ('rv_labor_cost', 'trad_labor_cost', '人力成本',
             self.config['无人车']['年度成本_元']['人力成本'],
             self.config['人驾电动轻卡']['年度成本_元']['人力成本'],
             '元/年', '无人车1K/人驾110K'),

            ('rv_software_cost', 'trad_software_cost', '软件订阅费',
             self.config['无人车']['年度成本_元']['软件订阅费'],
             self.config['人驾电动轻卡']['年度成本_元'].get('软件订阅费', 0),
             '元/年', '无人车34K/人驾0'),

            ('rv_insurance', 'trad_insurance', '保险费用',
             self.config['无人车']['年度成本_元']['保险费用'],
             self.config['人驾电动轻卡']['年度成本_元']['保险费用'],
             '元/年', '营运车辆'),

            ('rv_maintenance', 'trad_maintenance', '维保成本',
             self.config['无人车']['年度成本_元']['维保成本'],
             self.config['人驾电动轻卡']['年度成本_元']['维保成本'],
             '元/年', '定期保养'),

            ('rv_other_cost', 'trad_other_cost', '其他成本',
             self.config['无人车']['年度成本_元']['其他成本'],
             self.config['人驾电动轻卡']['年度成本_元']['其他成本'],
             '元/年', '杂项费用'),
        ]

        for rv_name, trad_name, display_name, rv_val, trad_val, unit, note in cost_params:
            ws.cell(row, 1, display_name)
            cell_rv = ws.cell(row, 2, rv_val)
            cell_trad = ws.cell(row, 3, trad_val)

            # 记录参数位置
            self.record_param(rv_name, 'B', row)
            self.record_param(trad_name, 'C', row)

            cell_rv.font = input_font
            cell_rv.fill = input_fill
            cell_trad.font = input_font
            cell_trad.fill = input_fill

            if isinstance(rv_val, int) and rv_val > 100:
                cell_rv.number_format = '#,##0'
                cell_trad.number_format = '#,##0'

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # ========================================
        # D区：厂商成本参数（v3.0新增）
        # ========================================
        ws[f'A{row}'] = '【D区】厂商成本参数（v3.0新增）'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        # 简化表头
        mfg_headers = ['参数项', '参数值', '单位', '说明']
        for col_idx, header in enumerate(mfg_headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        # 销售规模
        ws.cell(row, 1, '销售规模')
        cell = ws.cell(row, 2, self.config['厂商成本参数']['销售规模']['车队规模_台'])
        self.record_param('mfg_sales_scale', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '台')
        ws.cell(row, 4, '当前假设规模')
        row += 1

        # 研发成本
        ws.cell(row, 1, '固定研发投入')
        cell = ws.cell(row, 2, self.config['厂商成本参数']['研发成本']['固定研发投入_元'])
        self.record_param('mfg_rd_fixed', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '平台和算法开发（3.8亿）')
        row += 1

        ws.cell(row, 1, '单车边际研发')
        cell = ws.cell(row, 2, self.config['厂商成本参数']['研发成本']['单车边际研发_元'])
        self.record_param('mfg_rd_marginal', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/台')
        ws.cell(row, 4, '单车场景适配（3千）')
        row += 1

        # 总研发投入（计算字段）
        ws.cell(row, 1, '总研发投入')
        total_rd_formula = f'=mfg_rd_fixed+mfg_rd_marginal*mfg_sales_scale'
        cell = ws.cell(row, 2, total_rd_formula)
        self.record_param('mfg_rd_total', 'B', row)
        cell.font = calc_font
        cell.fill = calc_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '=固定+边际×规模')
        row += 1

        # 单车研发分摊（计算字段）
        ws.cell(row, 1, '单车研发分摊')
        per_vehicle_rd_formula = f'=mfg_rd_total/mfg_sales_scale'
        cell = ws.cell(row, 2, per_vehicle_rd_formula)
        self.record_param('mfg_rd_per_vehicle', 'B', row)
        cell.font = calc_font
        cell.fill = calc_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/台')
        ws.cell(row, 4, '=总研发÷规模')
        row += 1

        row += 1  # 空行

        # 单车制造成本
        ws.cell(row, 1, '硬件BOM成本')
        cell = ws.cell(row, 2, self.config['厂商成本参数']['单车制造成本']['硬件BOM成本_元'])
        self.record_param('mfg_bom_cost', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/台')
        ws.cell(row, 4, '雷达、传感器、车体')
        row += 1

        ws.cell(row, 1, '生产制造成本')
        cell = ws.cell(row, 2, self.config['厂商成本参数']['单车制造成本']['生产制造成本_元'])
        self.record_param('mfg_production_cost', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/台')
        ws.cell(row, 4, '组装、测试、质检')
        row += 1

        row += 1  # 空行

        # 销售与服务参数
        ws.cell(row, 1, '销售费用率')
        cell = ws.cell(row, 2, self.config['厂商成本参数']['销售与服务']['销售费用率'])
        self.record_param('mfg_sales_rate', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = '0%'
        ws.cell(row, 3, '%')
        ws.cell(row, 4, '占收入比例')
        row += 1

        ws.cell(row, 1, '软件订阅毛利率')
        cell = ws.cell(row, 2, self.config['厂商成本参数']['销售与服务']['软件订阅毛利率'])
        self.record_param('mfg_subscription_margin', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = '0%'
        ws.cell(row, 3, '%')
        ws.cell(row, 4, '订阅收入-运维成本')
        row += 1

        ws.cell(row, 1, '软件运维成本率')
        cell = ws.cell(row, 2, self.config['厂商成本参数']['销售与服务']['软件运维成本率'])
        self.record_param('mfg_opex_rate', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = '0%'
        ws.cell(row, 3, '%')
        ws.cell(row, 4, '占订阅收入比例')
        row += 1

        row += 1  # 空行

        # ========================================
        # E区：远程监控参数（v3.0新增）
        # ========================================
        ws[f'A{row}'] = '【E区】远程监控参数（v3.0新增）'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        for col_idx, header in enumerate(mfg_headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        # 监控固定成本
        ws.cell(row, 1, '监控中心年固定成本')
        cell = ws.cell(row, 2, self.config['远程监控参数']['固定成本']['监控中心年固定成本_元'])
        self.record_param('mon_fixed_cost', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '租金、系统、管理人员（500万）')
        row += 1

        # 安全员年薪
        ws.cell(row, 1, '安全员年薪')
        cell = ws.cell(row, 2, self.config['远程监控参数']['变动成本']['安全员年薪_元'])
        self.record_param('mon_salary', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '一线城市监控员（12万）')
        row += 1

        # 安全员配比
        ws.cell(row, 1, '安全员配比')
        cell = ws.cell(row, 2, self.config['远程监控参数']['变动成本']['安全员配比_车每人'])
        self.record_param('mon_ratio', 'B', row)
        cell.font = input_font
        cell.fill = input_fill
        ws.cell(row, 3, '车/人')
        ws.cell(row, 4, '成熟期1:20，初期1:5，极限1:50')
        row += 1

        # 需要安全员数（计算字段）
        ws.cell(row, 1, '需要安全员数')
        cell = ws.cell(row, 2, f'=mfg_sales_scale/mon_ratio')
        self.record_param('mon_staff_count', 'B', row)
        cell.font = calc_font
        cell.fill = calc_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '人')
        ws.cell(row, 4, '=规模÷配比')
        row += 1

        # 监控中心年度总成本（计算字段）
        ws.cell(row, 1, '监控中心年度总成本')
        cell = ws.cell(row, 2, f'=mon_fixed_cost+mon_staff_count*mon_salary')
        self.record_param('mon_total_cost', 'B', row)
        cell.font = calc_font
        cell.fill = calc_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '=固定成本+人数×年薪')
        row += 1

        # 单车年度监控成本（计算字段）
        ws.cell(row, 1, '单车年度监控成本')
        cell = ws.cell(row, 2, f'=mon_total_cost/mfg_sales_scale')
        self.record_param('mon_cost_per_vehicle', 'B', row)
        cell.font = calc_font
        cell.fill = calc_fill
        cell.number_format = '#,##0'
        ws.cell(row, 3, '元/台/年')
        ws.cell(row, 4, '=总成本÷规模')
        row += 1

        row += 1  # 空行

        # ========================================
        # F区：一次性买断参数（v3.0新增）
        # ========================================
        ws[f'A{row}'] = '【F区】一次性买断模式参数（v3.0新增）'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        for col_idx, header in enumerate(mfg_headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        # 买断模式参数
        buyout_params = [
            ('buy_total_price', '车辆加软件总价',
             self.config['一次性买断模式']['投资成本']['车辆加软件总价_元'],
             '元', 'vs订阅79,800'),

            ('buy_lifecycle', '生命周期',
             self.config['一次性买断模式']['投资成本']['生命周期_年'],
             '年', '同订阅模式6年'),

            ('buy_residual', '残值',
             self.config['一次性买断模式']['投资成本']['残值_元'],
             '元', '+3000元系统价值'),

            ('buy_support_fee', '技术支持费',
             self.config['一次性买断模式']['年度成本_元']['技术支持费'],
             '元/年', '买断后单独支付'),

            ('buy_labor_cost', '人力成本',
             self.config['一次性买断模式']['年度成本_元']['人力成本'],
             '元/年', '同订阅模式'),

            ('buy_insurance', '保险费用',
             self.config['一次性买断模式']['年度成本_元']['保险费用'],
             '元/年', '同订阅模式'),

            ('buy_maintenance', '维保成本',
             self.config['一次性买断模式']['年度成本_元']['维保成本'],
             '元/年', '同订阅模式'),

            ('buy_other_cost', '其他成本',
             self.config['一次性买断模式']['年度成本_元']['其他成本'],
             '元/年', '同订阅模式'),
        ]

        for name, display_name, value, unit, note in buyout_params:
            ws.cell(row, 1, display_name)
            cell = ws.cell(row, 2, value)
            self.record_param(name, 'B', row)

            cell.font = input_font
            cell.fill = input_fill
            if isinstance(value, int) and value > 100:
                cell.number_format = '#,##0'

            ws.cell(row, 3, unit)
            ws.cell(row, 4, note)
            row += 1

        print(f"\n✅ Sheet 1创建完成，共记录 {len(self.param_cells)} 个参数位置")

    def create_sheet_robovan_subscription(self):
        """创建Sheet 2: 运营商-订阅模式（6年TCO，使用Named Ranges）"""
        ws = self.wb.create_sheet("2-运营商-订阅模式")

        # 样式
        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_font = Font(bold=True, size=11, color='FFFFFF')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, size=10)
        value_font = Font(size=10, bold=True, color='0066CC')
        note_font = Font(size=9, italic=True, color='666666')

        # 列宽
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 50

        # 标题
        ws['A1'] = 'Robovan商业模式分析 v3.0 - 运营商视角（订阅模式6年TCO）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws.merge_cells('A1:C1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '📌 本Sheet通过命名范围引用参数，确保公式准确无误'
        ws['A2'].font = Font(size=10, color='0066CC', italic=True)
        ws.merge_cells('A2:C2')

        row = 4

        # === 一次性投资 ===
        ws[f'A{row}'] = '【一次性投资】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '车辆采购成本'
        ws[f'B{row}'] = '=rv_price'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：无人车购车成本'
        ws[f'C{row}'].font = note_font
        row += 1

        row += 1  # 空行

        # === 运营指标（6年） ===
        ws[f'A{row}'] = '【运营指标（基于6年生命周期）】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '年运营天数'
        ws[f'B{row}'] = '=rv_work_days'
        ws[f'B{row}'].font = value_font
        ws[f'C{row}'] = '引用参数表：年运营天数'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '日运营时长'
        ws[f'B{row}'] = '=rv_work_hours'
        ws[f'B{row}'].font = value_font
        ws[f'C{row}'] = '引用参数表：日有效运营时长（小时）'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '平均时速'
        ws[f'B{row}'] = '=rv_speed'
        ws[f'B{row}'].font = value_font
        ws[f'C{row}'] = '引用参数表：平均运营时速（km/h）'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '日均行驶里程'
        ws[f'B{row}'] = '=rv_work_hours*rv_speed'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=日运营时长 × 平均时速'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '年行驶里程'
        ws[f'B{row}'] = f'=B{row-1}*rv_work_days'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=日均里程 × 年运营天数'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '6年总里程'
        ws[f'B{row}'] = f'=B{row-1}*rv_lifecycle'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=年行驶里程 × 生命周期（6年）'
        ws[f'C{row}'].font = note_font
        lifecycle_km_row = row
        row += 1

        row += 1  # 空行

        # === 年度成本 ===
        ws[f'A{row}'] = '【年度成本结构】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '人力成本'
        ws[f'B{row}'] = '=rv_labor_cost'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：远程安全员1:100分摊'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '软件订阅费'
        ws[f'B{row}'] = '=rv_software_cost'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：智驾系统年订阅费'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '保险费用'
        ws[f'B{row}'] = '=rv_insurance'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：营运车辆保险'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '维保成本'
        ws[f'B{row}'] = '=rv_maintenance'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：定期保养费用'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '其他成本'
        ws[f'B{row}'] = '=rv_other_cost'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：杂项费用'
        ws[f'C{row}'].font = note_font
        row += 1

        # 能源成本（年）
        ws[f'A{row}'] = '能源成本（年）'
        ws[f'B{row}'] = f'=B{lifecycle_km_row-1}*(rv_power_consumption/100)*rv_electricity_price'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=年里程 × (百公里电耗/100) × 电价'
        ws[f'C{row}'].font = note_font
        row += 1

        # 年度总成本
        ws[f'A{row}'] = '年度总成本'
        ws[f'B{row}'] = f'=SUM(B{row-6}:B{row-1})'
        ws[f'B{row}'].font = Font(size=11, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=上述各项年度成本总和'
        ws[f'C{row}'].font = note_font
        annual_cost_row = row
        row += 1

        row += 1  # 空行

        # === 6年生命周期总成本（TCO） ===
        ws[f'A{row}'] = '【6年生命周期总成本（TCO）】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '车辆净成本'
        ws[f'B{row}'] = '=rv_price-rv_residual'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=购车成本 - 残值'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '6年运营成本'
        ws[f'B{row}'] = f'=B{annual_cost_row}*rv_lifecycle'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=年度总成本 × 6年'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '6年TCO总计'
        ws[f'B{row}'] = f'=B{row-2}+B{row-1}'
        ws[f'B{row}'].font = Font(size=12, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'B{row}'].fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        ws[f'C{row}'] = '=车辆净成本 + 6年运营成本（核心指标）'
        ws[f'C{row}'].font = note_font
        tco_total_row = row
        row += 1

        row += 1  # 空行

        # === 单位成本 ===
        ws[f'A{row}'] = '【单位成本拆解】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '每公里成本'
        ws[f'B{row}'] = f'=B{tco_total_row}/B{lifecycle_km_row}'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = '=6年TCO ÷ 6年总里程'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '每小时成本'
        ws[f'B{row}'] = f'=B{tco_total_row}/(rv_work_hours*rv_work_days*rv_lifecycle)'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = '=6年TCO ÷ (时长×天数×6年)'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '每趟成本'
        ws[f'B{row}'] = f'=B{tco_total_row}/(B{lifecycle_km_row}/route_distance)'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = '=6年TCO ÷ (总里程/单次往返里程)'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '每件成本'
        ws[f'B{row}'] = f'=B{tco_total_row}/(B{lifecycle_km_row}/route_distance*items_per_trip)'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '0.0000'
        ws[f'C{row}'] = '=6年TCO ÷ (总趟次 × 单次满载件数)'
        ws[f'C{row}'].font = note_font
        row += 1

        print("✅ Sheet 2（运营商-订阅模式）创建完成")

    def create_sheet_traditional(self):
        """创建Sheet 3: 运营商-人驾轻卡（7年TCO，使用Named Ranges）"""
        ws = self.wb.create_sheet("3-运营商-人驾轻卡")

        # 样式（与Sheet 2相同）
        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_font = Font(bold=True, size=11, color='FFFFFF')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        value_font = Font(size=10, bold=True, color='0066CC')
        note_font = Font(size=9, italic=True, color='666666')

        # 列宽
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 50

        # 标题
        ws['A1'] = 'Robovan商业模式分析 v3.0 - 运营商视角（人驾轻卡7年TCO）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws.merge_cells('A1:C1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '📌 传统车辆7年生命周期，机械部件更耐用'
        ws['A2'].font = Font(size=10, color='0066CC', italic=True)
        ws.merge_cells('A2:C2')

        row = 4

        # === 一次性投资 ===
        ws[f'A{row}'] = '【一次性投资】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '车辆采购成本'
        ws[f'B{row}'] = '=trad_price'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：人驾车购车成本'
        ws[f'C{row}'].font = note_font
        row += 1

        row += 1

        # === 运营指标（7年） ===
        ws[f'A{row}'] = '【运营指标（基于7年生命周期）】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '年运营天数'
        ws[f'B{row}'] = '=trad_work_days'
        ws[f'B{row}'].font = value_font
        ws[f'C{row}'] = '引用参数表：年运营天数'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '日运营时长'
        ws[f'B{row}'] = '=trad_work_hours'
        ws[f'B{row}'].font = value_font
        ws[f'C{row}'] = '引用参数表：司机8小时工作制'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '平均时速'
        ws[f'B{row}'] = '=trad_speed'
        ws[f'B{row}'].font = value_font
        ws[f'C{row}'] = '引用参数表：经验丰富速度稍快'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '日均行驶里程'
        ws[f'B{row}'] = '=trad_work_hours*trad_speed'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=日运营时长 × 平均时速'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '年行驶里程'
        ws[f'B{row}'] = f'=B{row-1}*trad_work_days'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=日均里程 × 年运营天数'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '7年总里程'
        ws[f'B{row}'] = f'=B{row-1}*trad_lifecycle'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=年行驶里程 × 生命周期（7年）'
        ws[f'C{row}'].font = note_font
        lifecycle_km_row = row
        row += 1

        row += 1

        # === 年度成本 ===
        ws[f'A{row}'] = '【年度成本结构】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '人力成本'
        ws[f'B{row}'] = '=trad_labor_cost'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：司机年薪+社保'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '软件订阅费'
        ws[f'B{row}'] = '=trad_software_cost'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：传统车无软件订阅'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '保险费用'
        ws[f'B{row}'] = '=trad_insurance'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：营运车辆保险'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '维保成本'
        ws[f'B{row}'] = '=trad_maintenance'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：定期保养费用'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '其他成本'
        ws[f'B{row}'] = '=trad_other_cost'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：杂项费用'
        ws[f'C{row}'].font = note_font
        row += 1

        # 能源成本（年）
        ws[f'A{row}'] = '能源成本（年）'
        ws[f'B{row}'] = f'=B{lifecycle_km_row-1}*(trad_power_consumption/100)*trad_electricity_price'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=年里程 × (百公里电耗/100) × 电价'
        ws[f'C{row}'].font = note_font
        row += 1

        # 年度总成本
        ws[f'A{row}'] = '年度总成本'
        ws[f'B{row}'] = f'=SUM(B{row-6}:B{row-1})'
        ws[f'B{row}'].font = Font(size=11, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=上述各项年度成本总和'
        ws[f'C{row}'].font = note_font
        annual_cost_row = row
        row += 1

        row += 1

        # === 7年生命周期总成本（TCO） ===
        ws[f'A{row}'] = '【7年生命周期总成本（TCO）】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '车辆净成本'
        ws[f'B{row}'] = '=trad_price-trad_residual'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=购车成本 - 残值'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '7年运营成本'
        ws[f'B{row}'] = f'=B{annual_cost_row}*trad_lifecycle'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=年度总成本 × 7年'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '7年TCO总计'
        ws[f'B{row}'] = f'=B{row-2}+B{row-1}'
        ws[f'B{row}'].font = Font(size=12, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'B{row}'].fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        ws[f'C{row}'] = '=车辆净成本 + 7年运营成本（核心指标）'
        ws[f'C{row}'].font = note_font
        tco_total_row = row
        row += 1

        row += 1

        # === 单位成本 ===
        ws[f'A{row}'] = '【单位成本拆解】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '每公里成本'
        ws[f'B{row}'] = f'=B{tco_total_row}/B{lifecycle_km_row}'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = '=7年TCO ÷ 7年总里程'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '每小时成本'
        ws[f'B{row}'] = f'=B{tco_total_row}/(trad_work_hours*trad_work_days*trad_lifecycle)'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = '=7年TCO ÷ (时长×天数×7年)'
        ws[f'C{row}'].font = note_font
        row += 1

        print("✅ Sheet 3（运营商-人驾轻卡）创建完成")

    def create_sheet_operator_comparison(self):
        """创建Sheet 4: 运营商对比分析（使用Named Ranges）"""
        ws = self.wb.create_sheet("4-运营商对比分析")

        # 样式
        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_font = Font(bold=True, size=11, color='FFFFFF')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, size=10)
        value_font = Font(size=10, bold=True)
        green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

        # 列宽
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 15

        # 标题
        ws['A1'] = 'Robovan商业模式分析 v3.0 - 运营商对比分析'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws.merge_cells('A1:E1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '📊 标准化为6年周期对比，绿色=优势，红色=劣势'
        ws['A2'].font = Font(size=10, color='0066CC', italic=True)
        ws.merge_cells('A2:E2')

        row = 4

        # === 表头 ===
        headers = ['对比项', '无人车(订阅)', '人驾轻卡', '差异', '差异率']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row, col_idx, header)
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
        row += 1

        # === 初始投资对比 ===
        ws[f'A{row}'] = '【初始投资】'
        ws[f'A{row}'].font = section_font
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '车辆采购成本'
        ws[f'B{row}'] = '=rv_price'
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=trad_price'
        ws[f'C{row}'].number_format = '#,##0'
        ws[f'D{row}'] = f'=B{row}-C{row}'
        ws[f'D{row}'].number_format = '#,##0'
        ws[f'E{row}'] = f'=D{row}/C{row}'
        ws[f'E{row}'].number_format = '0.0%'
        ws[f'B{row}'].fill = green_fill
        row += 1

        row += 1

        # === TCO对比（标准化为6年） ===
        ws[f'A{row}'] = '【6年TCO对比】'
        ws[f'A{row}'].font = section_font
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '6年TCO总成本'
        ws[f'B{row}'] = "='2-运营商-订阅模式'!B24"  # 指向Sheet 2的TCO总计行
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = "='3-运营商-人驾轻卡'!B24*6/7"  # 7年标准化为6年
        ws[f'C{row}'].number_format = '#,##0'
        ws[f'D{row}'] = f'=B{row}-C{row}'
        ws[f'D{row}'].number_format = '#,##0'
        ws[f'E{row}'] = f'=D{row}/C{row}'
        ws[f'E{row}'].number_format = '0.0%'
        ws[f'B{row}'].fill = green_fill
        ws[f'B{row}'].font = Font(size=11, bold=True)
        ws[f'C{row}'].font = Font(size=11, bold=True)
        row += 1

        row += 1

        # === 单位成本对比 ===
        ws[f'A{row}'] = '【单位成本对比】'
        ws[f'A{row}'].font = section_font
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '每公里成本'
        ws[f'B{row}'] = "='2-运营商-订阅模式'!B28"  # 指向Sheet 2的每公里成本
        ws[f'B{row}'].number_format = '0.00'
        ws[f'C{row}'] = "='3-运营商-人驾轻卡'!B28"  # 指向Sheet 3的每公里成本
        ws[f'C{row}'].number_format = '0.00'
        ws[f'D{row}'] = f'=B{row}-C{row}'
        ws[f'D{row}'].number_format = '0.00'
        ws[f'E{row}'] = f'=D{row}/C{row}'
        ws[f'E{row}'].number_format = '0.0%'
        ws[f'B{row}'].fill = green_fill
        row += 1

        row += 1

        # === 综合结论 ===
        ws[f'A{row}'] = '【综合结论】'
        ws[f'A{row}'].font = section_font
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '✅ 无人车(订阅)在初始投资、TCO、单位成本全面领先'
        ws[f'A{row}'].font = Font(size=11, bold=True, color='008000')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws[f'A{row}'] = '⭐ 核心优势：人力成本几乎为零（99%节省）'
        ws[f'A{row}'].font = Font(size=11, bold=True, color='0066CC')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        print("✅ Sheet 4（运营商对比分析）创建完成")

    def create_sheet_manufacturer(self):
        """创建Sheet 5: 厂商盈利分析（使用Named Ranges）"""
        ws = self.wb.create_sheet("5-厂商盈利分析")

        # 样式
        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_font = Font(bold=True, size=11, color='FFFFFF')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        value_font = Font(size=10, bold=True, color='0066CC')
        note_font = Font(size=9, italic=True, color='666666')

        # 列宽
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 50

        # 标题
        ws['A1'] = 'Robovan商业模式分析 v3.0 - 厂商盈利分析（订阅模式）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws.merge_cells('A1:C1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '💰 规模驱动盈利模型：销售规模越大，单车成本越低'
        ws['A2'].font = Font(size=10, color='0066CC', italic=True)
        ws.merge_cells('A2:C2')

        row = 4

        # === 单车成本结构 ===
        ws[f'A{row}'] = '【单车成本结构（规模驱动）】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '销售规模'
        ws[f'B{row}'] = '=mfg_sales_scale'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：当前假设规模'
        ws[f'C{row}'].font = note_font
        row += 1

        row += 1

        ws[f'A{row}'] = '研发成本分摊'
        ws[f'B{row}'] = '=mfg_rd_per_vehicle'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：总研发÷规模'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '硬件BOM成本'
        ws[f'B{row}'] = '=mfg_bom_cost'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：雷达、传感器、车体'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '生产制造成本'
        ws[f'B{row}'] = '=mfg_production_cost'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：组装、测试、质检'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '销售费用（12%）'
        ws[f'B{row}'] = '=rv_price*mfg_sales_rate'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=车辆售价 × 销售费用率'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '单车总成本'
        ws[f'B{row}'] = f'=SUM(B{row-4}:B{row-1})'
        ws[f'B{row}'].font = Font(size=11, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=以上各项成本总和'
        ws[f'C{row}'].font = note_font
        unit_cost_row = row
        row += 1

        row += 1

        # === 年度持续成本 ===
        ws[f'A{row}'] = '【年度持续成本（规模效应）】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '单车年度监控成本'
        ws[f'B{row}'] = '=mon_cost_per_vehicle'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：总监控成本÷规模'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '软件运维成本（年）'
        ws[f'B{row}'] = '=rv_software_cost*mfg_opex_rate'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=订阅收入 × 运维成本率(40%)'
        ws[f'C{row}'].font = note_font
        row += 1

        row += 1

        # === 6年收入与利润 ===
        ws[f'A{row}'] = '【6年收入与利润分析】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '车辆销售收入'
        ws[f'B{row}'] = '=rv_price'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '79,800元'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '6年软件订阅收入'
        ws[f'B{row}'] = '=rv_software_cost*rv_lifecycle'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=34,000 × 6年'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '6年总收入'
        ws[f'B{row}'] = f'=B{row-2}+B{row-1}'
        ws[f'B{row}'].font = Font(size=11, bold=True, color='008000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=车辆销售 + 软件订阅'
        ws[f'C{row}'].font = note_font
        total_revenue_row = row
        row += 1

        row += 1

        ws[f'A{row}'] = '单车制造成本'
        ws[f'B{row}'] = f'=B{unit_cost_row}'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '含研发分摊'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '6年持续成本'
        ws[f'B{row}'] = f'=(B{unit_cost_row+10}+B{unit_cost_row+11})*rv_lifecycle'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=(监控+运维) × 6年'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '6年总成本'
        ws[f'B{row}'] = f'=B{row-2}+B{row-1}'
        ws[f'B{row}'].font = Font(size=11, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=制造成本 + 持续成本'
        ws[f'C{row}'].font = note_font
        total_cost_row = row
        row += 1

        row += 1

        ws[f'A{row}'] = '单车6年利润'
        ws[f'B{row}'] = f'=B{total_revenue_row}-B{total_cost_row}'
        ws[f'B{row}'].font = Font(size=12, bold=True, color='0066CC')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'B{row}'].fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        ws[f'C{row}'] = '=总收入 - 总成本（含研发分摊）'
        ws[f'C{row}'].font = note_font
        row += 1

        row += 1

        # === 盈亏平衡分析 ===
        ws[f'A{row}'] = '【盈亏平衡分析】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws[f'A{row}'] = '总研发投入'
        ws[f'B{row}'] = '=mfg_rd_total'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '引用参数表：固定+边际×规模'
        ws[f'C{row}'].font = note_font
        row += 1

        ws[f'A{row}'] = '单车边际利润（不含研发）'
        ws[f'B{row}'] = f'=B{total_revenue_row}-B{total_cost_row}+mfg_rd_per_vehicle*rv_lifecycle'
        ws[f'B{row}'].font = value_font
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'C{row}'] = '=6年利润 + 研发分摊（还原边际利润）'
        ws[f'C{row}'].font = note_font
        margin_row = row
        row += 1

        ws[f'A{row}'] = '盈亏平衡台数'
        ws[f'B{row}'] = f'=B{margin_row-1}/B{margin_row}'
        ws[f'B{row}'].font = Font(size=12, bold=True, color='FF0000')
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'B{row}'].fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        ws[f'C{row}'] = '=总研发 ÷ 单车边际利润（需累计销售）'
        ws[f'C{row}'].font = note_font
        row += 1

        print("✅ Sheet 5（厂商盈利分析）创建完成")

    def create_sheet_comprehensive(self):
        """创建Sheet 6: 综合对比（使用Named Ranges）"""
        ws = self.wb.create_sheet("6-综合对比")

        # 样式
        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_font = Font(bold=True, size=11, color='FFFFFF')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')

        # 列宽
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 25

        # 标题
        ws['A1'] = 'Robovan商业模式分析 v3.0 - 双视角综合对比'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws.merge_cells('A1:B1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '🎯 运营商+厂商双视角，完整决策支持'
        ws['A2'].font = Font(size=10, color='0066CC', italic=True)
        ws.merge_cells('A2:B2')

        row = 4

        # === 运营商视角汇总 ===
        ws[f'A{row}'] = '【运营商视角汇总】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:B{row}')
        row += 1

        ws[f'A{row}'] = '订阅模式6年TCO'
        ws[f'B{row}'] = "='2-运营商-订阅模式'!B24"
        ws[f'B{row}'].number_format = '#,##0'
        row += 1

        ws[f'A{row}'] = '人驾轻卡7年TCO（标准化6年）'
        ws[f'B{row}'] = "='3-运营商-人驾轻卡'!B24*6/7"
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
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:B{row}')
        row += 1

        ws[f'A{row}'] = '销售规模'
        ws[f'B{row}'] = '=mfg_sales_scale'
        ws[f'B{row}'].number_format = '#,##0'
        row += 1

        ws[f'A{row}'] = '单车6年利润'
        ws[f'B{row}'] = "='5-厂商盈利分析'!B29"  # 指向Sheet 5的单车利润
        ws[f'B{row}'].number_format = '#,##0'
        row += 1

        ws[f'A{row}'] = '盈亏平衡台数'
        ws[f'B{row}'] = "='5-厂商盈利分析'!B34"  # 指向Sheet 5的盈亏平衡
        ws[f'B{row}'].number_format = '#,##0'
        ws[f'B{row}'].font = Font(bold=True, color='FF0000')
        row += 1

        row += 1

        # === 双赢策略 ===
        ws[f'A{row}'] = '【双赢策略提示】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
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

        print("✅ Sheet 6（综合对比）创建完成")

    def generate(self, output_path='output/robovan_analysis_v3.xlsx'):
        """主生成函数"""
        print("\n" + "=" * 60)
        print("🚀 开始生成Robovan商业模式分析 v3.0 (Named Ranges版)")
        print("=" * 60)

        # 1. 创建Sheet 1（参数配置）并记录位置
        print("\n📋 步骤 1/7: 创建参数配置Sheet...")
        self.create_sheet_parameters()

        # 2. 创建所有命名范围
        print("\n📋 步骤 2/7: 创建命名范围...")
        self.create_named_ranges()

        # 3. 创建Sheet 2（订阅模式）
        print("\n📋 步骤 3/7: 创建运营商-订阅模式Sheet...")
        self.create_sheet_robovan_subscription()

        # 4. 创建Sheet 3（人驾轻卡）
        print("\n📋 步骤 4/7: 创建运营商-人驾轻卡Sheet...")
        self.create_sheet_traditional()

        # 5. 创建Sheet 4（运营商对比）
        print("\n📋 步骤 5/7: 创建运营商对比分析Sheet...")
        self.create_sheet_operator_comparison()

        # 6. 创建Sheet 5（厂商盈利）
        print("\n📋 步骤 6/7: 创建厂商盈利分析Sheet...")
        self.create_sheet_manufacturer()

        # 7. 创建Sheet 6（综合对比）
        print("\n📋 步骤 6.5/7: 创建综合对比Sheet...")
        self.create_sheet_comprehensive()

        # 7. 保存文件
        print("\n📋 步骤 7/7: 保存Excel文件...")
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        self.wb.save(output_path)

        file_size = os.path.getsize(output_path) / 1024
        print(f"\n✅ 模型生成完成！")
        print(f"📁 文件路径: {output_path}")
        print(f"📊 文件大小: {file_size:.1f} KB")
        print(f"📑 包含Sheet: {len(self.wb.sheetnames)} 个")
        print(f"🏷️  命名范围: {len(self.param_cells)} 个")
        print("=" * 60)


def main():
    """主函数"""
    try:
        analyzer = RobovanAnalyzerV3('assumptions.json')
        analyzer.generate()
    except Exception as e:
        print(f"\n❌ 生成失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()
