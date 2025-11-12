#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Robovan商业模式分析 - 动态Excel模型生成器

功能：生成公式驱动的动态Excel模型，修改参数后自动重新计算
版本：2.0（动态模型版）
更新日期：2025-01-12
"""

import json
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


class RobovanAnalyzerDynamic:
    """Robovan商业模式分析器 - 动态模型版"""

    def __init__(self, config_path='assumptions.json'):
        """初始化分析器"""
        self.config_path = config_path
        self.config = self.load_config()
        self.wb = Workbook()
        self.wb.remove(self.wb.active)  # 删除默认sheet

        # 定义Sheet1参数单元格位置映射
        self.init_cell_mapping()

    def load_config(self):
        """加载配置文件"""
        with open(self.config_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def init_cell_mapping(self):
        """初始化参数单元格位置映射"""
        # Sheet1的参数位置（行号）
        self.param_rows = {
            # 基本信息区（从第7行开始）
            '车型': 7,
            '运力': 8,
            '购车成本': 9,
            '电池容量': 10,
            '续航里程': 11,
            '生命周期': 12,
            '残值': 13,

            # 运营时间区（从第17行开始）
            '年运营天数': 17,
            '日运营时长': 18,
            '平均时速': 19,

            # 年度成本区（从第23行开始）
            '人力成本': 23,
            '软件订阅': 24,
            '保险费用': 25,
            '维保成本': 26,
            '其他成本': 27,

            # 能源参数区（从第31行开始）
            '百公里电耗': 31,
            '电价': 32,

            # 业务场景区（从第36行开始）
            '单次往返里程': 36,
            '单次满载件数': 37,

            # 市场定价区（从第41行开始）
            '市场定价_公里': 41,
            '市场定价_件': 42,
        }

        # 列映射：B列=无人车，C列=人驾
        self.col_rv = 'B'   # Robovan
        self.col_tr = 'C'   # Traditional

    def get_param_cell(self, param_name, vehicle_type='无人车'):
        """获取参数单元格地址"""
        row = self.param_rows.get(param_name)
        if row is None:
            raise ValueError(f"未找到参数: {param_name}")

        col = self.col_rv if vehicle_type == '无人车' else self.col_tr
        return f"'{self.sheet1_name}'!{col}{row}"

    def create_sheet_parameters(self):
        """创建Sheet1: 关键假设参数（输入区）"""
        ws = self.wb.create_sheet("1-参数配置")
        self.sheet1_name = "1-参数配置"

        # 样式定义
        title_font = Font(size=16, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        section_font = Font(bold=True, color='FFFFFF', size=11)
        header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        header_font = Font(bold=True, size=10)
        input_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')  # 黄色=输入
        input_font = Font(size=10, bold=True, color='0000FF')  # 蓝色字体

        # 标题
        ws['A1'] = 'Robovan商业模式分析 - 参数配置表（动态模型）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:E1')
        ws.row_dimensions[1].height = 30

        # 说明
        ws['A2'] = f"生成日期：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws['A2'].font = Font(size=10, italic=True)
        ws['A3'] = "⭐ 黄色单元格可直接修改，其他Sheet会自动更新计算结果！"
        ws['A3'].font = Font(size=11, color='FF0000', bold=True)

        row = 5

        # 【A区】基本信息
        ws[f'A{row}'] = '【A区】车辆基本参数'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        # 表头
        ws.cell(row, 1, '参数项').font = header_font
        ws.cell(row, 1).fill = header_fill
        ws.cell(row, 2, '无人车').font = header_font
        ws.cell(row, 2).fill = header_fill
        ws.cell(row, 3, '人驾电动轻卡').font = header_font
        ws.cell(row, 3).fill = header_fill
        ws.cell(row, 4, '单位').font = header_font
        ws.cell(row, 4).fill = header_fill
        ws.cell(row, 5, '说明').font = header_font
        ws.cell(row, 5).fill = header_fill
        row += 1

        # 基本信息数据
        basic_data = [
            ('车型',
             self.config['无人车']['基本信息']['车型'],
             self.config['人驾电动轻卡']['基本信息']['车型'],
             '-', '市场主流车型', False),
            ('运力',
             self.config['无人车']['基本信息']['运力_立方米'],
             self.config['人驾电动轻卡']['基本信息']['运力_立方米'],
             'm³', '官方参数', True),
            ('车辆采购成本',
             self.config['无人车']['投资成本']['车辆采购成本_元'],
             self.config['人驾电动轻卡']['投资成本']['车辆采购成本_元'],
             '元', '官网价格', True),
            ('电池容量',
             self.config['无人车']['投资成本']['电池容量_kWh'],
             self.config['人驾电动轻卡']['投资成本']['电池容量_kWh'],
             'kWh', '官方参数', True),
            ('续航里程',
             self.config['无人车']['投资成本']['续航里程_公里'],
             self.config['人驾电动轻卡']['投资成本']['续航里程_公里'],
             'km', 'CLTC工况', True),
            ('生命周期',
             self.config['无人车']['投资成本']['生命周期_年'],
             self.config['人驾电动轻卡']['投资成本']['生命周期_年'],
             '年', '技术迭代快5年；传统车7年', True),
            ('残值',
             self.config['无人车']['投资成本']['残值_元'],
             self.config['人驾电动轻卡']['投资成本']['残值_元'],
             '元', '约购车成本38%/15%', True),
        ]

        for param_name, rv_val, tr_val, unit, note, is_input in basic_data:
            ws.cell(row, 1, param_name)

            cell_rv = ws.cell(row, 2, rv_val)
            cell_tr = ws.cell(row, 3, tr_val)

            if is_input and isinstance(rv_val, (int, float)):
                cell_rv.font = input_font
                cell_rv.fill = input_fill
                cell_tr.font = input_font
                cell_tr.fill = input_fill

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1

        # 【B区】运营时间参数
        ws[f'A{row}'] = '【B区】运营时间参数'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        # 表头
        for col, header in enumerate(['参数项', '无人车', '人驾电动轻卡', '单位', '说明'], 1):
            cell = ws.cell(row, col, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        operation_data = [
            ('年运营天数',
             self.config['无人车']['运营时间参数']['年运营天数_天'],
             self.config['人驾电动轻卡']['运营时间参数']['年运营天数_天'],
             '天', '少维护vs法定假日'),
            ('日有效运营时长',
             self.config['无人车']['运营时间参数']['日有效运营时长_小时'],
             self.config['人驾电动轻卡']['运营时间参数']['日有效运营时长_小时'],
             '小时', '夜间受限vs8小时工作制'),
            ('平均运营时速',
             self.config['无人车']['运营时间参数']['平均运营时速_公里每小时'],
             self.config['人驾电动轻卡']['运营时间参数']['平均运营时速_公里每小时'],
             'km/h', '保守驾驶vs经验丰富'),
        ]

        for param_name, rv_val, tr_val, unit, note in operation_data:
            ws.cell(row, 1, param_name)

            cell_rv = ws.cell(row, 2, rv_val)
            cell_tr = ws.cell(row, 3, tr_val)
            cell_rv.font = input_font
            cell_rv.fill = input_fill
            cell_tr.font = input_font
            cell_tr.fill = input_fill

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1

        # 【C区】年度成本参数
        ws[f'A{row}'] = '【C区】年度成本参数'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        # 表头
        for col, header in enumerate(['成本项', '无人车', '人驾电动轻卡', '单位', '说明'], 1):
            cell = ws.cell(row, col, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        cost_data = [
            ('人力成本',
             self.config['无人车']['年度成本_元']['人力成本'],
             self.config['人驾电动轻卡']['年度成本_元']['人力成本'],
             '元/年', '1:100分摊 vs 司机+社保'),
            ('软件订阅费',
             self.config['无人车']['年度成本_元']['软件订阅费'],
             self.config['人驾电动轻卡']['年度成本_元']['软件订阅费'],
             '元/年', '8500元/季度×4'),
            ('保险费用',
             self.config['无人车']['年度成本_元']['保险费用'],
             self.config['人驾电动轻卡']['年度成本_元']['保险费用'],
             '元/年', '风险低vs营运险'),
            ('维保成本',
             self.config['无人车']['年度成本_元']['维保成本'],
             self.config['人驾电动轻卡']['年度成本_元']['维保成本'],
             '元/年', '传感器贵vs保养成熟'),
            ('其他成本',
             self.config['无人车']['年度成本_元']['其他成本'],
             self.config['人驾电动轻卡']['年度成本_元']['其他成本'],
             '元/年', '停车、过路费等'),
        ]

        for param_name, rv_val, tr_val, unit, note in cost_data:
            ws.cell(row, 1, param_name)

            cell_rv = ws.cell(row, 2, rv_val)
            cell_tr = ws.cell(row, 3, tr_val)
            cell_rv.font = input_font
            cell_rv.fill = input_fill
            cell_tr.font = input_font
            cell_tr.fill = input_fill

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1

        # 【D区】能源参数
        ws[f'A{row}'] = '【D区】能源参数'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        # 表头
        for col, header in enumerate(['参数项', '无人车', '人驾电动轻卡', '单位', '说明'], 1):
            cell = ws.cell(row, col, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        energy_data = [
            ('百公里电耗',
             self.config['无人车']['能源参数']['百公里电耗_kWh'],
             self.config['人驾电动轻卡']['能源参数']['百公里电耗_kWh'],
             'kWh', '传感器耗电vs行业参考'),
            ('电价标准',
             self.config['无人车']['能源参数']['电价_元每kWh'],
             self.config['人驾电动轻卡']['能源参数']['电价_元每kWh'],
             '元/kWh', '工业用电峰谷混合'),
        ]

        for param_name, rv_val, tr_val, unit, note in energy_data:
            ws.cell(row, 1, param_name)

            cell_rv = ws.cell(row, 2, rv_val)
            cell_tr = ws.cell(row, 3, tr_val)
            cell_rv.font = input_font
            cell_rv.fill = input_fill
            cell_tr.font = input_font
            cell_tr.fill = input_fill

            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1

        # 【E区】业务场景
        ws[f'A{row}'] = '【E区】业务场景参数'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws.cell(row, 1, '单次往返里程')
        cell = ws.cell(row, 2, self.config['业务场景参数']['单次往返里程_公里'])
        cell.font = input_font
        cell.fill = input_fill
        ws.cell(row, 3, '公里')
        ws.cell(row, 4, '仓储物流场景')
        row += 1

        ws.cell(row, 1, '单次满载件数')
        cell = ws.cell(row, 2, self.config['业务场景参数']['单次满载件数_件'])
        cell.font = input_font
        cell.fill = input_fill
        ws.cell(row, 3, '件')
        ws.cell(row, 4, '8m³标准包裹')
        row += 1

        row += 1

        # 【F区】市场定价
        ws[f'A{row}'] = '【F区】市场定价参数（⚠️ 外部市场价，非成本！）'
        ws[f'A{row}'].font = Font(bold=True, size=11, color='FF0000')
        ws[f'A{row}'].fill = PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws.cell(row, 1, '按公里收费')
        cell = ws.cell(row, 2, self.config['市场定价参数']['按公里收费_元每公里'])
        cell.font = input_font
        cell.fill = input_fill
        ws.cell(row, 3, '元/km')
        ws.cell(row, 4, '短途货运市场价')
        row += 1

        ws.cell(row, 1, '按件收费')
        cell = ws.cell(row, 2, self.config['市场定价参数']['按件收费_元每件'])
        cell.font = input_font
        cell.fill = input_fill
        ws.cell(row, 3, '元/件')
        ws.cell(row, 4, '城市配送标准')

        # 设置列宽
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 40

        return ws

    def create_sheet_robovan_dynamic(self):
        """创建Sheet2: 无人车TCO分析（公式驱动）"""
        ws = self.wb.create_sheet("2-无人车TCO")

        # 样式定义
        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        section_font = Font(bold=True, color='FFFFFF', size=11)
        label_font = Font(size=10)
        formula_font = Font(size=10, italic=True, color='0066CC')  # 公式用斜体蓝色
        highlight_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')

        # 标题
        ws['A1'] = '无人车模式 - 5年TCO分析（动态模型）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:D1')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '💡 本Sheet所有数据自动计算，修改参数表即可更新'
        ws['A2'].font = Font(size=10, italic=True, color='008000')

        row = 4

        # 一、一次性投资
        ws[f'A{row}'] = '一、一次性投资'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        ws.cell(row, 1, '车辆采购成本').font = label_font
        ws.cell(row, 2, f"={self.get_param_cell('购车成本', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 2

        # 二、运营指标计算（中间计算区）
        ws[f'A{row}'] = '二、运营指标（自动计算）'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        # 定义命名区域方便后续引用
        calc_start_row = row

        ws.cell(row, 1, '日均行驶里程')
        daily_km_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('日运营时长', '无人车')}*{self.get_param_cell('平均时速', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0.0'
        ws.cell(row, 3, 'km/天')
        row += 1

        ws.cell(row, 1, '年行驶里程')
        yearly_km_cell = f'B{row}'
        ws.cell(row, 2, f"=B{row-1}*{self.get_param_cell('年运营天数', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, 'km/年')
        row += 1

        ws.cell(row, 1, '5年总里程')
        lifecycle_km_cell = f'B{row}'
        ws.cell(row, 2, f"=B{row-1}*{self.get_param_cell('生命周期', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, 'km')
        row += 1

        ws.cell(row, 1, '每公里电费')
        cost_per_km_energy_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('百公里电耗', '无人车')}/100*{self.get_param_cell('电价', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '0.000'
        ws.cell(row, 3, '元/km')
        row += 2

        # 三、年度成本结构
        ws[f'A{row}'] = '三、年度成本结构'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        year_cost_start_row = row

        ws.cell(row, 1, '【固定成本】').font = Font(bold=True)
        row += 1

        # 人力成本
        ws.cell(row, 1, '  人力成本')
        labor_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('人力成本', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 1

        # 软件订阅
        ws.cell(row, 1, '  软件订阅费')
        software_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('软件订阅', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 1

        # 保险
        ws.cell(row, 1, '  保险费用')
        insurance_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('保险费用', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 1

        # 固定成本小计
        ws.cell(row, 1, '固定成本小计').font = Font(bold=True)
        fixed_cost_year_cell = f'B{row}'
        ws.cell(row, 2, f"=SUM(B{year_cost_start_row+1}:B{row-1})")
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 2

        ws.cell(row, 1, '【可变成本】').font = Font(bold=True)
        row += 1

        variable_cost_start_row = row

        # 能源成本
        ws.cell(row, 1, '  能源成本')
        energy_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={yearly_km_cell}*{cost_per_km_energy_cell}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 1

        # 维保
        ws.cell(row, 1, '  维保成本')
        maintenance_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('维保成本', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 1

        # 其他
        ws.cell(row, 1, '  其他成本')
        other_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('其他成本', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 1

        # 可变成本小计
        ws.cell(row, 1, '可变成本小计').font = Font(bold=True)
        variable_cost_year_cell = f'B{row}'
        ws.cell(row, 2, f"=SUM(B{variable_cost_start_row}:B{row-1})")
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 2

        # 年度总成本
        ws.cell(row, 1, '年度总成本').font = Font(bold=True, size=11)
        total_cost_year_cell = f'B{row}'
        ws.cell(row, 2, f"={fixed_cost_year_cell}+{variable_cost_year_cell}")
        ws.cell(row, 2).font = Font(bold=True, size=11, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        ws.cell(row, 3, '元/年')
        row += 2

        # 四、5年TCO
        ws[f'A{row}'] = '四、5年总拥有成本(TCO)'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        tco_start_row = row

        ws.cell(row, 1, '5年固定成本')
        lifecycle_fixed_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={fixed_cost_year_cell}*{self.get_param_cell('生命周期', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 1

        ws.cell(row, 1, '5年可变成本')
        lifecycle_variable_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={variable_cost_year_cell}*{self.get_param_cell('生命周期', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 1

        ws.cell(row, 1, '车辆采购成本')
        ws.cell(row, 2, f"={self.get_param_cell('购车成本', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 1

        ws.cell(row, 1, '减：残值')
        ws.cell(row, 2, f"=-{self.get_param_cell('残值', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 1

        ws.cell(row, 1, '净车辆成本')
        net_vehicle_cost_cell = f'B{row}'
        ws.cell(row, 2, f"=B{row-2}+B{row-1}")
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 2

        ws.cell(row, 1, '5年TCO总成本').font = Font(bold=True, size=12, color='FF0000')
        tco_total_cell = f'B{row}'
        ws.cell(row, 2, f"={lifecycle_fixed_cost_cell}+{lifecycle_variable_cost_cell}+{net_vehicle_cost_cell}")
        ws.cell(row, 2).font = Font(bold=True, size=12, italic=True, color='FF0000')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        ws.cell(row, 3, '元')
        row += 2

        # 五、业务量指标
        ws[f'A{row}'] = '五、业务量指标（5年）'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        ws.cell(row, 1, '总里程')
        ws.cell(row, 2, f"={lifecycle_km_cell}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '公里')
        row += 1

        ws.cell(row, 1, '总时长')
        lifecycle_hours_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('日运营时长', '无人车')}*{self.get_param_cell('年运营天数', '无人车')}*{self.get_param_cell('生命周期', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '小时')
        row += 1

        ws.cell(row, 1, '总趟次')
        lifecycle_trips_cell = f'B{row}'
        ws.cell(row, 2, f"={lifecycle_km_cell}/{self.get_param_cell('单次往返里程', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '趟')
        row += 1

        ws.cell(row, 1, '总件数')
        lifecycle_items_cell = f'B{row}'
        ws.cell(row, 2, f"={lifecycle_trips_cell}*{self.get_param_cell('单次满载件数', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '件')
        row += 2

        # 六、单位成本
        ws[f'A{row}'] = '六、单位成本拆解（⚠️ 成本，非收费）'
        ws[f'A{row}'].font = Font(bold=True, size=11, color='FF0000')
        ws[f'A{row}'].fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        unit_costs = [
            ('每公里成本', f"={tco_total_cell}/{lifecycle_km_cell}", '0.00', '元/km'),
            ('每小时成本', f"={tco_total_cell}/{lifecycle_hours_cell}", '0.00', '元/h'),
            ('每趟次成本', f"={tco_total_cell}/{lifecycle_trips_cell}", '0.00', '元/趟'),
            ('每件成本', f"={tco_total_cell}/{lifecycle_items_cell}", '0.000', '元/件'),
            ('', '', '', ''),
            ('每天成本', f"={tco_total_cell}/({self.get_param_cell('年运营天数', '无人车')}*{self.get_param_cell('生命周期', '无人车')})", '0.00', '元/天'),
            ('每月成本', f"={tco_total_cell}/({self.get_param_cell('生命周期', '无人车')}*12)", '0.00', '元/月'),
            ('每年成本', f"={total_cost_year_cell}", '#,##0', '元/年'),
        ]

        for label, formula, num_format, unit in unit_costs:
            if label:
                ws.cell(row, 1, label)
                ws.cell(row, 2, formula)
                ws.cell(row, 2).font = formula_font
                ws.cell(row, 2).number_format = num_format
                ws.cell(row, 3, unit)
            row += 1

        # 设置列宽
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 15

        # 保存一些关键单元格地址供Sheet4使用
        self.rv_cells = {
            'tco_total': tco_total_cell,
            'total_cost_year': total_cost_year_cell,
            'lifecycle_km': lifecycle_km_cell,
            'lifecycle_items': lifecycle_items_cell,
            'daily_km': daily_km_cell,
            'yearly_km': yearly_km_cell,
            'net_vehicle_cost': net_vehicle_cost_cell,
            'labor_cost': labor_cost_cell,
            'software_cost': software_cost_cell,
            'insurance_cost': insurance_cost_cell,
            'energy_cost': energy_cost_cell,
            'maintenance_cost': maintenance_cost_cell,
            'other_cost': other_cost_cell,
        }

        return ws

    def create_sheet_traditional_dynamic(self):
        """创建Sheet3: 人驾轻卡TCO分析（公式驱动）"""
        ws = self.wb.create_sheet("3-人驾轻卡TCO")

        # 样式（复用）
        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        section_font = Font(bold=True, color='FFFFFF', size=11)
        label_font = Font(size=10)
        formula_font = Font(size=10, italic=True, color='0066CC')
        highlight_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')

        # 标题
        ws['A1'] = '人驾电动轻卡模式 - 5年TCO分析（动态模型）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:D1')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '💡 本Sheet所有数据自动计算，修改参数表即可更新'
        ws['A2'].font = Font(size=10, italic=True, color='008000')
        ws['A3'] = '注：实际生命周期7年，为对比统一按5年计算'
        ws['A3'].font = Font(size=10, italic=True, color='FF0000')

        row = 5

        # 一、一次性投资
        ws[f'A{row}'] = '一、一次性投资'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        ws.cell(row, 1, '车辆采购成本').font = label_font
        ws.cell(row, 2, f"={self.get_param_cell('购车成本', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 2

        # 二、运营指标
        ws[f'A{row}'] = '二、运营指标（自动计算）'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        ws.cell(row, 1, '日均行驶里程')
        daily_km_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('日运营时长', '人驾')}*{self.get_param_cell('平均时速', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0.0'
        ws.cell(row, 3, 'km/天')
        row += 1

        ws.cell(row, 1, '年行驶里程')
        yearly_km_cell = f'B{row}'
        ws.cell(row, 2, f"=B{row-1}*{self.get_param_cell('年运营天数', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, 'km/年')
        row += 1

        ws.cell(row, 1, '5年总里程')
        lifecycle_km_cell = f'B{row}'
        ws.cell(row, 2, f"=B{row-1}*5")  # 固定5年对比
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, 'km')
        row += 1

        ws.cell(row, 1, '每公里电费')
        cost_per_km_energy_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('百公里电耗', '人驾')}/100*{self.get_param_cell('电价', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '0.000'
        ws.cell(row, 3, '元/km')
        row += 2

        # 三、年度成本
        ws[f'A{row}'] = '三、年度成本结构'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        year_cost_start_row = row

        ws.cell(row, 1, '【固定成本】').font = Font(bold=True)
        row += 1

        ws.cell(row, 1, '  人力成本')
        labor_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('人力成本', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年 ⭐')
        row += 1

        ws.cell(row, 1, '  保险费用')
        insurance_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('保险费用', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 1

        ws.cell(row, 1, '固定成本小计').font = Font(bold=True)
        fixed_cost_year_cell = f'B{row}'
        ws.cell(row, 2, f"=SUM(B{year_cost_start_row+1}:B{row-1})")
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 2

        ws.cell(row, 1, '【可变成本】').font = Font(bold=True)
        row += 1

        variable_cost_start_row = row

        ws.cell(row, 1, '  能源成本')
        energy_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={yearly_km_cell}*{cost_per_km_energy_cell}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 1

        ws.cell(row, 1, '  维保成本')
        maintenance_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('维保成本', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 1

        ws.cell(row, 1, '  其他成本')
        other_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('其他成本', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 1

        ws.cell(row, 1, '可变成本小计').font = Font(bold=True)
        variable_cost_year_cell = f'B{row}'
        ws.cell(row, 2, f"=SUM(B{variable_cost_start_row}:B{row-1})")
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        row += 2

        ws.cell(row, 1, '年度总成本').font = Font(bold=True, size=11)
        total_cost_year_cell = f'B{row}'
        ws.cell(row, 2, f"={fixed_cost_year_cell}+{variable_cost_year_cell}")
        ws.cell(row, 2).font = Font(bold=True, size=11, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        ws.cell(row, 3, '元/年')
        row += 2

        # 四、5年TCO
        ws[f'A{row}'] = '四、5年总拥有成本(TCO)'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        ws.cell(row, 1, '5年固定成本')
        lifecycle_fixed_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={fixed_cost_year_cell}*5")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 1

        ws.cell(row, 1, '5年可变成本')
        lifecycle_variable_cost_cell = f'B{row}'
        ws.cell(row, 2, f"={variable_cost_year_cell}*5")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 1

        ws.cell(row, 1, '车辆采购成本')
        ws.cell(row, 2, f"={self.get_param_cell('购车成本', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 1

        ws.cell(row, 1, '减：残值')
        ws.cell(row, 2, f"=-{self.get_param_cell('残值', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 1

        ws.cell(row, 1, '净车辆成本')
        net_vehicle_cost_cell = f'B{row}'
        ws.cell(row, 2, f"=B{row-2}+B{row-1}")
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        row += 2

        ws.cell(row, 1, '5年TCO总成本').font = Font(bold=True, size=12, color='FF0000')
        tco_total_cell = f'B{row}'
        ws.cell(row, 2, f"={lifecycle_fixed_cost_cell}+{lifecycle_variable_cost_cell}+{net_vehicle_cost_cell}")
        ws.cell(row, 2).font = Font(bold=True, size=12, italic=True, color='FF0000')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        ws.cell(row, 3, '元')
        row += 2

        # 五、业务量指标
        ws[f'A{row}'] = '五、业务量指标（5年）'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        ws.cell(row, 1, '总里程')
        ws.cell(row, 2, f"={lifecycle_km_cell}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '公里')
        row += 1

        ws.cell(row, 1, '总时长')
        lifecycle_hours_cell = f'B{row}'
        ws.cell(row, 2, f"={self.get_param_cell('日运营时长', '人驾')}*{self.get_param_cell('年运营天数', '人驾')}*5")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '小时')
        row += 1

        ws.cell(row, 1, '总趟次')
        lifecycle_trips_cell = f'B{row}'
        ws.cell(row, 2, f"={lifecycle_km_cell}/{self.get_param_cell('单次往返里程', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '趟')
        row += 1

        ws.cell(row, 1, '总件数')
        lifecycle_items_cell = f'B{row}'
        ws.cell(row, 2, f"={lifecycle_trips_cell}*{self.get_param_cell('单次满载件数', '人驾')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '件')
        row += 2

        # 六、单位成本
        ws[f'A{row}'] = '六、单位成本拆解（⚠️ 成本，非收费）'
        ws[f'A{row}'].font = Font(bold=True, size=11, color='FF0000')
        ws[f'A{row}'].fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        unit_costs = [
            ('每公里成本', f"={tco_total_cell}/{lifecycle_km_cell}", '0.00', '元/km'),
            ('每小时成本', f"={tco_total_cell}/{lifecycle_hours_cell}", '0.00', '元/h'),
            ('每趟次成本', f"={tco_total_cell}/{lifecycle_trips_cell}", '0.00', '元/趟'),
            ('每件成本', f"={tco_total_cell}/{lifecycle_items_cell}", '0.000', '元/件'),
            ('', '', '', ''),
            ('每天成本', f"={tco_total_cell}/({self.get_param_cell('年运营天数', '人驾')}*5)", '0.00', '元/天'),
            ('每月成本', f"={tco_total_cell}/60", '0.00', '元/月'),
            ('每年成本', f"={total_cost_year_cell}", '#,##0', '元/年'),
        ]

        for label, formula, num_format, unit in unit_costs:
            if label:
                ws.cell(row, 1, label)
                ws.cell(row, 2, formula)
                ws.cell(row, 2).font = formula_font
                ws.cell(row, 2).number_format = num_format
                ws.cell(row, 3, unit)
            row += 1

        # 设置列宽
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 12

        # 保存关键单元格地址
        self.tr_cells = {
            'tco_total': tco_total_cell,
            'total_cost_year': total_cost_year_cell,
            'lifecycle_km': lifecycle_km_cell,
            'lifecycle_items': lifecycle_items_cell,
            'daily_km': daily_km_cell,
            'yearly_km': yearly_km_cell,
            'net_vehicle_cost': net_vehicle_cost_cell,
            'labor_cost': labor_cost_cell,
            'insurance_cost': insurance_cost_cell,
            'energy_cost': energy_cost_cell,
            'maintenance_cost': maintenance_cost_cell,
            'other_cost': other_cost_cell,
        }

        return ws

    def create_sheet_comparison_dynamic(self):
        """创建Sheet4: 成本对比分析（公式驱动）"""
        ws = self.wb.create_sheet("4-成本对比")

        # 样式
        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        section_font = Font(bold=True, color='FFFFFF', size=11)
        header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        header_font = Font(bold=True, size=10)
        formula_font = Font(size=10, italic=True, color='0066CC')
        advantage_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

        # 标题
        ws['A1'] = '无人车 vs 人驾轻卡：成本对比分析（动态模型）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:F1')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '💡 本Sheet所有对比自动计算，修改参数表即可更新'
        ws['A2'].font = Font(size=10, italic=True, color='008000')

        row = 4

        # 一、初始投资对比
        ws[f'A{row}'] = '一、初始投资对比'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        headers = ['项目', '无人车', '人驾轻卡', '绝对差异', '相对差异', '']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row, col, header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
        row += 1

        ws.cell(row, 1, '采购成本')
        ws.cell(row, 2, f"={self.get_param_cell('购车成本', '无人车')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, f"={self.get_param_cell('购车成本', '人驾')}")
        ws.cell(row, 3).font = formula_font
        ws.cell(row, 3).number_format = '#,##0'
        ws.cell(row, 4, f"=B{row}-C{row}")
        ws.cell(row, 4).font = formula_font
        ws.cell(row, 4).number_format = '#,##0'
        ws.cell(row, 4).fill = advantage_fill
        ws.cell(row, 5, f"=(B{row}-C{row})/C{row}")
        ws.cell(row, 5).font = formula_font
        ws.cell(row, 5).number_format = '0.0%'
        ws.cell(row, 5).fill = advantage_fill
        row += 2

        # 二、5年TCO对比
        ws[f'A{row}'] = '二、5年TCO对比（核心）'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row, col, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        ws.cell(row, 1, '5年TCO总成本').font = Font(bold=True, size=11)
        ws.cell(row, 2, f"='2-无人车TCO'!{self.rv_cells['tco_total']}")
        ws.cell(row, 2).font = Font(bold=True, size=11, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, f"='3-人驾轻卡TCO'!{self.tr_cells['tco_total']}")
        ws.cell(row, 3).font = Font(bold=True, size=11, italic=True, color='0066CC')
        ws.cell(row, 3).number_format = '#,##0'
        ws.cell(row, 4, f"=B{row}-C{row}")
        ws.cell(row, 4).font = Font(bold=True, size=11, italic=True, color='FF0000')
        ws.cell(row, 4).number_format = '#,##0'
        ws.cell(row, 4).fill = advantage_fill
        ws.cell(row, 5, f"=(B{row}-C{row})/C{row}")
        ws.cell(row, 5).font = Font(bold=True, size=11, italic=True, color='FF0000')
        ws.cell(row, 5).number_format = '0.0%'
        ws.cell(row, 5).fill = advantage_fill
        ws.cell(row, 6, '⭐⭐⭐').font = Font(bold=True, color='FF0000')
        row += 2

        # 三、单位成本对比
        ws[f'A{row}'] = '三、单位成本对比'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row, col, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        # 每公里成本
        ws.cell(row, 1, '每公里成本')
        ws.cell(row, 2, f"='2-无人车TCO'!{self.rv_cells['tco_total']}/'2-无人车TCO'!{self.rv_cells['lifecycle_km']}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '0.00'
        ws.cell(row, 3, f"='3-人驾轻卡TCO'!{self.tr_cells['tco_total']}/'3-人驾轻卡TCO'!{self.tr_cells['lifecycle_km']}")
        ws.cell(row, 3).font = formula_font
        ws.cell(row, 3).number_format = '0.00'
        ws.cell(row, 4, f"=B{row}-C{row}")
        ws.cell(row, 4).font = formula_font
        ws.cell(row, 4).number_format = '0.00'
        ws.cell(row, 4).fill = advantage_fill
        ws.cell(row, 5, f"=(B{row}-C{row})/C{row}")
        ws.cell(row, 5).font = formula_font
        ws.cell(row, 5).number_format = '0.0%'
        ws.cell(row, 5).fill = advantage_fill
        ws.cell(row, 6, '⭐').font = Font(bold=True, color='FF0000')
        row += 1

        # 每件成本
        ws.cell(row, 1, '每件成本')
        ws.cell(row, 2, f"='2-无人车TCO'!{self.rv_cells['tco_total']}/'2-无人车TCO'!{self.rv_cells['lifecycle_items']}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '0.000'
        ws.cell(row, 3, f"='3-人驾轻卡TCO'!{self.tr_cells['tco_total']}/'3-人驾轻卡TCO'!{self.tr_cells['lifecycle_items']}")
        ws.cell(row, 3).font = formula_font
        ws.cell(row, 3).number_format = '0.000'
        ws.cell(row, 4, f"=B{row}-C{row}")
        ws.cell(row, 4).font = formula_font
        ws.cell(row, 4).number_format = '0.000'
        ws.cell(row, 4).fill = advantage_fill
        ws.cell(row, 5, f"=(B{row}-C{row})/C{row}")
        ws.cell(row, 5).font = formula_font
        ws.cell(row, 5).number_format = '0.0%'
        ws.cell(row, 5).fill = advantage_fill
        ws.cell(row, 6, '⭐').font = Font(bold=True, color='FF0000')
        row += 2

        # 四、运营效率对比
        ws[f'A{row}'] = '四、运营效率对比'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row, col, header)
            cell.font = header_font
            cell.fill = header_fill
        row += 1

        efficiency_items = [
            ('年运营天数', '年运营天数', '天'),
            ('日运营时长', '日运营时长', '小时'),
            ('日均里程', 'daily_km', 'km'),
            ('5年总里程', 'lifecycle_km', 'km'),
        ]

        for label, param_or_cell, unit in efficiency_items:
            ws.cell(row, 1, label)
            if param_or_cell in ['年运营天数', '日运营时长']:
                ws.cell(row, 2, f"={self.get_param_cell(param_or_cell, '无人车')}")
                ws.cell(row, 3, f"={self.get_param_cell(param_or_cell, '人驾')}")
            else:
                ws.cell(row, 2, f"='2-无人车TCO'!{self.rv_cells[param_or_cell]}")
                ws.cell(row, 3, f"='3-人驾轻卡TCO'!{self.tr_cells[param_or_cell]}")

            ws.cell(row, 2).font = formula_font
            ws.cell(row, 2).number_format = '#,##0'
            ws.cell(row, 3).font = formula_font
            ws.cell(row, 3).number_format = '#,##0'
            ws.cell(row, 4, f"=B{row}-C{row}")
            ws.cell(row, 4).font = formula_font
            ws.cell(row, 4).number_format = '+#,##0;-#,##0;0'
            ws.cell(row, 4).fill = advantage_fill
            ws.cell(row, 5, f"=(B{row}-C{row})/C{row}")
            ws.cell(row, 5).font = formula_font
            ws.cell(row, 5).number_format = '+0.0%;-0.0%;0%'
            ws.cell(row, 5).fill = advantage_fill
            row += 1

        # 设置列宽
        ws.column_dimensions['A'].width = 18
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 8

        return ws

    def generate(self, output_path='output/robovan_analysis.xlsx'):
        """生成完整动态Excel模型"""
        print("🚀 开始生成Robovan商业模式分析Excel（动态模型）...")

        print("📝 创建Sheet 1: 参数配置（输入区）...")
        self.create_sheet_parameters()

        print("📝 创建Sheet 2: 无人车TCO（公式驱动）...")
        self.create_sheet_robovan_dynamic()

        print("📝 创建Sheet 3: 人驾轻卡TCO（公式驱动）...")
        self.create_sheet_traditional_dynamic()

        print("📝 创建Sheet 4: 成本对比（公式驱动）...")
        self.create_sheet_comparison_dynamic()

        # 保存文件
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        self.wb.save(output_path)

        print(f"\n✅ 动态Excel模型生成成功！")
        print(f"📁 文件路径: {output_path}")
        print(f"\n⭐ 核心特性：")
        print(f"  • 所有黄色单元格可直接修改")
        print(f"  • 修改参数后，其他Sheet自动更新")
        print(f"  • 无需重新运行Python程序")
        print(f"  • 真正的动态模型！")


def main():
    """主函数"""
    analyzer = RobovanAnalyzerDynamic('assumptions.json')
    analyzer.generate('output/robovan_analysis.xlsx')


if __name__ == '__main__':
    main()
