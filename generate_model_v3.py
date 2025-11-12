#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Robovan Business Model Analyzer v3.0
生成动态Excel模型，包含双视角分析（运营商+厂商）
"""

import json
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


class RobovanAnalyzerV3:
    """Robovan商业模式分析器 v3.0 - 双视角动态模型"""

    def __init__(self, config_path='assumptions.json'):
        self.config_path = config_path
        self.config = self.load_config()
        self.wb = Workbook()
        # 删除默认sheet
        if 'Sheet' in self.wb.sheetnames:
            del self.wb['Sheet']

        # 初始化单元格映射
        self.init_cell_mapping()

    def load_config(self):
        """加载配置文件"""
        with open(self.config_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def init_cell_mapping(self):
        """初始化参数单元格位置映射（扩展版-包含厂商参数）"""
        # Sheet1的参数位置（行号）
        self.param_rows = {
            # === A区：车辆基本参数 (7-18行) ===
            '车型': 7,
            '运力': 8,
            '购车成本': 9,
            '电池容量': 10,
            '续航里程': 11,
            '生命周期': 12,
            '残值': 13,

            # === B区：运营参数 (22-39行) ===
            '年运营天数': 22,
            '日运营时长': 23,
            '平均时速': 24,
            '百公里电耗': 28,
            '电价': 29,
            '单次往返里程': 33,
            '单次满载件数': 34,
            '市场定价_公里': 38,
            '市场定价_件': 39,

            # === C区：运营商年度成本 (43-48行) ===
            '人力成本': 43,
            '软件订阅': 44,
            '保险费用': 45,
            '维保成本': 46,
            '其他成本': 47,

            # === D区：厂商成本参数 (52-65行) ===
            '销售规模': 52,
            '固定研发': 53,
            '单车边际研发': 54,
            '总研发投入': 55,  # 计算字段
            '单车研发分摊': 56,  # 计算字段
            '硬件BOM': 58,
            '生产成本': 59,
            '销售费用率': 60,
            '软件毛利率': 61,

            # === E区：监控成本参数 (65-71行) ===
            '监控固定成本': 65,
            '安全员年薪': 66,
            '安全员配比': 67,
            '需要安全员': 68,  # 计算字段
            '监控总成本': 69,  # 计算字段
            '单车监控成本': 70,  # 计算字段

            # === F区：一次性买断参数 (75-82行) ===
            '买断总价': 75,
            '买断生命周期': 76,
            '买断残值': 77,
            '买断技术支持费': 78,
            '买断人力成本': 79,
            '买断保险': 80,
            '买断维保': 81,
            '买断其他': 82,
        }

        # 列映射：B列=无人车/订阅，C列=人驾/买断
        self.col_operator_sub = 'B'    # 运营商-订阅模式
        self.col_operator_trad = 'C'   # 运营商-人驾模式
        self.col_manufacturer = 'B'     # 厂商参数列
        self.col_buyout = 'C'          # 买断模式列

    def get_param_cell(self, param_name, column='B'):
        """获取参数单元格地址

        Args:
            param_name: 参数名称
            column: 列名（B/C等）
        """
        row = self.param_rows.get(param_name)
        if row is None:
            raise ValueError(f"未找到参数: {param_name}")
        return f"'1-参数配置'!{column}{row}"

    def create_sheet_parameters(self):
        """创建Sheet 1: 参数配置总表（5个区块）"""
        ws = self.wb.create_sheet("1-参数配置")
        self.sheet1_name = "1-参数配置"

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

        # === 标题 ===
        ws['A1'] = 'Robovan商业模式分析 v3.0 - 参数配置（双视角动态模型）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:F1')
        ws.row_dimensions[1].height = 30

        ws['A2'] = f"生成日期：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws['A2'].font = Font(size=10, italic=True)
        ws['A3'] = "⭐ 黄色单元格可修改，绿色单元格自动计算，其他Sheet会实时更新！"
        ws['A3'].font = Font(size=11, color='FF0000', bold=True)
        ws['A4'] = "📊 v3.0新增：厂商视角分析、一次性买断模式、规模化成本模型"
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
        basic_data = [
            ('车型', self.config['无人车']['基本信息']['车型'],
             self.config['人驾电动轻卡']['基本信息']['车型'], '-', '市场主流车型', False),
            ('运力', self.config['无人车']['基本信息']['运力_立方米'],
             self.config['人驾电动轻卡']['基本信息']['运力_立方米'], 'm³', '官方参数', True),
            ('车辆采购成本', self.config['无人车']['投资成本']['车辆采购成本_元'],
             self.config['人驾电动轻卡']['投资成本']['车辆采购成本_元'], '元', '官网价格', True),
            ('电池容量', self.config['无人车']['投资成本']['电池容量_kWh'],
             self.config['人驾电动轻卡']['投资成本']['电池容量_kWh'], 'kWh', '官方参数', True),
            ('续航里程', self.config['无人车']['投资成本']['续航里程_公里'],
             self.config['人驾电动轻卡']['投资成本']['续航里程_公里'], 'km', 'CLTC工况', True),
            ('生命周期', self.config['无人车']['投资成本']['生命周期_年'],
             self.config['人驾电动轻卡']['投资成本']['生命周期_年'], '年', '无人车6年/人驾7年', True),
            ('残值', self.config['无人车']['投资成本']['残值_元'],
             self.config['人驾电动轻卡']['投资成本']['残值_元'], '元', '约购车成本38%/15%', True),
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
                if isinstance(rv_val, int) and rv_val > 100:
                    cell_rv.number_format = '#,##0'
                    cell_tr.number_format = '#,##0'

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

        # B区数据 - 运营时间
        ws.cell(row, 1, '年运营天数')
        ws.cell(row, 2, self.config['无人车']['运营时间参数']['年运营天数_天']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, self.config['人驾电动轻卡']['运营时间参数']['年运营天数_天']).fill = input_fill
        ws.cell(row, 3).font = input_font
        ws.cell(row, 4, '天')
        ws.cell(row, 5, '无人车350天/人驾330天')
        row += 1

        ws.cell(row, 1, '日有效运营时长')
        ws.cell(row, 2, self.config['无人车']['运营时间参数']['日有效运营时长_小时']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, self.config['人驾电动轻卡']['运营时间参数']['日有效运营时长_小时']).fill = input_fill
        ws.cell(row, 3).font = input_font
        ws.cell(row, 4, '小时')
        ws.cell(row, 5, '无人车10h/人驾8h')
        row += 1

        ws.cell(row, 1, '平均运营时速')
        ws.cell(row, 2, self.config['无人车']['运营时间参数']['平均运营时速_公里每小时']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, self.config['人驾电动轻卡']['运营时间参数']['平均运营时速_公里每小时']).fill = input_fill
        ws.cell(row, 3).font = input_font
        ws.cell(row, 4, 'km/h')
        ws.cell(row, 5, '无人车25/人驾28')
        row += 1

        row += 1  # 空行

        # B区数据 - 能源参数
        ws.cell(row, 1, '百公里电耗')
        ws.cell(row, 2, self.config['无人车']['能源参数']['百公里电耗_kWh']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, self.config['人驾电动轻卡']['能源参数']['百公里电耗_kWh']).fill = input_fill
        ws.cell(row, 3).font = input_font
        ws.cell(row, 4, 'kWh')
        ws.cell(row, 5, '无人车含传感器耗电')
        row += 1

        ws.cell(row, 1, '电价')
        ws.cell(row, 2, self.config['无人车']['能源参数']['电价_元每kWh']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, self.config['人驾电动轻卡']['能源参数']['电价_元每kWh']).fill = input_fill
        ws.cell(row, 3).font = input_font
        ws.cell(row, 4, '元/kWh')
        ws.cell(row, 5, '工业用电峰谷混合价')
        row += 1

        row += 1  # 空行

        # B区数据 - 业务场景
        ws.cell(row, 1, '单次往返里程')
        ws.cell(row, 2, self.config['业务场景参数']['单次往返里程_公里']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, self.config['业务场景参数']['单次往返里程_公里']).fill = input_fill
        ws.cell(row, 3).font = input_font
        ws.cell(row, 4, 'km')
        ws.cell(row, 5, '仓库到配送中心')
        row += 1

        ws.cell(row, 1, '单次满载件数')
        ws.cell(row, 2, self.config['业务场景参数']['单次满载件数_件']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, self.config['业务场景参数']['单次满载件数_件']).fill = input_fill
        ws.cell(row, 3).font = input_font
        ws.cell(row, 4, '件')
        ws.cell(row, 5, '标准包裹数量')
        row += 1

        row += 1  # 空行

        # B区数据 - 市场定价
        ws.cell(row, 1, '市场定价(公里)')
        ws.cell(row, 2, self.config['市场定价参数']['按公里收费_元每公里']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, self.config['市场定价参数']['按公里收费_元每公里']).fill = input_fill
        ws.cell(row, 3).font = input_font
        ws.cell(row, 4, '元/km')
        ws.cell(row, 5, '外部市场价格')
        row += 1

        ws.cell(row, 1, '市场定价(件)')
        ws.cell(row, 2, self.config['市场定价参数']['按件收费_元每件']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, self.config['市场定价参数']['按件收费_元每件']).fill = input_fill
        ws.cell(row, 3).font = input_font
        ws.cell(row, 4, '元/件')
        ws.cell(row, 5, '外部市场价格')
        row += 1

        row += 1  # 空行

        # ========================================
        # C区：运营商年度成本
        # ========================================
        ws[f'A{row}'] = '【C区】运营商年度成本'
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
        cost_data = [
            ('人力成本', self.config['无人车']['年度成本_元']['人力成本'],
             self.config['人驾电动轻卡']['年度成本_元']['人力成本'], '元', '远程监控/司机工资'),
            ('软件订阅费', self.config['无人车']['年度成本_元']['软件订阅费'],
             self.config['人驾电动轻卡']['年度成本_元']['软件订阅费'], '元', '自动驾驶系统订阅'),
            ('保险费用', self.config['无人车']['年度成本_元']['保险费用'],
             self.config['人驾电动轻卡']['年度成本_元']['保险费用'], '元', '交强险+商业险'),
            ('维保成本', self.config['无人车']['年度成本_元']['维保成本'],
             self.config['人驾电动轻卡']['年度成本_元']['维保成本'], '元', '维修保养费用'),
            ('其他成本', self.config['无人车']['年度成本_元']['其他成本'],
             self.config['人驾电动轻卡']['年度成本_元']['其他成本'], '元', '停车/杂费'),
        ]

        for param_name, rv_val, tr_val, unit, note in cost_data:
            ws.cell(row, 1, param_name)
            cell_rv = ws.cell(row, 2, rv_val)
            cell_tr = ws.cell(row, 3, tr_val)
            cell_rv.font = input_font
            cell_rv.fill = input_fill
            cell_rv.number_format = '#,##0'
            cell_tr.font = input_font
            cell_tr.fill = input_fill
            cell_tr.number_format = '#,##0'
            ws.cell(row, 4, unit)
            ws.cell(row, 5, note)
            row += 1

        row += 1  # 空行

        # ========================================
        # D区：厂商成本参数
        # ========================================
        ws[f'A{row}'] = '【D区】厂商成本参数（规模驱动）'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        ws.cell(row, 1, '参数项').font = header_font
        ws.cell(row, 1).fill = header_fill
        ws.cell(row, 2, '数值').font = header_font
        ws.cell(row, 2).fill = header_fill
        ws.cell(row, 3, '单位').font = header_font
        ws.cell(row, 3).fill = header_fill
        ws.cell(row, 4, '说明').font = header_font
        ws.cell(row, 4).fill = header_fill
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # D区数据 - 规模
        ws.cell(row, 1, '★销售规模')
        ws.cell(row, 2, self.config['厂商成本参数']['销售规模']['车队规模_台']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '台')
        ws.cell(row, 4, '⭐关键参数：影响研发分摊和监控成本')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # D区数据 - 研发成本
        ws.cell(row, 1, '固定研发投入')
        ws.cell(row, 2, self.config['厂商成本参数']['研发成本']['固定研发投入_元']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '平台、算法等基础研发')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '单车边际研发')
        ws.cell(row, 2, self.config['厂商成本参数']['研发成本']['单车边际研发_元']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '场景适配成本')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # 计算字段：总研发投入
        ws.cell(row, 1, '总研发投入')
        total_rd_formula = f"=B53+B54*B52"
        ws.cell(row, 2, total_rd_formula).fill = calc_fill
        ws.cell(row, 2).font = calc_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '=固定研发+单车边际×规模')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # 计算字段：单车研发分摊
        ws.cell(row, 1, '单车研发分摊')
        unit_rd_formula = f"=B55/B52"
        ws.cell(row, 2, unit_rd_formula).fill = calc_fill
        ws.cell(row, 2).font = calc_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '=总研发÷销售规模（规模效应关键）')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        row += 1  # 空行

        # D区数据 - 单车制造成本
        ws.cell(row, 1, '硬件BOM成本')
        ws.cell(row, 2, self.config['厂商成本参数']['单车制造成本']['硬件BOM成本_元']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '激光雷达+传感器+车体')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '生产制造成本')
        ws.cell(row, 2, self.config['厂商成本参数']['单车制造成本']['生产制造成本_元']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '组装+测试+质检')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '销售费用率')
        ws.cell(row, 2, self.config['厂商成本参数']['销售与服务']['销售费用率']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '0%'
        ws.cell(row, 3, '%')
        ws.cell(row, 4, '占销售收入比例')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '软件订阅毛利率')
        ws.cell(row, 2, self.config['厂商成本参数']['销售与服务']['软件订阅毛利率']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '0%'
        ws.cell(row, 3, '%')
        ws.cell(row, 4, '软件毛利（成本=订阅费×40%）')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        row += 1  # 空行

        # ========================================
        # E区：远程监控参数
        # ========================================
        ws[f'A{row}'] = '【E区】远程监控成本（规模效应）'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        ws.cell(row, 1, '参数项').font = header_font
        ws.cell(row, 1).fill = header_fill
        ws.cell(row, 2, '数值').font = header_font
        ws.cell(row, 2).fill = header_fill
        ws.cell(row, 3, '单位').font = header_font
        ws.cell(row, 3).fill = header_fill
        ws.cell(row, 4, '说明').font = header_font
        ws.cell(row, 4).fill = header_fill
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # E区数据
        ws.cell(row, 1, '监控中心固定成本')
        ws.cell(row, 2, self.config['远程监控参数']['固定成本']['监控中心年固定成本_元']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '场地+系统+管理人员')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '安全员年薪')
        ws.cell(row, 2, self.config['远程监控参数']['变动成本']['安全员年薪_元']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '监控员人力成本')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '★安全员配比')
        ws.cell(row, 2, self.config['远程监控参数']['变动成本']['安全员配比_车每人']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, '车/人')
        ws.cell(row, 4, '⭐关键参数：1人监控N台车')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # 计算字段：需要安全员数量
        ws.cell(row, 1, '需要安全员数量')
        safety_count_formula = f"=B52/B67"
        ws.cell(row, 2, safety_count_formula).fill = calc_fill
        ws.cell(row, 2).font = calc_font
        ws.cell(row, 2).number_format = '0'
        ws.cell(row, 3, '人')
        ws.cell(row, 4, '=销售规模÷配比')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # 计算字段：监控总成本
        ws.cell(row, 1, '监控年总成本')
        monitor_total_formula = f"=B65+B68*B66"
        ws.cell(row, 2, monitor_total_formula).fill = calc_fill
        ws.cell(row, 2).font = calc_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '=固定成本+人数×年薪')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # 计算字段：单车监控成本
        ws.cell(row, 1, '单车年监控成本')
        monitor_unit_formula = f"=B69/B52"
        ws.cell(row, 2, monitor_unit_formula).fill = calc_fill
        ws.cell(row, 2).font = calc_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '=监控总成本÷销售规模（规模效应）')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        row += 1  # 空行

        # ========================================
        # F区：一次性买断参数
        # ========================================
        ws[f'A{row}'] = '【F区】一次性买断模式参数'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

        ws.cell(row, 1, '参数项').font = header_font
        ws.cell(row, 1).fill = header_fill
        ws.cell(row, 2, '数值').font = header_font
        ws.cell(row, 2).fill = header_fill
        ws.cell(row, 3, '单位').font = header_font
        ws.cell(row, 3).fill = header_fill
        ws.cell(row, 4, '说明').font = header_font
        ws.cell(row, 4).fill = header_fill
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # F区数据
        ws.cell(row, 1, '车辆+软件总价')
        ws.cell(row, 2, self.config['一次性买断模式']['投资成本']['车辆加软件总价_元']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '一次性买断价格')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '生命周期')
        ws.cell(row, 2, self.config['一次性买断模式']['投资成本']['生命周期_年']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 3, '年')
        ws.cell(row, 4, '与订阅模式相同（6年）')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '残值')
        ws.cell(row, 2, self.config['一次性买断模式']['投资成本']['残值_元']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '+3000元系统价值')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '技术支持费')
        ws.cell(row, 2, self.config['一次性买断模式']['年度成本_元']['技术支持费']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '买断需单独支付')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '人力成本')
        ws.cell(row, 2, self.config['一次性买断模式']['年度成本_元']['人力成本']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '与订阅模式相同')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '保险费用')
        ws.cell(row, 2, self.config['一次性买断模式']['年度成本_元']['保险费用']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '与订阅模式相同')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '维保成本')
        ws.cell(row, 2, self.config['一次性买断模式']['年度成本_元']['维保成本']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '与订阅模式相同')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        ws.cell(row, 1, '其他成本')
        ws.cell(row, 2, self.config['一次性买断模式']['年度成本_元']['其他成本']).fill = input_fill
        ws.cell(row, 2).font = input_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '与订阅模式相同')
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # 设置列宽
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 15

        print("  ✓ Sheet 1: 参数配置（5个区块）已创建")

    def create_sheet_operator_subscription(self):
        """创建Sheet 2: 运营商-订阅模式TCO（6年）"""
        ws = self.wb.create_sheet("2-运营商-订阅模式")

        # 样式定义
        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        section_font = Font(bold=True, color='FFFFFF', size=11)
        formula_font = Font(size=10, italic=True, color='0066CC')
        highlight_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        note_font = Font(size=9, italic=True, color='666666')
        note_fill = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')

        # 标题
        ws['A1'] = '运营商视角 - 无人车订阅模式TCO分析（6年）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:E1')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '💡 本Sheet所有数据自动计算，修改参数表即可更新'
        ws['A2'].font = Font(size=10, italic=True, color='0066CC')

        row = 4

        # 一、年度成本
        ws[f'A{row}'] = '一、年度成本拆解'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        # 表头
        ws.cell(row, 1, '成本项').font = Font(bold=True)
        ws.cell(row, 2, '金额').font = Font(bold=True)
        ws.cell(row, 3, '单位').font = Font(bold=True)
        ws.cell(row, 4, '计算说明').font = Font(bold=True)
        row += 1

        # 成本数据
        costs = [
            ('人力成本', f"={self.get_param_cell('人力成本', 'B')}", '元/年', '远程监控分摊'),
            ('软件订阅费', f"={self.get_param_cell('软件订阅', 'B')}", '元/年', '自动驾驶系统年费'),
            ('保险费用', f"={self.get_param_cell('保险费用', 'B')}", '元/年', '交强+商业险'),
            ('维保成本', f"={self.get_param_cell('维保成本', 'B')}", '元/年', '维修保养'),
            ('其他成本', f"={self.get_param_cell('其他成本', 'B')}", '元/年', '停车/杂费'),
        ]

        cost_start_row = row
        for label, formula, unit, note in costs:
            ws.cell(row, 1, label)
            ws.cell(row, 2, formula)
            ws.cell(row, 2).font = formula_font
            ws.cell(row, 2).number_format = '#,##0'
            ws.cell(row, 3, unit)
            ws.cell(row, 4, note)
            ws.cell(row, 4).font = note_font
            ws.cell(row, 4).fill = note_fill
            row += 1

        # 年度总成本
        ws.cell(row, 1, '年度总成本').font = Font(bold=True, size=11)
        total_year_formula = f"=SUM(B{cost_start_row}:B{row-1})"
        ws.cell(row, 2, total_year_formula)
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '运营成本合计')
        ws.cell(row, 4).font = note_font
        total_year_cell = f"'2-运营商-订阅模式'!B{row}"
        row += 2

        # 二、6年TCO
        ws[f'A{row}'] = '二、6年总拥有成本(TCO)'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws.cell(row, 1, '车辆采购成本')
        ws.cell(row, 2, f"={self.get_param_cell('购车成本', 'B')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '一次性投入')
        ws.cell(row, 4).font = note_font
        row += 1

        ws.cell(row, 1, '6年运营成本')
        ws.cell(row, 2, f"={total_year_cell}*{self.get_param_cell('生命周期', 'B')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '=年成本×6年')
        ws.cell(row, 4).font = note_font
        row += 1

        ws.cell(row, 1, '减：残值回收')
        ws.cell(row, 2, f"=-{self.get_param_cell('残值', 'B')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '6年后转卖')
        ws.cell(row, 4).font = note_font
        row += 1

        ws.cell(row, 1, '6年TCO总成本').font = Font(bold=True, size=12, color='FF0000')
        tco_total_formula = f"=SUM(B{row-3}:B{row-1})"
        ws.cell(row, 2, tco_total_formula)
        ws.cell(row, 2).font = Font(bold=True, size=12, italic=True, color='FF0000')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '总拥有成本')
        ws.cell(row, 4).font = Font(bold=True, size=9, color='FF0000')
        tco_cell = f"'2-运营商-订阅模式'!B{row}"
        row += 2

        # 三、单位成本
        ws[f'A{row}'] = '三、单位成本'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws.cell(row, 1, '年均成本')
        ws.cell(row, 2, f"={tco_cell}/{self.get_param_cell('生命周期', 'B')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '=TCO÷6年')
        ws.cell(row, 4).font = note_font
        row += 1

        # 列宽
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 35
        ws.column_dimensions['E'].width = 15

        self.sheet2_tco_cell = tco_cell
        self.sheet2_year_cell = total_year_cell
        print("  ✓ Sheet 2: 运营商-订阅模式（6年TCO+计算说明）已创建")

    def create_sheet_operator_traditional(self):
        """创建Sheet 3: 运营商-人驾模式TCO（7年）"""
        ws = self.wb.create_sheet("3-运营商-人驾轻卡")

        # 样式（复用）
        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        section_font = Font(bold=True, color='FFFFFF', size=11)
        formula_font = Font(size=10, italic=True, color='0066CC')
        highlight_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        note_font = Font(size=9, italic=True, color='666666')
        note_fill = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')

        # 标题
        ws['A1'] = '运营商视角 - 人驾电动轻卡TCO分析（7年）'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:E1')
        ws.row_dimensions[1].height = 25

        ws['A2'] = '💡 本Sheet所有数据自动计算，修改参数表即可更新'
        ws['A2'].font = Font(size=10, italic=True, color='0066CC')

        row = 4

        # 一、年度成本
        ws[f'A{row}'] = '一、年度成本拆解'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws.cell(row, 1, '成本项').font = Font(bold=True)
        ws.cell(row, 2, '金额').font = Font(bold=True)
        ws.cell(row, 3, '单位').font = Font(bold=True)
        ws.cell(row, 4, '计算说明').font = Font(bold=True)
        row += 1

        costs = [
            ('司机工资', f"={self.get_param_cell('人力成本', 'C')}", '元/年', '含社保'),
            ('保险费用', f"={self.get_param_cell('保险费用', 'C')}", '元/年', '营运车辆险'),
            ('维保成本', f"={self.get_param_cell('维保成本', 'C')}", '元/年', '维修保养'),
            ('其他成本', f"={self.get_param_cell('其他成本', 'C')}", '元/年', '过路费等'),
        ]

        cost_start_row = row
        for label, formula, unit, note in costs:
            ws.cell(row, 1, label)
            ws.cell(row, 2, formula)
            ws.cell(row, 2).font = formula_font
            ws.cell(row, 2).number_format = '#,##0'
            ws.cell(row, 3, unit)
            ws.cell(row, 4, note)
            ws.cell(row, 4).font = note_font
            ws.cell(row, 4).fill = note_fill
            row += 1

        ws.cell(row, 1, '年度总成本').font = Font(bold=True, size=11)
        total_year_formula = f"=SUM(B{cost_start_row}:B{row-1})"
        ws.cell(row, 2, total_year_formula)
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '运营成本合计')
        ws.cell(row, 4).font = note_font
        total_year_cell = f"'3-运营商-人驾轻卡'!B{row}"
        row += 2

        # 二、7年TCO
        ws[f'A{row}'] = '二、7年总拥有成本(TCO)'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws.cell(row, 1, '车辆采购成本')
        ws.cell(row, 2, f"={self.get_param_cell('购车成本', 'C')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '一次性投入')
        ws.cell(row, 4).font = note_font
        row += 1

        ws.cell(row, 1, '7年运营成本')
        ws.cell(row, 2, f"={total_year_cell}*{self.get_param_cell('生命周期', 'C')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '=年成本×7年')
        ws.cell(row, 4).font = note_font
        row += 1

        ws.cell(row, 1, '减：残值回收')
        ws.cell(row, 2, f"=-{self.get_param_cell('残值', 'C')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '7年后转卖')
        ws.cell(row, 4).font = note_font
        row += 1

        ws.cell(row, 1, '7年TCO总成本').font = Font(bold=True, size=12, color='FF0000')
        tco_total_formula = f"=SUM(B{row-3}:B{row-1})"
        ws.cell(row, 2, tco_total_formula)
        ws.cell(row, 2).font = Font(bold=True, size=12, italic=True, color='FF0000')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        ws.cell(row, 3, '元')
        ws.cell(row, 4, '总拥有成本')
        ws.cell(row, 4).font = Font(bold=True, size=9, color='FF0000')
        tco_cell = f"'3-运营商-人驾轻卡'!B{row}"
        row += 2

        # 三、单位成本
        ws[f'A{row}'] = '三、单位成本'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:E{row}')
        row += 1

        ws.cell(row, 1, '年均成本')
        ws.cell(row, 2, f"={tco_cell}/{self.get_param_cell('生命周期', 'C')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '元/年')
        ws.cell(row, 4, '=TCO÷7年')
        ws.cell(row, 4).font = note_font

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 35

        self.sheet3_tco_cell = tco_cell
        self.sheet3_year_cell = total_year_cell
        print("  ✓ Sheet 3: 运营商-人驾轻卡（7年TCO+计算说明）已创建")

    def create_sheet_operator_comparison(self):
        """创建Sheet 4: 运营商对比分析"""
        ws = self.wb.create_sheet("4-运营商对比")

        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_font = Font(bold=True, color='FFFFFF', size=11)
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        formula_font = Font(size=10, italic=True, color='0066CC')
        green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

        ws['A1'] = '运营商视角 - 订阅模式 vs 人驾模式对比'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:D1')
        ws.row_dimensions[1].height = 25

        row = 3

        ws[f'A{row}'] = '核心指标对比'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 1

        ws.cell(row, 1, '对比项').font = Font(bold=True)
        ws.cell(row, 2, '订阅模式(6年)').font = Font(bold=True)
        ws.cell(row, 3, '人驾模式(7年)').font = Font(bold=True)
        ws.cell(row, 4, '优势方').font = Font(bold=True)
        row += 1

        ws.cell(row, 1, '生命周期TCO')
        ws.cell(row, 2, f"={self.sheet2_tco_cell}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, f"={self.sheet3_tco_cell}")
        ws.cell(row, 3).font = formula_font
        ws.cell(row, 3).number_format = '#,##0'
        ws.cell(row, 4, f"=IF(B{row}<C{row},\"订阅\",\"人驾\")")
        ws.cell(row, 4).font = Font(bold=True, color='006100')
        row += 1

        ws.cell(row, 1, '年均成本')
        ws.cell(row, 2, f"={self.sheet2_tco_cell}/6")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, f"={self.sheet3_tco_cell}/7")
        ws.cell(row, 3).font = formula_font
        ws.cell(row, 3).number_format = '#,##0'
        ws.cell(row, 4, f"=IF(B{row}<C{row},\"订阅\",\"人驾\")")
        ws.cell(row, 4).font = Font(bold=True, color='006100')
        row += 1

        ws.cell(row, 1, '初始投资')
        ws.cell(row, 2, f"={self.get_param_cell('购车成本', 'B')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, f"={self.get_param_cell('购车成本', 'C')}")
        ws.cell(row, 3).font = formula_font
        ws.cell(row, 3).number_format = '#,##0'
        ws.cell(row, 4, f"=IF(B{row}<C{row},\"订阅\",\"人驾\")")
        ws.cell(row, 4).font = Font(bold=True, color='006100')

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 15

        print("  ✓ Sheet 4: 运营商对比分析已创建")

    def create_sheet_manufacturer_analysis(self):
        """创建Sheet 5: 厂商盈利分析"""
        ws = self.wb.create_sheet("5-厂商盈利分析")

        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        section_font = Font(bold=True, color='FFFFFF', size=11)
        section_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        formula_font = Font(size=10, italic=True, color='0066CC')
        highlight_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        note_font = Font(size=9, italic=True, color='666666')

        ws['A1'] = '厂商视角 - 订阅模式 vs 一次性买断盈利分析'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:F1')
        ws.row_dimensions[1].height = 25

        ws['A2'] = f"当前场景：销售规模 {self.config['厂商成本参数']['销售规模']['车队规模_台']} 台 | " + \
                  f"安全员配比 1:{self.config['远程监控参数']['变动成本']['安全员配比_车每人']}"
        ws['A2'].font = Font(size=10, italic=True, color='0066CC')

        row = 4

        # 左侧：订阅模式
        ws[f'A{row}'] = '【订阅模式 - 厂商盈利分析】'
        ws[f'A{row}'].font = section_font
        ws[f'A{row}'].fill = section_fill
        ws.merge_cells(f'A{row}:C{row}')

        # 右侧：买断模式
        ws[f'D{row}'] = '【一次性买断 - 厂商盈利分析】'
        ws[f'D{row}'].font = section_font
        ws[f'D{row}'].fill = section_fill
        ws.merge_cells(f'D{row}:F{row}')
        row += 1

        # 订阅模式 - 单车成本
        ws.cell(row, 1, '单车初期成本').font = Font(bold=True)
        row += 1

        ws.cell(row, 1, '  硬件BOM')
        ws.cell(row, 2, f"={self.get_param_cell('硬件BOM')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '引用参数表')
        ws.cell(row, 3).font = note_font
        sub_bom_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '  研发分摊')
        ws.cell(row, 2, f"={self.get_param_cell('单车研发分摊')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '规模效应关键')
        ws.cell(row, 3).font = note_font
        sub_rd_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '  生产成本')
        ws.cell(row, 2, f"={self.get_param_cell('生产成本')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        sub_prod_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '  销售费用')
        sales_revenue = self.config['无人车']['投资成本']['车辆采购成本_元']
        ws.cell(row, 2, f"={sales_revenue}*{self.get_param_cell('销售费用率')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '收入×12%')
        ws.cell(row, 3).font = note_font
        sub_sales_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '初期成本合计').font = Font(bold=True)
        ws.cell(row, 2, f"=SUM({sub_bom_cell}:{sub_sales_cell})")
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        sub_init_cost_cell = f"B{row}"
        row += 2

        # 订阅模式 - 年度成本
        ws.cell(row, 1, '年度持续成本').font = Font(bold=True)
        row += 1

        ws.cell(row, 1, '  监控成本')
        ws.cell(row, 2, f"={self.get_param_cell('单车监控成本')}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '规模效应')
        ws.cell(row, 3).font = note_font
        sub_monitor_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '  软件运维')
        software_sub = self.config['无人车']['年度成本_元']['软件订阅费']
        ws.cell(row, 2, f"={software_sub}*(1-{self.get_param_cell('软件毛利率')})")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '订阅费×40%')
        ws.cell(row, 3).font = note_font
        sub_software_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '  技术支持')
        ws.cell(row, 2, 2000)
        ws.cell(row, 2).number_format = '#,##0'
        sub_support_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '年度成本小计').font = Font(bold=True)
        ws.cell(row, 2, f"=SUM({sub_monitor_cell}:{sub_support_cell})")
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        sub_year_cost_cell = f"B{row}"
        row += 2

        # 订阅模式 - 收入
        ws.cell(row, 1, '收入结构').font = Font(bold=True)
        row += 1

        ws.cell(row, 1, '  车辆销售收入')
        ws.cell(row, 2, sales_revenue)
        ws.cell(row, 2).number_format = '#,##0'
        sub_vehicle_rev_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '  6年软件订阅')
        ws.cell(row, 2, f"={software_sub}*6")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        sub_software_rev_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '6年总收入').font = Font(bold=True)
        ws.cell(row, 2, f"={sub_vehicle_rev_cell}+{sub_software_rev_cell}")
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='0066CC')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        sub_total_rev_cell = f"B{row}"
        row += 2

        # 订阅模式 - 利润
        ws.cell(row, 1, '6年利润分析').font = Font(bold=True, size=11, color='FF0000')
        row += 1

        ws.cell(row, 1, '  总收入')
        ws.cell(row, 2, f"={sub_total_rev_cell}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        row += 1

        ws.cell(row, 1, '  初期成本')
        ws.cell(row, 2, f"=-{sub_init_cost_cell}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        row += 1

        ws.cell(row, 1, '  6年持续成本')
        ws.cell(row, 2, f"=-{sub_year_cost_cell}*6")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        row += 1

        ws.cell(row, 1, '6年总利润').font = Font(bold=True, size=12, color='FF0000')
        ws.cell(row, 2, f"=SUM(B{row-3}:B{row-1})")
        ws.cell(row, 2).font = Font(bold=True, size=12, italic=True, color='FF0000')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 2).fill = highlight_fill
        sub_profit_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '单车净利率')
        ws.cell(row, 2, f"={sub_profit_cell}/{sub_total_rev_cell}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '0.0%'
        row += 2

        # 盈亏平衡分析
        ws.cell(row, 1, '盈亏平衡分析').font = Font(bold=True, size=11)
        ws.merge_cells(f'A{row}:C{row}')
        row += 1

        ws.cell(row, 1, '  单车利润(不含研发)')
        ws.cell(row, 2, f"={sub_profit_cell}+{sub_rd_cell}")
        ws.cell(row, 2).font = formula_font
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '回推单车边际利润')
        ws.cell(row, 3).font = note_font
        sub_margin_cell = f"B{row}"
        row += 1

        ws.cell(row, 1, '  盈亏平衡台数').font = Font(bold=True)
        ws.cell(row, 2, f"={self.get_param_cell('总研发投入')}/{sub_margin_cell}")
        ws.cell(row, 2).font = Font(bold=True, italic=True, color='FF0000')
        ws.cell(row, 2).number_format = '#,##0'
        ws.cell(row, 3, '需卖多少台盈利')
        ws.cell(row, 3).font = Font(bold=True, size=9, color='FF0000')

        # 列宽
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 22
        ws.column_dimensions['E'].width = 18
        ws.column_dimensions['F'].width = 20

        print("  ✓ Sheet 5: 厂商盈利分析（订阅模式+盈亏平衡）已创建")

    def create_sheet_comprehensive_comparison(self):
        """创建Sheet 6: 综合对比（简化版）"""
        ws = self.wb.create_sheet("6-综合对比")

        title_font = Font(size=14, bold=True, color='FFFFFF')
        title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')

        ws['A1'] = '双视角综合对比 - 运营商 & 厂商'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:D1')
        ws.row_dimensions[1].height = 25

        ws['A3'] = '运营商视角：订阅模式6年TCO更优'
        ws['A3'].font = Font(size=11, bold=True)

        ws['A5'] = '厂商视角：订阅模式更易盈亏平衡'
        ws['A5'].font = Font(size=11, bold=True)

        ws['A7'] = '💡 调整Sheet 1的"销售规模"和"安全员配比"参数可观察规模效应'
        ws['A7'].font = Font(size=10, italic=True, color='0066CC')

        ws.column_dimensions['A'].width = 60

        print("  ✓ Sheet 6: 综合对比已创建")

    def generate(self):
        """生成完整Excel模型"""
        print("\n🚀 开始生成Robovan商业模式分析 v3.0...")
        print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

        self.create_sheet_parameters()
        self.create_sheet_operator_subscription()
        self.create_sheet_operator_traditional()
        self.create_sheet_operator_comparison()
        self.create_sheet_manufacturer_analysis()
        self.create_sheet_comprehensive_comparison()

        print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

        # 保存
        output_dir = 'output'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        output_path = os.path.join(output_dir, 'robovan_analysis_v3.xlsx')
        self.wb.save(output_path)

        print(f"\n✅ 模型生成完成！")
        print(f"📁 文件路径: {output_path}")
        return output_path


def main():
    analyzer = RobovanAnalyzerV3()
    analyzer.generate()


if __name__ == '__main__':
    main()
