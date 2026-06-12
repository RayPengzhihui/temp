#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生产成本结转凭证生成脚本
输出格式：会计软件导入模板（30列标准格式）

生成规则（参照参考脚本）：
  - 筛选：FR列="结转" 且 AV列有数字
  - 借方：科目=FT列，摘要=FS列，金额=AV列
  - 贷方：AW-BP列中，第5行有5001开头科目 且 当前行有金额 的列
  - 会计年/会计期间/制单日期/凭证号/制单人 → 留空
  - 币种名称 → 人民币（借贷行均填）
  - 凭证ID → 借方和贷方行都填同一编号
  - 借方科目以66开头或=6901 → 不做项目核算（项目大类/项目编码留空），部门编码填来源值
  - 贷方行 → 项目大类编码和项目编码始终填写
"""

import sys
import openpyxl
import xlwt
from openpyxl.utils import column_index_from_string

# ── 可配置字段 ─────────────────────────────────────
VOUCHER_TYPE     = '转'   # 凭证类别
PROJECT_CATEGORY = '97'   # 项目大类编码

# ── 源列索引（0-based）────────────────────────────
COL_PROJECT = 3                                       # D   项目编号
# 结转类型
COL_AV = column_index_from_string('AV') - 1          # 47  借方金额
COL_AW = column_index_from_string('AW') - 1          # 48  贷方明细起始
COL_BP = column_index_from_string('BP') - 1          # 67  贷方明细结束
# 售后运维费用类型
COL_DG = column_index_from_string('DG') - 1          # 110 借方金额
COL_DH = column_index_from_string('DH') - 1          # 111 贷方明细起始
COL_EA = column_index_from_string('EA') - 1          # 130 贷方明细结束

COL_FR = column_index_from_string('FR') - 1          # 173 是否结转
COL_FS = column_index_from_string('FS') - 1          # 174 摘要
COL_FT = column_index_from_string('FT') - 1          # 175 借方科目
COL_FU = column_index_from_string('FU') - 1          # 176 部门编码（售后运维费用借方使用）

if len(sys.argv) > 1:
    INPUT_FILE = sys.argv[1]
else:
    print('============================================')
    print('  生产成本结转凭证生成工具')
    print('============================================')
    INPUT_FILE = input('请输入文件名（如 生产成本0713.xlsx）：').strip()
OUTPUT_FILE = '凭证.xls'

# 输出文件列头（30列标准格式）
HEADERS = [
    '凭证ID', '会计年', '会计期间', '制单日期', '凭证类别', '凭证号', '制单人',
    '所附单据数', '备注1', '备注2',
    '科目编码', '摘要', '结算方式编码', '票据号', '票据日期', '币种名称', '汇率',
    '单价', '借方数量', '贷方数量', '原币借方', '原币贷方',
    '借方金额', '贷方金额',
    '部门编码', '职员编码', '客户编码', '供应商编码',
    '项目大类编码', '项目编码',
]

# 输出列索引
H = {h: i for i, h in enumerate(HEADERS)}


def acct_str(val):
    """科目编码转字符串（去掉浮点小数点）"""
    if isinstance(val, float):
        return str(int(val))
    return str(val).strip() if val else ''


def load_data():
    wb = openpyxl.load_workbook(INPUT_FILE, read_only=True, data_only=True)
    ws = wb['生产成本']
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows


def build_credit_cols(row5, start, end):
    """从第5行提取指定列范围内有科目编码的列 -> {列索引: 科目字符串}"""
    cols = {}
    for i in range(start, end + 1):
        acct = row5[i] if i < len(row5) else None
        if acct and str(acct).strip():
            cols[i] = acct_str(acct)
    return cols


def build_voucher_groups(rows):
    row5 = rows[4]  # 第5行：贷方科目行

    # 按类型分别构建贷方科目列映射
    credit_cols_map = {
        '结转':       build_credit_cols(row5, COL_AW, COL_BP),  # AW-BP
        '售后运维费用': build_credit_cols(row5, COL_DH, COL_EA),  # DH-EA
    }

    # FR 值 -> 凭证ID 映射
    VOUCHER_ID_MAP = {
        '结转':       1,
        '售后运维费用': 2,
    }

    groups = []

    for row in rows[6:]:  # 第7行起是数据
        fr_val  = row[COL_FR] if COL_FR < len(row) else None
        fs_val  = row[COL_FS] if COL_FS < len(row) else None
        ft_val  = row[COL_FT] if COL_FT < len(row) else None
        fu_val  = row[COL_FU] if COL_FU < len(row) else None
        proj_no = row[COL_PROJECT] if COL_PROJECT < len(row) else None

        fr_str = str(fr_val).strip()
        if fr_str not in VOUCHER_ID_MAP:
            continue

        voucher_id = VOUCHER_ID_MAP[fr_str]

        # 按类型选取借方金额列
        if fr_str == '售后运维费用':
            av_val = row[COL_DG] if COL_DG < len(row) else None
        else:
            av_val = row[COL_AV] if COL_AV < len(row) else None
        if not isinstance(av_val, (int, float)) or av_val == 0:
            continue

        summary    = str(fs_val).strip() if fs_val else ''
        debit_acct = acct_str(ft_val)
        dept_code  = str(fu_val).strip() if fu_val else ''
        proj_str   = str(proj_no).strip() if proj_no else ''

        # 是否需要项目核算（借方科目以66开头或等于6901则不需要）
        need_proj = not (debit_acct.startswith('66') or debit_acct == '6901')

        def make_entry(is_debit, account, amount):
            r = [''] * len(HEADERS)
            r[H['凭证ID']]   = voucher_id
            r[H['凭证类别']] = VOUCHER_TYPE
            r[H['科目编码']] = account
            r[H['摘要']]     = summary
            r[H['币种名称']] = '人民币'
            if is_debit:
                r[H['借方金额']]    = round(float(amount), 2)
                # 售后运维费用类型借方填部门编码（FU列）
                r[H['部门编码']]    = dept_code if fr_str == '售后运维费用' else ''
                r[H['项目大类编码']] = PROJECT_CATEGORY if need_proj else ''
                r[H['项目编码']]    = proj_str if need_proj else ''
            else:
                r[H['贷方金额']]    = round(float(amount), 2)
                r[H['项目大类编码']] = PROJECT_CATEGORY
                r[H['项目编码']]    = proj_str
            return r

        entries = []
        # 借方行
        entries.append(make_entry(True, debit_acct, av_val))

        # 贷方行（按类型选对应贷方列范围）
        credit_total = 0
        for col_idx, acct in credit_cols_map[fr_str].items():
            val = row[col_idx] if col_idx < len(row) else None
            if isinstance(val, (int, float)) and val != 0:
                entries.append(make_entry(False, acct, val))
                credit_total += val

        # 借贷平衡检查
        if abs(av_val - credit_total) > 0.01:
            print(f'  警告: {proj_str} 借贷不平衡: 借方={av_val}, 贷方合计={credit_total:.2f}')

        if len(entries) > 1:
            groups.append(entries)

    return groups


def write_xls(groups):
    wb_out = xlwt.Workbook(encoding='utf-8')
    ws = wb_out.add_sheet('Sheet1')

    for col, h in enumerate(HEADERS):
        ws.write(0, col, h)

    row_idx = 1
    for group in groups:
        for entry in group:
            for col, val in enumerate(entry):
                ws.write(row_idx, col, val)
            row_idx += 1

    wb_out.save(OUTPUT_FILE)
    return row_idx - 1


def main():
    print(f'读取 {INPUT_FILE} ...')
    rows = load_data()

    groups = build_voucher_groups(rows)
    total = sum(len(g) for g in groups)
    print(f'找到 {len(groups)} 条结转凭证，共 {total} 行分录')

    write_xls(groups)
    print(f'已生成 {OUTPUT_FILE}')

    print('\n前3条凭证预览：')
    for g in groups[:3]:
        print('─' * 72)
        for e in g:
            d = f"{e[H['借方金额']]:>12.2f}" if e[H['借方金额']] != '' else ' ' * 12
            c = f"{e[H['贷方金额']]:>12.2f}" if e[H['贷方金额']] != '' else ' ' * 12
            print(f"  ID:{e[H['凭证ID']]:>4}  科目:{e[H['科目编码']]:<14} 借:{d}  贷:{c}  "
                  f"项目:{e[H['项目编码']]}  币种:{e[H['币种名称']]}")


if __name__ == '__main__':
    main()
    input('\n按回车键退出...')
