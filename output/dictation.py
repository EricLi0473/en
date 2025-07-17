#!/usr/bin/env python3
"""
Dictation PDF Generator with Wrongbook Support
Author: ChatGPT
Requirements:
    pip install pandas reportlab

Usage examples:
----------------
✅ 常用命令用例全集（新版 dictation.py 支持错题本）
📘 1. 全默写指定 List（例：List 1, 5, 7）
python dictation.py --excel 英文单词.xlsx generate --mode full --lists 1,5,7

🎲 2. 抽查模式（例：从 List 2, 3 随机抽取 30 个）
python dictation.py --excel 英文单词.xlsx generate --mode sample --lists 2,3 --count 30

📘 3. 加入错题本内容一起输出（例：List 1 + 错题本）
python dictation.py --excel 英文单词.xlsx generate --mode full --lists 1 --include-wb

✍️ 4. 添加错题本条目（交互式输入 10-1、2-5 等）
python dictation.py --excel 英文单词.xlsx wb add

❌ 5. 从错题本中删除条目
python dictation.py --excel 英文单词.xlsx wb remove

🧾 6. 单独输出错题本为 PDF（题目 + 答案）
python dictation.py --excel 英文单词.xlsx wb output

🌱 7. 设置随机种子（可复现的抽样）
python dictation.py --excel 英文单词.xlsx generate --mode sample --lists 1,2,3 --count 20 --seed 42
"""

import argparse
import random
import re
import os
from typing import List, Set
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Frame, PageTemplate
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN_LEFT = 12   # ~0.17 inch
MARGIN_RIGHT = 12
MARGIN_TOP = 36
MARGIN_BOTTOM = 36
COLUMN_GAP = 18
FONT_SIZE = 15

pdfmetrics.registerFont(TTFont('SimSun', 'simsun.ttc'))  # Windows SimSun, fallback to NotoSansCJK if needed

styles = getSampleStyleSheet()
para_style = styles['Normal']
para_style.fontName = 'SimSun'
para_style.fontSize = FONT_SIZE
para_style.leading = FONT_SIZE + 2

# --------------------------- 数据加载与过滤 --------------------------- #

def load_words(excel_path: str) -> pd.DataFrame:
    """读取 Excel，返回 DataFrame(list_no, index, chinese, pos, english)"""
    df = pd.read_excel(excel_path, header=None, names=['index_raw', 'chinese', 'pos', 'english'])
    df['list_no'] = df['index_raw'].str.extract(r'List(\d+)', expand=False).astype(int)
    df['index'] = df['index_raw'].str.extract(r'-(\d+)', expand=False).astype(int)
    return df[['list_no', 'index', 'chinese', 'pos', 'english']]

def filter_by_lists(df: pd.DataFrame, lists: List[int]) -> pd.DataFrame:
    return df[df['list_no'].isin(lists)].sort_values(['list_no', 'index']).reset_index(drop=True)

def random_sample(df: pd.DataFrame, count: int) -> pd.DataFrame:
    if count > len(df):
        raise ValueError(f"Requested {count} words but only {len(df)} available from selected lists.")
    return df.sample(n=count).sort_values(['list_no', 'index']).reset_index(drop=True)

# --------------------------- 错题本工具 --------------------------- #

WRONG_PATTERN = re.compile(r'^(\d+)[-–](\d+)$')  # 支持 10-1 或 10–1

def parse_wrong_ref(ref: str):
    m = WRONG_PATTERN.match(ref.strip())
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))

def read_wrongbook(path: str) -> Set[str]:
    if not os.path.exists(path):
        return set()
    with open(path, 'r', encoding='utf-8') as f:
        return set(line.strip() for line in f if line.strip())

def write_wrongbook(path: str, refs: Set[str]):
    with open(path, 'w', encoding='utf-8') as f:
        for r in sorted(refs, key=lambda x: (int(x.split('-')[0]), int(x.split('-')[1]))):
            f.write(r + '\n')

# --------------------------- PDF 生成 --------------------------- #

def build_two_column_story(rows):
    story = []
    for text in rows:
        story.append(Paragraph(text, para_style))
        story.append(Spacer(1, 8))
    return story

def export_pdf(rows, output_path: str):
    column_width = (PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT - COLUMN_GAP) / 2

    frame1 = Frame(MARGIN_LEFT, MARGIN_BOTTOM, column_width, PAGE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM,
                leftPadding=0, rightPadding=6, topPadding=0, bottomPadding=0, id='col1')
    frame2 = Frame(MARGIN_LEFT + column_width + COLUMN_GAP, MARGIN_BOTTOM, column_width, PAGE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM,
                leftPadding=6, rightPadding=0, topPadding=0, bottomPadding=0, id='col2')

    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            leftMargin=MARGIN_LEFT, rightMargin=MARGIN_RIGHT,
                            topMargin=MARGIN_TOP, bottomMargin=MARGIN_BOTTOM)

    template = PageTemplate(id='TwoCol', frames=[frame1, frame2])
    doc.addPageTemplates([template])
    doc.build(build_two_column_story(rows))

# --------------------------- 文本格式化 --------------------------- #

def format_answer_rows(df: pd.DataFrame):
    return [f"{r.list_no}-{r.index}. {r.chinese} ({r.pos}) — {r.english}" for r in df.itertuples()]

def format_dictation_rows(df: pd.DataFrame):
    return [f"{r.chinese} ({r.pos})" for r in df.itertuples()]

# --------------------------- 错题本交互 --------------------------- #

def wrongbook_interactive(action: str, df: pd.DataFrame, wb_path: str):
    refs = read_wrongbook(wb_path)
    print(f"当前错题本已有 {len(refs)} 条。输入条目如 10-1，空行或 q 退出。")
    while True:
        user_in = input('> ').strip()
        if user_in == '' or user_in.lower() in {'q', 'quit', 'exit'}:
            break
        parsed = parse_wrong_ref(user_in)
        if not parsed:
            print('格式错误，应为 ListIndex-WordIndex (如 10-1)')
            continue
        ref_str = f"{parsed[0]}-{parsed[1]}"
        # 验证是否存在
        exists = not df[(df['list_no'] == parsed[0]) & (df['index'] == parsed[1])].empty
        if not exists:
            print('该编号在词库中不存在！')
            continue
        if action == 'add':
            refs.add(ref_str)
            print(f'已添加 {ref_str}')
        elif action == 'remove':
            if ref_str in refs:
                refs.remove(ref_str)
                print(f'已删除 {ref_str}')
            else:
                print('错题本中无此条')
    write_wrongbook(wb_path, refs)
    print(f"已写入 {wb_path}，共 {len(refs)} 条错题。")

# --------------------------- 主程序 --------------------------- #

def main():
    parser = argparse.ArgumentParser(description="Generate English dictation PDFs with wrongbook support")
    # 通用
    parser.add_argument('--excel', type=str, required=True, help='Path to Excel file')
    parser.add_argument('--wb-file', type=str, default='wrongbook.txt', help='Wrongbook storage file')

    subparsers = parser.add_subparsers(dest='command', required=True)

    # full / sample 生成模式
    gen_parser = subparsers.add_parser('generate', help='Generate PDFs (full or sample)')
    gen_parser.add_argument('--mode', choices=['full', 'sample'], required=True)
    gen_parser.add_argument('--lists', type=str, required=True, help='Comma-separated list numbers, e.g., 1,3,5')
    gen_parser.add_argument('--count', type=int, default=0, help='Number of words to sample (sample mode)')
    gen_parser.add_argument('--seed', type=int, default=None, help='Random seed')
    gen_parser.add_argument('--output', type=str, default='output', help='Output directory')
    gen_parser.add_argument('--include-wb', action='store_true', help='Include wrongbook entries in result')

    # 错题本 add/remove/output
    wb_parser = subparsers.add_parser('wb', help='Manage or output wrongbook')
    wb_parser.add_argument('action', choices=['add', 'remove', 'output'], help='add/remove/output')
    wb_parser.add_argument('--output', type=str, default='output', help='Output directory when action=output')

    args = parser.parse_args()

    df_all = load_words(args.excel)

    # -------------------- 错题本子命令 -------------------- #
    if args.command == 'wb':
        if args.action in {'add', 'remove'}:
            wrongbook_interactive(args.action, df_all, args.wb_file)
        elif args.action == 'output':
            refs = read_wrongbook(args.wb_file)
            if not refs:
                print('错题本为空')
                return
            df_refs = []
            for ref in refs:
                l_no, idx = map(int, ref.split('-'))
                row = df_all[(df_all['list_no'] == l_no) & (df_all['index'] == idx)]
                if not row.empty:
                    df_refs.append(row.iloc[0])
            if not df_refs:
                print('错题本中的条目在词库中未找到')
                return
            df_refs = pd.DataFrame(df_refs)
            os.makedirs(args.output, exist_ok=True)
            ans_pdf = os.path.join(args.output, 'wrongbook_answer.pdf')
            ques_pdf = os.path.join(args.output, 'wrongbook_dictation.pdf')
            export_pdf(format_answer_rows(df_refs), ans_pdf)
            export_pdf(format_dictation_rows(df_refs), ques_pdf)
            print(f'已输出:\n  {ans_pdf}\n  {ques_pdf}')
        return

    # -------------------- 生成 full / sample -------------------- #
    lists = [int(x) for x in re.split(r'[，,]', args.lists) if x.strip().isdigit()]
    if not lists:
        raise ValueError('No valid list numbers provided.')

    random.seed(args.seed)
    df_filtered = filter_by_lists(df_all, lists)

    if args.mode == 'sample':
        if args.count <= 0:
            raise ValueError('--count must be positive when mode=sample')
        df_final = random_sample(df_filtered, args.count)
    else:
        df_final = df_filtered.sample(frac=1).reset_index(drop=True)  # 打乱顺序

    # 加入错题本
    if args.include_wb:
        refs = read_wrongbook(args.wb_file)
        ref_rows = []
        for ref in refs:
            l_no, idx = map(int, ref.split('-'))
            row = df_all[(df_all['list_no'] == l_no) & (df_all['index'] == idx)]
            if not row.empty:
                ref_rows.append(row.iloc[0])
        if ref_rows:
            df_wb = pd.DataFrame(ref_rows)
            df_final = pd.concat([df_final, df_wb]).drop_duplicates(['list_no', 'index'])
            df_final = df_final.sort_values(['list_no', 'index']).reset_index(drop=True)

    os.makedirs(args.output, exist_ok=True)
    base_name = f"{args.mode}_{'-'.join(map(str, lists))}"

    answer_pdf_path = os.path.join(args.output, base_name + '_answer.pdf')
    dictation_pdf_path = os.path.join(args.output, base_name + '_dictation.pdf')

    export_pdf(format_answer_rows(df_final), answer_pdf_path)
    export_pdf(format_dictation_rows(df_final), dictation_pdf_path)

    print(f'Generated:\n  {answer_pdf_path}\n  {dictation_pdf_path}')


if __name__ == '__main__':
    main()
