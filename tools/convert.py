#!/usr/bin/env python3
"""
Word(.docx)の問題・解答ファイルをJSON形式に変換するスクリプト。

使い方:
  py convert.py

新しい問題を追加するには:
  対象フォルダにファイルを追加して再実行してください。
"""
import zipfile
import xml.etree.ElementTree as ET
import json
import re
import os
import sys

NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

SOURCE_DIR = r'C:\Users\haya1\OneDrive\ドキュメント\仕事関係\1　法令等遵守責任者\AI\問題、解答・解説(テーマ別)'
DIFF_DIR   = r'C:\Users\haya1\OneDrive\ドキュメント\仕事関係\1　法令等遵守責任者\AI\問題、解答・解説(難易度別テスト)'
OUTPUT_FILE = r'C:\Users\haya1\brain\20_Projects\compliance-quiz\data\questions.json'

QUESTION_RE  = re.compile(r'^第(\d+)問(?:【(.+?)】)?\s+(.+)', re.DOTALL)
ANSWER_RE    = re.compile(r'^第(\d+)問[\s　]+正解[：:]\s*(\d)')
DIFF_Q_RE    = re.compile(r'^第(\d+)問[\s　]+(.+)')
CHOICE_RE    = re.compile(r'^[アイウエ][．.]\s*(.+)')
KANA_TO_NUM  = {'ア': 1, 'イ': 2, 'ウ': 3, 'エ': 4}


# ── テキスト抽出 ──────────────────────────────────────────────
def extract_paragraphs(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        with z.open('word/document.xml') as f:
            content = f.read().decode('utf-8')
    tree = ET.fromstring(content)
    paragraphs = []
    for p in tree.iter(NS + 'p'):
        texts = [t.text or '' for t in p.iter(NS + 't')]
        line = ''.join(texts).strip()
        paragraphs.append(line)
    return paragraphs


def extract_table_rows(docx_path):
    """Word表を [[cell, cell, ...], ...] のリストとして返す"""
    with zipfile.ZipFile(docx_path) as z:
        with z.open('word/document.xml') as f:
            content = f.read().decode('utf-8')
    tree = ET.fromstring(content)
    rows = []
    for tbl in tree.iter(NS + 'tbl'):
        for tr in tbl.iter(NS + 'tr'):
            cells = []
            for tc in tr.iter(NS + 'tc'):
                texts = [t.text or '' for t in tc.iter(NS + 't')]
                cells.append(''.join(texts).strip())
            if cells:
                rows.append(cells)
    return rows


# ── 通常形式（001〜048）───────────────────────────────────────
def parse_questions(paras):
    questions = []
    i = 0
    while i < len(paras):
        m = QUESTION_RE.match(paras[i])
        if m:
            q_num  = int(m.group(1))
            q_type = m.group(2) or '一般問題'
            q_text = m.group(3).strip()
            choices = []
            j = i + 1
            while j < len(paras) and len(choices) < 4:
                line = paras[j].strip()
                if QUESTION_RE.match(line) or line.startswith('■'):
                    break
                if line:
                    choices.append(line)
                j += 1
            questions.append({
                'num': q_num, 'type': q_type,
                'question': q_text, 'choices': choices,
                'answer': None, 'explanation': ''
            })
            i = j
        else:
            i += 1
    return questions


def parse_answers(paras):
    answers = {}
    i = 0
    while i < len(paras):
        m = ANSWER_RE.match(paras[i])
        if m:
            q_num = int(m.group(1))
            ans   = int(m.group(2))
            parts = []
            j = i + 1
            while j < len(paras):
                line = paras[j].strip()
                if ANSWER_RE.match(line):
                    break
                if line and line != '【解説】':
                    parts.append(line)
                j += 1
            answers[q_num] = {'answer': ans, 'explanation': '\n'.join(parts)}
            i = j
        else:
            i += 1
    return answers


def get_subject_name(filename):
    name = re.sub(r'^\d+[\s　]*', '', filename)
    name = re.sub(r'（第\d+問[^）]*）\.docx$', '', name)
    name = re.sub(r'[\s　]+第\d+問.*$', '', name)
    name = name.replace('.docx', '').strip()
    return name


# ── 難易度別形式（ア/イ/ウ/エ + 表形式解答）──────────────────
def parse_diff_questions(paras):
    questions = []
    i = 0
    while i < len(paras):
        m = DIFF_Q_RE.match(paras[i])
        if m:
            q_num  = int(m.group(1))
            q_text = m.group(2).strip()
            choices = []
            j = i + 1
            while j < len(paras) and len(choices) < 4:
                line = paras[j].strip()
                if DIFF_Q_RE.match(line):
                    break
                cm = CHOICE_RE.match(line)
                if cm:
                    choices.append(cm.group(1).strip())
                j += 1
            questions.append({
                'num': q_num, 'type': '一般問題',
                'question': q_text, 'choices': choices,
                'answer': None, 'explanation': ''
            })
            i = j
        else:
            i += 1
    return questions


def parse_diff_answers(docx_path):
    """Word表から解答・解説を抽出する（列: 問/解答/解説/参照）"""
    rows = extract_table_rows(docx_path)
    answers = {}
    for row in rows:
        if len(row) < 3:
            continue
        try:
            q_num = int(row[0])
        except ValueError:
            continue  # ヘッダー行はスキップ
        letter      = row[1].strip()
        explanation = row[2].strip()
        reference   = row[3].strip() if len(row) > 3 else ''
        answer_num  = KANA_TO_NUM.get(letter)
        if answer_num:
            exp_text = explanation
            if reference:
                exp_text += f'\n【参照】{reference}'
            answers[q_num] = {'answer': answer_num, 'explanation': exp_text}
    return answers


# ── メイン ────────────────────────────────────────────────────
def main():
    subjects = []

    # 1. 通常問題（001〜048）
    files = sorted(os.listdir(SOURCE_DIR))
    q_files = [f for f in files if f.endswith('.docx') and int(f[:3]) % 2 == 1]

    for q_file in q_files:
        q_num_int = int(q_file[:3])
        a_num_str = str(q_num_int + 1).zfill(3)
        a_file = next((f for f in files if f.startswith(a_num_str) and f.endswith('.docx')), None)
        if not a_file:
            print(f'警告: 解答ファイルなし → {q_file}', file=sys.stderr)
            continue

        print(f'処理中: {q_file}')
        q_paras = extract_paragraphs(os.path.join(SOURCE_DIR, q_file))
        a_paras = extract_paragraphs(os.path.join(SOURCE_DIR, a_file))
        questions = parse_questions(q_paras)
        answers   = parse_answers(a_paras)

        for q in questions:
            if q['num'] in answers:
                q['answer']      = answers[q['num']]['answer']
                q['explanation'] = answers[q['num']]['explanation']

        subjects.append({
            'id': q_file[:3],
            'name': get_subject_name(q_file),
            'questions': questions
        })

    # 2. 難易度別問題
    DIFF_FILES = [
        ('低難易度_問題.docx', '低難易度_解答解説.docx', '【難易度別】低難易度テスト', 'D01'),
        ('中難易度_問題.docx', '中難易度_解答解説.docx', '【難易度別】中難易度テスト', 'D02'),
        ('高難易度_問題.docx', '高難易度_解答解説.docx', '【難易度別】高難易度テスト', 'D03'),
    ]

    for q_fname, a_fname, name, sid in DIFF_FILES:
        q_path = os.path.join(DIFF_DIR, q_fname)
        a_path = os.path.join(DIFF_DIR, a_fname)
        if not os.path.exists(q_path) or not os.path.exists(a_path):
            print(f'警告: ファイルなし → {q_fname}', file=sys.stderr)
            continue

        print(f'処理中: {q_fname}')
        q_paras   = extract_paragraphs(q_path)
        questions = parse_diff_questions(q_paras)
        answers   = parse_diff_answers(a_path)

        for q in questions:
            if q['num'] in answers:
                q['answer']      = answers[q['num']]['answer']
                q['explanation'] = answers[q['num']]['explanation']

        subjects.append({'id': sid, 'name': name, 'questions': questions})

    # 3. 出力
    output = {'subjects': subjects}
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    total_q = sum(len(s['questions']) for s in subjects)
    print(f'\n完了！ {len(subjects)}テキスト / {total_q}問')
    print(f'出力先: {OUTPUT_FILE}')


if __name__ == '__main__':
    main()
