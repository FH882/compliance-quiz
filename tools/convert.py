#!/usr/bin/env python3
"""
Word(.docx)の問題・解答ファイルをJSON形式に変換するスクリプト。

使い方:
  py convert.py

新しい問題を追加するには:
  SOURCE_DIRに問題ファイルと解答ファイルを配置して再実行してください。
"""
import zipfile
import xml.etree.ElementTree as ET
import json
import re
import os
import sys

NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

SOURCE_DIR = r'C:\Users\haya1\OneDrive\ドキュメント\仕事関係\1　法令等遵守責任者\AI\問題、解答・解説'
OUTPUT_FILE = r'C:\Users\haya1\brain\20_Projects\compliance-quiz\data\questions.json'

QUESTION_RE = re.compile(r'^第(\d+)問(?:【(.+?)】)?\s+(.+)', re.DOTALL)
ANSWER_RE   = re.compile(r'^第(\d+)問[\s　]+正解[：:]\s*(\d)')


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
                'num': q_num,
                'type': q_type,
                'question': q_text,
                'choices': choices,
                'answer': None,
                'explanation': ''
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
            explanation_parts = []
            j = i + 1
            while j < len(paras):
                line = paras[j].strip()
                if ANSWER_RE.match(line):
                    break
                if line and line != '【解説】':
                    explanation_parts.append(line)
                j += 1
            answers[q_num] = {
                'answer': ans,
                'explanation': '\n'.join(explanation_parts)
            }
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


def main():
    files = sorted(os.listdir(SOURCE_DIR))
    # 奇数番号ファイルが問題、偶数番号ファイルが解答（001→問題、002→解答...）
    q_files = [f for f in files if f.endswith('.docx')
               and int(f[:3]) % 2 == 1]

    subjects = []

    for q_file in q_files:
        q_num_int = int(q_file[:3])
        a_num_str = str(q_num_int + 1).zfill(3)
        a_file = next((f for f in files if f.startswith(a_num_str) and f.endswith('.docx')), None)

        if not a_file:
            print(f'警告: 解答ファイルが見つかりません → {q_file}', file=sys.stderr)
            continue

        q_path = os.path.join(SOURCE_DIR, q_file)
        a_path = os.path.join(SOURCE_DIR, a_file)

        print(f'処理中: {q_file}')

        q_paras = extract_paragraphs(q_path)
        a_paras = extract_paragraphs(a_path)

        questions = parse_questions(q_paras)
        answers   = parse_answers(a_paras)

        for q in questions:
            if q['num'] in answers:
                q['answer']      = answers[q['num']]['answer']
                q['explanation'] = answers[q['num']]['explanation']
            else:
                print(f'  警告: 第{q["num"]}問の解答がありません', file=sys.stderr)

        subject_name = get_subject_name(q_file)
        subjects.append({
            'id':        q_file[:3],
            'name':      subject_name,
            'questions': questions
        })

    output = {'subjects': subjects}
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    total_q = sum(len(s['questions']) for s in subjects)
    print(f'\n完了！ {len(subjects)}テキスト / {total_q}問')
    print(f'出力先: {OUTPUT_FILE}')


if __name__ == '__main__':
    main()
