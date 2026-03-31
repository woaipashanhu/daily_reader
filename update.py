#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日阅读更新脚本
用法: python3 update.py <word文件路径> [标题]
"""

import os
import sys
import json
from datetime import datetime
from docx import Document

def update_article(word_path, title=None):
    """更新文章到网站"""
    
    # 读取Word文件
    doc = Document(word_path)
    
    # 提取所有段落文本
    paragraphs = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            paragraphs.append(text)
    
    content = '\n'.join(paragraphs)
    
    # 生成日期和文件名
    today = datetime.now().strftime('%Y-%m-%d')
    json_filename = f'{today}.json'
    
    # 如果没有提供标题，使用第一个段落或日期
    if not title:
        title = paragraphs[0] if paragraphs else today
    
    # 文章JSON内容
    article_data = {
        'date': today,
        'title': title,
        'content': content
    }
    
    # 保存文章JSON文件
    articles_dir = os.path.join(os.path.dirname(__file__), 'articles')
    json_path = os.path.join(articles_dir, json_filename)
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(article_data, f, ensure_ascii=False, indent=2)
    
    print(f'已保存文章: {json_path}')
    
    # 更新文章列表
    list_path = os.path.join(articles_dir, 'list.json')
    
    if os.path.exists(list_path):
        with open(list_path, 'r', encoding='utf-8') as f:
            articles_list = json.load(f)
    else:
        articles_list = []
    
    # 检查今天是否已有文章，有则更新，无则添加
    existing_index = None
    for i, article in enumerate(articles_list):
        if article.get('date') == today:
            existing_index = i
            break
    
    article_info = {
        'date': today,
        'title': title,
        'file': json_filename
    }
    
    if existing_index is not None:
        articles_list[existing_index] = article_info
        print(f'已更新今天的文章')
    else:
        articles_list.append(article_info)
        print(f'已添加新文章')
    
    # 按日期排序（最新在前）
    articles_list.sort(key=lambda x: x.get('date', ''), reverse=True)
    
    with open(list_path, 'w', encoding='utf-8') as f:
        json.dump(articles_list, f, ensure_ascii=False, indent=2)
    
    print(f'已更新文章列表，共 {len(articles_list)} 篇文章')
    print(f'\n请运行以下命令部署到GitHub:')
    print(f'  cd {os.path.dirname(__file__)}')
    print(f'  git add .')
    print(f'  git commit -m "更新文章 {today}"')
    print(f'  git push')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python3 update.py <word文件路径> [标题]')
        sys.exit(1)
    
    word_path = sys.argv[1]
    title = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(word_path):
        print(f'错误: 文件不存在 - {word_path}')
        sys.exit(1)
    
    update_article(word_path, title)
