#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
法律人AI实操夜校 - 学员名单抓取脚本
从 https://www.legalagi.cn 抓取并整理学员名单
"""

import requests
from bs4 import BeautifulSoup
import json
from datetime import datetime


def scrape_students():
    """抓取学员名单"""
    url = "https://www.legalagi.cn/students/00_往期学员名录.html"

    print(f"正在访问: {url}")

    try:
        response = requests.get(url, timeout=30)
        response.encoding = 'utf-8'
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"请求失败: {e}")
        return None

    soup = BeautifulSoup(response.text, 'html.parser')

    students = []

    # 查找学员区块
    alumni_section = soup.find('section', class_='alumni-directory__section')

    if alumni_section:
        # 获取期数名称
        header = alumni_section.find('div', class_='alumni-directory__header')
        if header:
            season = header.find('p', class_='alumni-directory__season')
            cohort = header.find('h2')
            current_season = season.text.strip() if season else ""
            current_cohort = cohort.text.strip() if cohort else ""
        else:
            current_season = ""
            current_cohort = ""

        # 查找所有学员卡片
        avatars = alumni_section.find_all('article', class_='alumni-avatar')

        for avatar in avatars:
            name_elem = avatar.find('strong', class_='student-avatar__name')
            note_elem = avatar.find('span', class_='student-avatar__note')

            name = name_elem.text.strip() if name_elem else ""
            note = note_elem.text.strip() if note_elem else ""

            students.append({
                "name": name,
                "note": note
            })

        return {
            "season": current_season,
            "cohort": current_cohort,
            "count": len(students),
            "students": students
        }

    return None


def export_to_markdown(data, filename="学员名单.md"):
    """导出为 Markdown 格式"""
    if not data:
        print("没有数据可导出")
        return

    content = f"""# {data['season']} - {data['cohort']}

> 抓取时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
> 总人数: {data['count']}人

## 学员名单

| 序号 | 姓名 | 备注 |
|------|------|------|
"""

    for i, student in enumerate(data['students'], 1):
        note = student['note'] if student['note'] else "-"
        content += f"| {i} | {student['name']} | {note} |\n"

    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)

    print(f"✅ 已导出到: {filename}")


def export_to_json(data, filename="学员名单.json"):
    """导出为 JSON 格式"""
    if not data:
        print("没有数据可导出")
        return

    data['scraped_at'] = datetime.now().isoformat()

    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"✅ 已导出到: {filename}")


def export_to_csv(data, filename="学员名单.csv"):
    """导出为 CSV 格式"""
    if not data:
        print("没有数据可导出")
        return

    with open(filename, 'w', encoding='utf-8-sig') as f:
        f.write("序号,姓名,备注\n")
        for i, student in enumerate(data['students'], 1):
            note = student['note'] if student['note'] else ""
            f.write(f"{i},{student['name']},{note}\n")

    print(f"✅ 已导出到: {filename}")


def print_summary(data):
    """打印摘要"""
    if not data:
        return

    print("\n" + "="*50)
    print(f"📚 {data['season']} - {data['cohort']}")
    print(f"👥 总人数: {data['count']} 人")
    print("="*50 + "\n")

    for i, student in enumerate(data['students'], 1):
        note = f" ({student['note']})" if student['note'] else ""
        print(f"{i:2d}. {student['name']}{note}")


if __name__ == "__main__":
    # 抓取数据
    data = scrape_students()

    if data:
        # 打印摘要
        print_summary(data)

        # 导出多种格式
        export_to_markdown(data)
        export_to_json(data)
        export_to_csv(data)

        print("\n✨ 完成!")
    else:
        print("❌ 未能获取数据")