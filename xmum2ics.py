import re
import argparse
import os
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from icalendar import Calendar, Event, vRecur
import pytz

SEMESTER_START = datetime(2026, 4, 6) 
TIMEZONE = pytz.timezone('Asia/Kuala_Lumpur')
START_HOUR = 8 

def parse_schedule(html):
    soup = BeautifulSoup(html, 'html.parser')
    rows = soup.find('tbody').find_all('tr')
    
    skip_cells = {i: 0 for i in range(7)}
    courses = []

    for row_idx, row in enumerate(rows):
        cells = row.find_all('td')
        col_idx = 0  
        td_idx = 1   

        while col_idx < 7:
            if skip_cells[col_idx] > 0:
                skip_cells[col_idx] -= 1
                col_idx += 1
                continue

            if td_idx < len(cells):
                cell = cells[td_idx]
                td_idx += 1
                rowspan = int(cell.get('rowspan', 1))
                
                if 'row_kb' in cell.get('class', []):
                    info = [text.strip() for text in cell.stripped_strings]
                    week_str = info[-1]
                    match = re.search(r'Week\s*(\d+)-(\d+)', week_str, re.IGNORECASE)
                    
                    if match:
                        start_week = int(match.group(1))
                        end_week = int(match.group(2))
                        total_weeks = end_week - start_week + 1
                        start_offset_weeks = start_week - 1
                    else:
                        total_weeks = 14
                        start_offset_weeks = 0
                        
                    courses.append({
                        'day': col_idx,               
                        'start_offset': row_idx,      
                        'duration': rowspan,          
                        'code': info[0] if len(info) > 0 else '',
                        'name': info[1] if len(info) > 1 else '',
                        'lecturer': info[2] if len(info) > 2 else '',
                        'location': info[3] if len(info) > 3 else '',
                        'total_weeks': total_weeks,
                        'start_offset_weeks': start_offset_weeks
                    })
                    
                skip_cells[col_idx] += rowspan - 1
            col_idx += 1
            
    return courses

def generate_ics(courses, filename):
    cal = Calendar()
    cal.add('prodid', '-//Dynamic Course Schedule//')
    cal.add('version', '2.0')

    for course in courses:
        event = Event()
        course_date = SEMESTER_START + timedelta(days=course['day']) + timedelta(weeks=course['start_offset_weeks'])
        start_time = course_date.replace(hour=START_HOUR + course['start_offset'], minute=0, second=0)
        end_time = start_time + timedelta(hours=course['duration'])
        
        start_time = TIMEZONE.localize(start_time)
        end_time = TIMEZONE.localize(end_time)

        event.add('summary', f"[{course['code']}] {course['name']}")
        event.add('location', course['location'])
        event.add('description', f"Lecturer: {course['lecturer']}")
        event.add('dtstart', start_time)
        event.add('dtend', end_time)
        event.add('rrule', vRecur({'freq': 'weekly', 'count': course['total_weeks']}))
        
        cal.add_component(event)

    with open(filename, 'wb') as f:
        f.write(cal.to_ical())
    print(f"✅ 成功导出 {len(courses)} 门课程的排课信息至 {filename}")

def main():
    parser = argparse.ArgumentParser(description="将教务系统 HTML 课表转换为 ICS 日历文件。")
    
    parser.add_argument("html_path", help="本地 HTML 课表文件的路径")
    
    parser.add_argument("-o", "--output", default="schedule.ics", help="输出的 ICS 文件名 (默认: schedule.ics)")

    args = parser.parse_args()

    if not os.path.exists(args.html_path):
        print(f"❌ 错误: 找不到文件 '{args.html_path}'")
        return

    try:
        with open(args.html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
    except Exception as e:
        print(f"❌ 读取文件时发生错误: {e}")
        return

    print(f"正在解析: {args.html_path} ...")
    parsed_data = parse_schedule(html_content)
    
    if not parsed_data:
        print("⚠️ 未在 HTML 中找到任何课程信息，请检查文件内容或解析逻辑。")
        return
        
    generate_ics(parsed_data, args.output)

if __name__ == "__main__":
    main()