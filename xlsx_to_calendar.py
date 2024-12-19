import pandas as pd
from datetime import datetime, timedelta
import random
import os
import pytz
from icalendar import Calendar, Event
import warnings
warnings.filterwarnings('ignore', category=UserWarning)

def generate_random_color():
    """Generate a light pastel color"""
    hue = random.random()
    saturation = 0.3 + random.random() * 0.2
    value = 0.9 + random.random() * 0.1
    
    h = hue * 6
    c = value * saturation
    x = c * (1 - abs(h % 2 - 1))
    m = value - c
    
    if h < 1:
        r, g, b = c, x, 0
    elif h < 2:
        r, g, b = x, c, 0
    elif h < 3:
        r, g, b = 0, c, x
    elif h < 4:
        r, g, b = 0, x, c
    elif h < 5:
        r, g, b = x, 0, c
    else:
        r, g, b = c, 0, x
    
    r = int((r + m) * 255)
    g = int((g + m) * 255)
    b = int((b + m) * 255)
    
    return f"#{r:02x}{g:02x}{b:02x}"

def parse_meeting_pattern(pattern):
    """Parse the complete meeting pattern including dates, times, and location"""
    if not isinstance(pattern, str) or pd.isna(pattern):
        return []

    try:
        # Split into blocks and clean them
        blocks = [block.strip() for block in pattern.split('\n\n') if block.strip()]
        results = []
        
        for block in blocks:
            try:
                # Split the single line into parts using |
                parts = [p.strip() for p in block.split('|') if p.strip()]
                if len(parts) < 3:
                    continue

                # Parse date range
                date_range = parts[0]
                date_parts = date_range.split('-')
                if len(date_parts) != 2:
                    continue
                start_date = date_parts[0].strip()
                end_date = date_parts[1].strip()

                # Parse days and times
                day_time = parts[1]
                days = day_time.split()  # Split into individual days
                
                time_parts = parts[2].strip().split('-')
                if len(time_parts) != 2:
                    continue
                start_time = time_parts[0].strip()
                end_time = time_parts[1].strip()

                # Get location
                location = parts[3].strip() if len(parts) > 3 else ""

                # Add each day to results
                for day in days:
                    results.append({
                        'start_date': start_date,
                        'end_date': end_date,
                        'day': day,
                        'start_time': start_time,
                        'end_time': end_time,
                        'location': location
                    })
                    print(f"Added meeting: {day} {start_time}-{end_time} at {location}")

            except Exception as block_error:
                print(f"Error processing block: {block_error}")
                continue

        return results

    except Exception as e:
        print(f"Error in parse_meeting_pattern: {e}")
        return []

def get_term_dates(term):
    """Get start and end dates for the specified term"""
    if term.lower() == 'term1':
        return '2024/09/03', '2024/12/06'
    elif term.lower() == 'term2':
        return '2025/01/06', '2025/04/06'
    else:
        raise ValueError("Invalid term specified")

def is_in_term(start_date, term):
    """Check if a course is in the selected term"""
    try:
        course_date = datetime.strptime(start_date, '%Y/%m/%d')
        
        if term == 'term1':
            term_start = datetime(2024, 9, 3)
            term_end = datetime(2024, 12, 6)
        else:  # term2
            term_start = datetime(2025, 1, 6)
            term_end = datetime(2025, 4, 6)
            
        return term_start <= course_date <= term_end
    except Exception as e:
        print(f"Error checking term: {e}")
        return False

def generate_course_calendar(df, selected_term):
    """Generate HTML calendar from course data"""
    schedule = {
        'Monday': [],
        'Tuesday': [],
        'Wednesday': [],
        'Thursday': [],
        'Friday': []
    }
    
    course_styles = []
    course_colors = {}
    time_slots = {day: {} for day in schedule.keys()}
    
    # Track earliest and latest times
    earliest_time = 24 * 60  # Initialize to end of day
    latest_time = 0  # Initialize to start of day
    
    # Process courses and find time range
    for _, row in df.iterrows():
        try:
            course_name = str(row['Course Listing'])
            meeting_pattern = str(row['Meeting Patterns'])
            course_type = str(row.get('Instructional Format', 'Lecture'))
            
            if pd.isna(meeting_pattern):
                continue
            
            if course_name not in course_colors:
                course_colors[course_name] = generate_random_color()
                safe_name = course_name.replace(' ', '_').replace('.', '_')
                course_styles.append(f".course_{safe_name} {{ background-color: {course_colors[course_name]}; }}")
            
            meetings = parse_meeting_pattern(meeting_pattern)
            
            for meeting in meetings:
                if not is_in_term(meeting['start_date'], selected_term):
                    continue
                    
                start_parts = meeting['start_time'].split(':')
                end_parts = meeting['end_time'].split(':')
                start_mins = int(start_parts[0]) * 60 + int(start_parts[1])
                end_mins = int(end_parts[0]) * 60 + int(end_parts[1])
                
                # Update time range
                earliest_time = min(earliest_time, start_mins)
                latest_time = max(latest_time, end_mins)
                
                day_map = {
                    'Mon': 'Monday',
                    'Tue': 'Tuesday',
                    'Wed': 'Wednesday',
                    'Thu': 'Thursday',
                    'Fri': 'Friday'
                }
                day = day_map.get(meeting['day'])
                
                if day and day in schedule:
                    time_key = f"{meeting['start_time']}-{meeting['end_time']}"
                    
                    if time_key in time_slots[day] and time_slots[day][time_key] == course_name:
                        continue
                    
                    course_info = {
                        'name': course_name,
                        'start': meeting['start_time'],
                        'end': meeting['end_time'],
                        'location': meeting['location'],
                        'safe_name': course_name.replace(' ', '_').replace('.', '_'),
                        'type': course_type,
                        'start_mins': start_mins,
                        'end_mins': end_mins
                    }
                    schedule[day].append(course_info)
                    time_slots[day][time_key] = course_name
            
        except Exception as e:
            print(f"Error processing course {course_name}: {e}")
            continue

    # Round time range to nearest hour
    start_hour = (earliest_time // 60) - 1  # One hour before earliest class
    end_hour = (latest_time // 60) + 1  # One hour after latest class
    hour_height = 100  # 100px per hour

    # Generate HTML
    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Course Schedule</title>
        <style>
            body { 
                font-family: system-ui, -apple-system, sans-serif; 
                padding: 20px;
                margin: 0;
                background-color: #f5f5f5;
            }
            .calendar-header {
                display: grid;
                grid-template-columns: 120px repeat(5, 1fr);
                gap: 2px;
                margin-bottom: 2px;
            }
            .calendar-container {
                display: grid;
                grid-template-columns: 120px repeat(5, 1fr);
                gap: 2px;
                background: #fff;
                border: 1px solid #ddd;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }
            .header-cell {
                background: #f8f9fa;
                padding: 15px;
                text-align: center;
                font-weight: bold;
                font-size: 18px;
                color: #212529;
                border: 1px solid #dee2e6;
            }
            .time-column {
                background: #f8f9fa;
                padding: 0;
                border-right: 2px solid #dee2e6;
            }
            .time-slot {
                height: 100px;
                border-bottom: 1px solid #dee2e6;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 18px;
                color: #495057;
                font-weight: 500;
                position: relative;
            }
            .time-slot::after {
                content: '';
                position: absolute;
                right: -2px;
                width: 10px;
                height: 1px;
                background: #dee2e6;
            }
            .day-column {
                background: white;
                position: relative;
                border-right: 1px solid #dee2e6;
                min-height: 100px;
            }
            .course {
                position: absolute;
                left: 5px;
                right: 5px;
                padding: 8px;
                border-radius: 6px;
                font-size: 13px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                overflow: hidden;
                transition: transform 0.2s;
                cursor: default;
                background-color: white;
            }
            .course:hover {
                transform: scale(1.02);
                z-index: 100;
                box-shadow: 0 4px 8px rgba(0,0,0,0.15);
            }
            .course.laboratory {
                border-left: 4px solid #e74c3c;
            }
            .course-type {
                font-size: 11px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                margin-bottom: 3px;
                opacity: 0.8;
                font-weight: 600;
            }
            .time {
                font-weight: bold;
                margin-bottom: 3px;
                font-size: 13px;
            }
            .location {
                font-size: 11px;
                color: #495057;
                margin-top: 3px;
            }
    """
    
    # Add course styles
    for style in course_styles:
        html += style + "\n"
    
    html += """
        </style>
    </head>
    <body>
        <h2>Course Schedule</h2>
        <!-- Separate header row -->
        <div class="calendar-header">
            <div class="header-cell"></div>
            <div class="header-cell">Monday</div>
            <div class="header-cell">Tuesday</div>
            <div class="header-cell">Wednesday</div>
            <div class="header-cell">Thursday</div>
            <div class="header-cell">Friday</div>
        </div>
        <div class="calendar-container">
            <!-- Time column -->
            <div class="time-column">
    """
    
    # Add time slots
    for hour in range(start_hour, end_hour + 1):
        html += f'<div class="time-slot">{hour:02d}:00</div>'
    
    html += "</div>"
    
    # Add day columns
    for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']:
        html += f'<div class="day-column">'
        
        # Add courses for this day
        for course in sorted(schedule[day], key=lambda x: x['start_mins']):
            # Calculate position and height
            top = (course['start_mins'] - start_hour * 60) * (hour_height / 60)
            height = (course['end_mins'] - course['start_mins']) * (hour_height / 60)
            
            is_lab = course['type'].lower() == 'laboratory'
            lab_class = ' laboratory' if is_lab else ''
            
            html += f"""
                <div class="course course_{course['safe_name']}{lab_class}"
                     style="top: {top}px; height: {height}px;">
                    <div class="course-type">{course['type']}</div>
                    <div class="time">{course['start']} - {course['end']}</div>
                    <strong>{course['name']}</strong>
                    <div class="location">{course['location']}</div>
                </div>
            """
        
        html += "</div>"
    
    html += """
        </div>
    </body>
    </html>
    """
    
    return html

def generate_ics_calendar(df, term_start, term_end, selected_term):
    """Generate ICS calendar from course data"""
    cal = Calendar()
    cal.add('prodid', '-//Course Schedule Calendar//mxm.dk//')
    cal.add('version', '2.0')
    
    tz = pytz.timezone('America/Vancouver')
    
    # Track added events to avoid duplicates
    added_events = set()
    
    for _, row in df.iterrows():
        try:
            course_name = str(row['Course Listing'])
            meeting_pattern = str(row['Meeting Patterns'])
            
            if pd.isna(meeting_pattern):
                continue
            
            # Get additional information
            section = str(row.get('Section', ''))
            course_type = str(row.get('Instructional Format', ''))
            instructor = str(row.get('Instructor', ''))
            
            # Parse meeting patterns
            meetings = parse_meeting_pattern(meeting_pattern)
            
            for meeting in meetings:
                try:
                    # Check if the meeting is in the selected term
                    if not is_in_term(meeting['start_date'], selected_term):
                        continue
                    
                    # Create unique event identifier
                    event_key = f"{course_name}_{meeting['day']}_{meeting['start_time']}_{meeting['location']}"
                    if event_key in added_events:
                        continue
                    added_events.add(event_key)
                    
                    # Create event
                    event = Event()
                    # Add course type to summary if it's a lab
                    if course_type.lower() == 'laboratory':
                        event.add('summary', f"{course_name} (Lab)")
                    else:
                        event.add('summary', course_name)
                    
                    # Add detailed description
                    description = (
                        f"Course: {course_name}\n"
                        f"Type: {course_type}\n"
                        f"Section: {section}\n"
                        f"Location: {meeting['location']}\n"
                        f"Instructor: {instructor}"
                    )
                    event.add('description', description)
                    event.add('location', meeting['location'])
                    
                    # Parse dates and times
                    start_date = datetime.strptime(meeting['start_date'], '%Y/%m/%d')
                    end_date = datetime.strptime(meeting['end_date'], '%Y/%m/%d')
                    
                    # Get day number (0 = Monday)
                    day_map = {
                        'Mon': 0, 'Tue': 1, 'Wed': 2, 'Thu': 3, 'Fri': 4
                    }
                    day_num = day_map.get(meeting['day'])
                    if day_num is None:
                        continue
                    
                    # Find first occurrence of this day
                    current_date = start_date
                    while current_date.weekday() != day_num:
                        current_date += timedelta(days=1)
                    
                    # Parse times
                    try:
                        start_time = datetime.strptime(meeting['start_time'], '%H:%M').time()
                        end_time = datetime.strptime(meeting['end_time'], '%H:%M').time()
                    except ValueError:
                        print(f"Error parsing time for {course_name}")
                        continue
                    
                    # Combine date and time
                    event_start = datetime.combine(current_date.date(), start_time)
                    event_end = datetime.combine(current_date.date(), end_time)
                    
                    # Localize times
                    event_start = tz.localize(event_start)
                    event_end = tz.localize(event_end)
                    
                    event.add('dtstart', event_start)
                    event.add('dtend', event_end)
                    
                    # Set up weekly recurrence until the section end date
                    event.add('rrule', {
                        'freq': 'weekly',
                        'until': tz.localize(datetime.combine(end_date.date(), end_time)),
                        'byday': meeting['day'][:2].upper()
                    })
                    
                    cal.add_component(event)
                    print(f"Added {course_type} event for {course_name} on {meeting['day']}")
                    
                except Exception as event_error:
                    print(f"Error creating event for {course_name}: {event_error}")
                    continue
                    
        except Exception as course_error:
            print(f"Error processing course {course_name}: {course_error}")
            continue
    
    return cal

def process_excel_file(excel_file):
    """Process Excel file and return cleaned DataFrame"""
    try:
        df = pd.read_excel(excel_file, engine='openpyxl', header=2)
        df = df.dropna(how='all')
        print("\nDetected columns:", df.columns.tolist())
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        raise

def main():
    try:
        # List Excel files
        excel_files = [f for f in os.listdir() if f.endswith(('.xlsx', '.xls'))]
        for i, file in enumerate(excel_files, 1):
            print(f"{i}. {file}")
        
        # Select file
        choice = int(input("\nEnter file number: "))
        if not (1 <= choice <= len(excel_files)):
            print("Invalid selection")
            return
        
        excel_file = excel_files[choice - 1]
        print(f"\nProcessing {excel_file}...")
        
        # Select term
        print("\nSelect term:")
        print("1. Term 1 (Sep 3 - Dec 6)")
        print("2. Term 2 (Jan 6 - Apr 6)")
        term_choice = input("Enter term number (1 or 2): ")
        term = 'term1' if term_choice == '1' else 'term2'
        term_start, term_end = get_term_dates(term)
        
        # Process Excel file
        df = process_excel_file(excel_file)
        
        # Generate and save HTML calendar
        html_content = generate_course_calendar(df, term)
        with open('course_calendar.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
        print("\nHTML calendar generated successfully!")
        
        # Generate and save ICS calendar
        cal = generate_ics_calendar(df, term_start, term_end, term)
        with open('course_calendar.ics', 'wb') as f:
            f.write(cal.to_ical())
        print("ICS calendar generated successfully!")
        
        print("\nGenerated files:")
        print("1. course_calendar.html - Open in your web browser")
        print("2. course_calendar.ics - Import into your calendar application")
        
    except Exception as e:
        print(f"\nError: {str(e)}")
        print("\nTroubleshooting steps:")
        print("1. Make sure the Excel file contains required columns")
        print("2. Check if the meeting patterns are in the correct format")
        print("3. Try saving the Excel file with a different name")

if __name__ == "__main__":
    main()