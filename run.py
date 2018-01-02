#!/usr/bin/python3

# Standard imports
import glob
import logging
import os
from datetime import datetime, timedelta

# Third party imports (requirements.txt)
from bs4 import BeautifulSoup
import xlsxwriter

# Global variables
total = 0
workbook = xlsxwriter.Workbook('out/options.xlsx')


class Section:

    def __init__(self, num='', title='', state='', info=''):
        self.num = num
        self.title = title
        self.state = state
        self.info = info

        self.sem = ''
        self.lab = ''
        self.lec = ''
        self.dates = {}

    def parse_info(self):
        for string in self.info.split('\n'):
            substring = string.replace(', ', ',').split(" ")

            if substring[1] == 'LAB':
                start = clock12to24(substring[3])
                finish = clock12to24(substring[5].split(',')[0])
                self.lab = substring[2] + "||" + start + "-" + finish

            if substring[1] == 'SEM':
                start = clock12to24(substring[3])
                finish = clock12to24(substring[5].split(',')[0])
                self.sem = substring[2] + "||" + start + "-" + finish

            if substring[1] == 'EXAM':
                pass

            if substring[1] == 'LEC':
                start = clock12to24(substring[3])
                finish = clock12to24(substring[5].split(',')[0])
                self.lec = substring[2] + "||" + start + "-" + finish

    def print(self):
        if self.num != '':
            print("\t" + self.num + ".\t" + self.title)

    def format_section_times(self):
        days = {}

        if len(self.lec) > 1:
            lec = self.lec.split('||')
            for day in lec[0].split(','):
                if day not in days.keys():
                    days.setdefault(day, [lec[1]])
                else:
                    days[day].append(lec[1])

        if len(self.sem) > 1:
            sem = self.sem.split('||')
            for day in sem[0].split(','):
                if day not in days.keys():
                    days.setdefault(day, [sem[1]])
                else:
                    days[day].append(sem[1])

        if len(self.lab) > 1:
            lab = self.lab.split('||')
            for day in lab[0].split(','):
                if day not in days.keys():
                    days.setdefault(day, [lab[1]])
                else:
                    days[day].append(lab[1])

        return days


def is_valid_html_file(html):
    for string in html.find_all("h1"):
        if string.text.upper() == "Section Selection Results".upper():
            return True
    return False


def format_section_times(section):
    days = {}

    if len(section.lec) > 1:
        lec = section.lec.split('||')
        for day in lec[0].split(','):
            if day not in days.keys():
                days.setdefault(day, [lec[1]])
            else:
                days[day].append(lec[1])

    if len(section.sem) > 1:
        sem = section.sem.split('||')
        for day in sem[0].split(','):
            if day not in days.keys():
                days.setdefault(day, [sem[1]])
            else:
                days[day].append(sem[1])

    if len(section.lab) > 1:
        lab = section.lab.split('||')
        for day in lab[0].split(','):
            if day not in days.keys():
                days.setdefault(lab, [lab[1]])
            else:
                days[day].append(lab[1])

    return days


def can_be_together(section1, section2):

    # dates of the sections must be already formatted
    if len(section1.dates) <= 0 and len(section2.dates) <= 0:
        return False

    # print(section1.dates)
    # print(section2.dates)

    for day in section1.dates.keys():
        # if day is in both sections
        if day in list(section2.dates.keys()):
            # print(day)
            # Need to check times
            for time_1 in section1.dates[day]:

                if time_1 in list(section2.dates[day]):
                    return False

                time_c1 = time_1.replace(':', '').split('-')

                for time_2 in section2.dates[day]:
                    time_c2 = time_2.replace(':', '').split('-')
                    # print(time_1 + " | " + time_2)

                    if int(time_c1[1]) < int(time_c2[0]):
                        continue

                    if int(time_c2[1]) < int(time_c1[0]):
                        continue

                    return False

    return True


def clock12to24(time12):
    in_time = datetime.strptime(time12.replace(" ", "").strip(), "%I:%M%p")
    return datetime.strftime(in_time, "%H:%M")


def get_day_pos(day):
    day = day.lower()
    if day == 'mon':
        return 'B'
    if day == 'tues':
        return 'C'
    if day == 'wed':
        return 'D'
    if day == 'thur':
        return 'E'
    if day == 'fri':
        return 'F'
    return 'A'


def create_option_menu(sections):
    global workbook

    days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    day_format = workbook.add_format({'bold': True, 'align': 'center'})
    time_format = workbook.add_format({'bold': True, 'align': 'center'})

    worksheet = workbook.add_worksheet()
    worksheet.set_column('B:H', 20)

    for i in range(0, len(days_of_week)):
        worksheet.write(0, i + 1, days_of_week[i], day_format)

    start = datetime.strptime('7:30', "%H:%M")
    stop = datetime.strptime('22:30', "%H:%M")

    times = {}

    count = 1
    while start.time() != stop.time():
        start += timedelta(minutes=30)
        worksheet.write(count, 0, start.strftime("%I:%M %p"), time_format)
        count += 1

        times.setdefault(start, count)

    colours = ['#27AE60', '#2980B9', '#8E44AD', '#2C3E50', '#FLC40F', '#E74C3C', '#E67E22']

    count = 0
    for section in sections:

        node_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
        node_format.set_font_color('white')
        node_format.set_fg_color(colours[count])

        for key, values in section.dates.items():
            row = get_day_pos(key)
            if row != 'A':
                for value in values:
                    start = value.split('-')[0]
                    stop = value.split('-')[1]

                    loc_start = 0
                    loc_stop = 0

                    for ti, loc in times.items():
                        if ti.strftime("%H:%M") == start:
                            loc_start = loc

                        if (ti + timedelta(minutes=-10)).strftime("%H:%M") == stop:
                            loc_stop = loc - 1

                    worksheet.merge_range(row + str(loc_start) + ':' + row + str(loc_stop), section.title, node_format)

        count += 1
        if count == len(colours):
            count = 0


def main():
    logging.basicConfig(filename=None, level=logging.DEBUG)

    global total

    sections = {}
    lengths = {}

    exceptions = input("What exceptions would you like to add? ").split(' ')
    print(exceptions)

    if not os.path.exists("out"):
        os.makedirs("out")

    if not os.path.exists("html_courses"):
        os.makedirs("html_courses")

    for path in glob.glob(".\html_courses\*.html"):

        with open(path, 'r') as data:
            soup = BeautifulSoup(data.read(), 'html.parser')

            if not is_valid_html_file(soup):
                logging.error("File: " + path + " is not valid.")
            else:
                print("Reading file: '" + path + "'")

                for row in soup.find('div', {'id': 'GROUP_Grp_WSS_COURSE_SECTIONS'}).find_all('tr'):

                    section = Section()
                    for subsection in row.find_all('td'):

                        subclasses = subsection.get('class')

                        if 'windowIdx' in subclasses:
                            if subsection.text == '':
                                break
                            else:
                                section.num = subsection.text.rstrip('\n')

                        if 'LIST_VAR1' in subclasses or 'LIST_VAR2' in subclasses:
                            if subsection.text.lower().strip() == 'open' or subsection.text.lower().strip() == 'closed':
                                section.state = subsection.text.strip()

                        if 'SEC_SHORT_TITLE' in subclasses:
                            section.title = subsection.text.strip().split(" ")[0]

                        if 'SEC_MEETING_INFO' in subclasses:
                            section.info = subsection.text.strip()
                            section.parse_info()

                            if len(section.state) > 0:

                                title = section.title.split("*")
                                title = title[0] + "*" + title[1]

                                if section.state.upper() == 'OPEN' or section.title.upper() in exceptions:

                                    section.dates = section.format_section_times()

                                    if title in sections.keys():
                                        sections.get(title).append(section)
                                        lengths[title] += 1
                                    else:
                                        sections.setdefault(title, [section])
                                        lengths.setdefault(title, 1)

                                else:
                                    if title not in sections.keys():
                                        sections.setdefault(title, [])
                                        lengths.setdefault(title, 0)

    num_courses = len(sections.keys())
    print("\nNum of courses: " + str(num_courses))

    for key, values in sections.items():
        print("===== " + key + " (" + str(len(values)) + ") ======")
        for value in values:
            value.print()

    work(sections)

    print("\nTotal options found: " + str(total))


def work(sections, current=0, courses=None, path=[]):
    global total

    if courses is None:
        courses = list(sections.keys())

    if current >= len(courses):

        if len(path) != 0:
            for element1 in path:
                for element2 in path:
                    if can_be_together(element1, element2) or element2 == element1:
                        continue
                    else:
                        return

        # print(total)
        # for s in path:
        #    s.print()

        total += 1
        create_option_menu(path)
        return

    for section in sections[courses[current]]:
        path.append(section)
        work(sections, current + 1, courses, path)
        del path[-1]


if __name__ == '__main__':
    main()
    workbook.close()
