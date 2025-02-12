import pandas as pd
import xlsxwriter
from collections import defaultdict

workbook = None
sections = []   # sections = {sect, submissions, percent, target}
question_categories = []   # e.g. ['coding', 'diagramming']
question_groups = {}    # groups = {[q1, q2], [q1%, q2%]}

color_list = ['#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc', '#efefef']
color_formats = []

def parseInputSheet(sheet):
    """ Goes through the input excel sheet and pulls information from it.

    Returns
        List[TA]: List of ta names and their grading weights = {'name', 'percent'}
    """

    global sections, question_groups, question_categories

    input = pd.read_excel(sheet)
    name_percent_df = pd.DataFrame({
        'name': input['TA Name'],
        'percent': input['% to Grade']
    })
    name_percent_df['percent'] = name_percent_df['percent'].fillna(100.0)
    name_percent_df = name_percent_df.sort_values(by='percent', ascending=False).reset_index(drop=True)

    # tas = {name, percent}
    tas = dict(zip(name_percent_df['name'], name_percent_df['percent']))
    print(tas)

    # looks at 'category ->' row and finds the names
    cat_row = input.iloc[0, 10:].tolist()
    for cat in cat_row:
        if isinstance(cat, str):
            question_categories.append(cat)
    num_categories = len(question_categories)
    print("question_categories:", question_categories)

    # get section, section submission numbers
    total_submissions = 0
    num_sections = int(input.iloc[0, 4])
    for i in range(num_sections):
        submissions = input.iloc[i, 7]
        total_submissions += submissions

        sections.append({
            'sect': input.iloc[i, 6],
            'submissions': submissions,
            'percent': 0,
        })
    
    # calculate each section's weight based on its submission compared to total submissions
    for section in sections:
        section['percent'] = round((section['submissions'] / total_submissions), 2)
    print("sections:", sections)

    # TODO: fix here
    col = 10
    # question_groups = {[q1, q2], [q1 percent, q2 percent]}
    for i in range(num_categories):
        cat = question_categories[i]
        qp = []

        questions = input.iloc[1:, col].dropna().tolist()
        qp.append(questions)

        percents = input.iloc[1:, col+1].dropna().tolist()
        qp.append(percents)
        
        question_groups[cat] = qp
        col += 2

    print("question_groups:", question_groups)
    return tas

def distribute_integer_parts(total, percentages):
    """Distributes total into integer parts based on relative percentages while preserving sum."""
    sum_percentages = sum(percentages)
    normalized_percentages = [p / sum_percentages for p in percentages]
    print(percentages)

    int_parts = [int(total * pct) for pct in normalized_percentages]
    remainders = [(total * pct) - int_parts[i] for i, pct in enumerate(normalized_percentages)]
    
    remaining = total - sum(int_parts)

    # distribute the remaining units to groups with the highest remainders
    for i in sorted(range(len(remainders)), key=lambda i: -remainders[i]):
        if remaining > 0:
            int_parts[i] += 1
            remaining -= 1
        else:
            break

    return int_parts

def groupTAs(tas):
    """Allocates names to each section while balancing total weight."""

    print("-------------- Grouping TAs --------------")
    num_sections = len(sections)
    total_weight = sum(tas.values())

    section_percentages = []    # percentages of each section
    for i in range(num_sections):
        section_percentages.append(sections[i]['percent'])

    # compute target integer weights for each section
    for i in range(num_sections):
        sections[i]['target_weight'] = (total_weight * sections[i]['percent'])

    # sort names by weight descending for better allocation
    sorted_names = sorted(tas.items(), key=lambda item: item[1], reverse=True)

    # assign names to sections while balancing total weight
    section_allocations = []
    for i in range(num_sections):
        section_allocations.append([])
    section_weights = [0] * num_sections

    for name, weight in sorted_names:
        # find the section that is the most under its target weight
        target_sect = max(range(num_sections), key=lambda g: sections[g]['target_weight'] - section_weights[g])
        section_allocations[target_sect].append({'name': name, 'percent': weight})
        section_weights[target_sect] += weight

    # sort alphabetically
    for i in range(num_sections):
        for item in section_allocations[i]:
            item['name'] = item['name'].title()

    for i in range(num_sections):
        print(section_allocations[i])

    return section_allocations

# tas = {name, percent}
# question_group = {bst, avls} or {skiplist, heaps}
def calcRanges(tas, section, question_group):
    print("------------------ calculating ranges ------------------")
    
    # split based on question weights
    question_percents = question_group[1]
    question_split = distribute_integer_parts(section['target_weight'], question_percents)
    print("question_split:", question_split)
    
    # groups per question
    groups = []
    for _ in range (len(question_percents)):
        groups.append([])
    group_totals = [0] * len(question_percents)

    for i in range(len(tas)):
        per = tas[i]['percent']

        min_diff_index = min(range(len(question_percents)), 
                             key=lambda j: (group_totals[j] + int(per)) - question_split[j])
        groups[min_diff_index].append(tas[i])
        group_totals[min_diff_index] += per

    # ranges = {ta, start, # of submissions, question}
    ranges = []
    STA_submissions = 8
    for i in range(len(groups)):
        sub = 1
        sect_submissions = section['submissions']

        if section['sect'] == 'Version A':
            sub += STA_submissions 
            sect_submissions -= STA_submissions

        groups[i] = sorted(groups[i], key=lambda item: item['name'])
        percents = [ta['percent'] / 100 for ta in groups[i]]
        per_ta_splits = distribute_integer_parts(sect_submissions, percents)

        for j in range(len(groups[i])):
            if (j == 0 and section['sect'] == 'Version A'):
                 ranges.append({
                    'ta': {'name': 'STA'},
                    'start': 1,
                    'num_submissions': STA_submissions,
                    'question': i
                })

            ranges.append({
                'ta': groups[i][j],
                'start': sub,
                'num_submissions': per_ta_splits[j],
                'question': i
            })

            sub += per_ta_splits[j]

    print(ranges)
    return ranges

"""Generates format with a new color for each question"""
def generate_formats():
    # find max num of questions for any category
    max_num = 0
    for cat in question_categories:
        questions = question_groups[cat][0]
        max_num = max(len(questions), max_num)

    formats = []
    color_idx = 0
    for _ in range(len(sections)):
        for _ in range(max_num):
            set = []
            set.append(workbook.add_format({
                'bg_color': color_list[color_idx],
            }))
            set.append(workbook.add_format({
                'bg_color': color_list[color_idx],
                'align': 'center'
            }))
            set.append(workbook.add_format({
                'bg_color': color_list[color_idx],
                'align': 'center',
                'font_size': 12,
            }))
            formats.append(set)
            color_idx += 1
    return formats

def printSheets(sheet, ranges, category, sect_idx):
    title_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'font_size': 15,
        'bg_color': '#4285f4',
        'color': 'white'
    })

    question_title_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'font_size': 13,
        'bg_color': '#4285f4',
        'color': 'white'
    })

    section = sections[sect_idx]['sect']

    # headers
    sheet.write(0, 0, 'TA Name', title_format)
    sheet.set_column('A:A', 26)
    sheet.write(0, 1, 'First', title_format)
    sheet.set_column('B:B', 10)
    sheet.write(0, 2, 'Last', title_format)
    sheet.set_column('C:C', 10)
    sheet.write(0, 3, '# of Submissions', title_format)
    sheet.set_column('D:D', 26)
    sheet.freeze_panes(1, 0)

    # print ranges
    row = 1
    for i in range(len(ranges)):
        start = ranges[i]['start']
        num_submissions = ranges[i]['num_submissions']

        question = ranges[i]['question']
        color_idx = (sect_idx * 2) + question
        name_format = color_formats[color_idx][0]
        submission_format = color_formats[color_idx][1]

        sheet.write(row, 0, ranges[i]['ta']['name'], submission_format)
        sheet.write(row, 1, start, submission_format)
        sheet.write(row, 2, start + num_submissions - 1, submission_format)
        sheet.write(row, 3, ranges[i]['num_submissions'], submission_format)
        row += 1

    row = 3
    sheet.write(row, 5, f"{section.upper()} QUESTIONS", question_title_format)
    sheet.set_column('F:F', 36)

    for i in range(len(question_groups)):
        color_idx = (sect_idx * 2) + i
        question_format = color_formats[color_idx][2]

        row += 1
        sheet.write(row, 5, f"{question_groups[category][0][i]} {category.lower().capitalize()}", question_format)

def createSheets(tas):
    global workbook, color_formats, sections
    tas_per_section = groupTAs(tas)

    color_formats = generate_formats()

    # make num ranges for each group
    for sect_idx in range(len(sections)):     # go through each section

        # new page and ranges for diff categories (e.g. coding campus a, diagramming campus a)
        for cat_idx in range(len(question_categories)):
            cat = question_categories[cat_idx]

            sheet_name = f"{cat} {sections[sect_idx]['sect']}"   
            sheet = workbook.add_worksheet(sheet_name)

            sheet_ranges = calcRanges(tas_per_section[sect_idx], sections[sect_idx], question_groups[cat])
            printSheets(sheet, sheet_ranges, cat, sect_idx)

def main():
    tas = parseInputSheet('input.xlsx')
    outputFileName = 'GradingAssignments.xlsx'

    global workbook
    workbook = xlsxwriter.Workbook(outputFileName)
    createSheets(tas)
    workbook.close()

if __name__ == "__main__":
    main()