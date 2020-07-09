import os
from collections import defaultdict, namedtuple, OrderedDict
from xml.dom.minidom import Document

import xlrd

Case = namedtuple('Case', ['summary', 'precond', 'exectype', 'steps', 'expected', 'importance', 'keywords'])

IMPORTANCE = {'低': '1', '高': '3', '中': '2'}
EXECTYPE = {'手工': '1', '自动': '2'}
CASES = 'cases'


def tree():
    return defaultdict(tree)


def add_case(t, path, case):
    if path != '':
        paths = path.split('/')
        for item in paths:
            t.setdefault(item, OrderedDict())
            t = t[item]
    t.setdefault(CASES, OrderedDict())
    if type(t[CASES]) == list:
        t[CASES].append(case)
    elif type(t[CASES]) == OrderedDict:
        t[CASES] = [case]


def read_excel_and_build_trees(excel_file):
    wkbook = xlrd.open_workbook(excel_file)
    sheets = wkbook.sheets()
    result = {}

    for sheet in sheets:
        if sheet.name != '模板说明':
            root_tree = OrderedDict()
            root_tree.setdefault(sheet.name, OrderedDict())
            t = root_tree[sheet.name]

            for row in range(1, sheet.nrows):
                path = sheet.cell_value(row, 0).strip()
                summary = sheet.cell_value(row, 1).strip()
                if summary == '':
                    continue

                precond = sheet.cell_value(row, 2).strip()
                steps = sheet.cell_value(row, 3).strip()
                expected = sheet.cell_value(row, 4).strip()
                importance = sheet.cell_value(row, 5).strip()
                importance = IMPORTANCE.get(importance, '3')
                keywords = sheet.cell_value(row, 6).strip()
                keywords = keywords.split(',') if keywords != '' else []
                exectype = sheet.cell_value(row, 7).strip()
                exectype = EXECTYPE.get(exectype, '1')

                case = Case(summary, precond, exectype, steps, expected, importance, keywords)
                add_case(t, path, case)

            result[sheet.name] = root_tree
        else:
            pass

    return result


def fxml(text: str):
    return text.replace('<', '&lt;').replace('>', '&gt;')


def generate_testcase(doc, case):
    testcase = doc.createElement('testcase')
    testcase.setAttribute('name', case.summary)

    # summary
    summary = doc.createElement('summary')
    text = doc.createCDATASection(fxml(case.summary))
    summary.appendChild(text)
    testcase.appendChild(summary)

    # precond
    precond = doc.createElement('preconditions')
    text = doc.createCDATASection(fxml(case.precond))
    precond.appendChild(text)
    testcase.appendChild(precond)

    # execution type
    exectype = doc.createElement('execution_type')
    text = doc.createCDATASection(case.exectype)
    exectype.appendChild(text)
    testcase.appendChild(exectype)

    # importance
    importance = doc.createElement('importance')
    text = doc.createCDATASection(case.importance)
    importance.appendChild(text)
    testcase.appendChild(importance)

    # steps
    if case.steps != '' and case.expected != '':
        steps = doc.createElement('steps')
        step = doc.createElement('step')

        step_number = doc.createElement('step_number')
        text = doc.createCDATASection('1')
        step_number.appendChild(text)
        step.appendChild(step_number)

        if case.steps != '':
            actions = doc.createElement('actions')
            lst = case.steps.split('\n')
            texts = []
            for line in lst:
                line = '<p>' + fxml(line.strip()) + '</p>'
                texts.append(line)
            texts = '\n'.join(texts)

            text = doc.createCDATASection(texts)
            actions.appendChild(text)
            step.appendChild(actions)

        if case.expected != '':
            expected = doc.createElement('expectedresults')
            lst = case.expected.split('\n')
            texts = []
            for line in lst:
                line = '<p>' + fxml(line.strip()) + '</p>'
                texts.append(line)
            texts = '\n'.join(texts)

            text = doc.createCDATASection(texts)
            expected.appendChild(text)
            step.appendChild(expected)

        exectype = doc.createElement('execution_type')
        text = doc.createCDATASection(case.exectype)
        exectype.appendChild(text)
        step.appendChild(exectype)

        steps.appendChild(step)
        testcase.appendChild(steps)

    # keywords
    if len(case.keywords) > 0:
        keywords = doc.createElement('keywords')
        for kw in case.keywords:
            keyword = doc.createElement('keyword')
            keyword.setAttribute('name', kw)
            keywords.appendChild(keyword)
        testcase.appendChild(keywords)

    return testcase


def generate_testsuit(doc, name):
    testsuit = doc.createElement('testsuite')
    testsuit.setAttribute('name', name)
    return testsuit


def generate_recursion(doc, elem, t):
    for k in t:
        v = t[k]
        if type(v) == OrderedDict:
            testsuit = generate_testsuit(doc, k)
            elem.appendChild(testsuit)
            generate_recursion(doc, testsuit, v)
        else:
            for case in v:
                testcase = generate_testcase(doc, case)
                elem.appendChild(testcase)


def generate_xml(xml_dir, name, t):
    doc = Document()
    generate_recursion(doc, doc, t)

    xml_file = os.path.join(xml_dir, name + '.xml')
    with open(xml_file, 'w', encoding='utf-8') as f:
        doc.writexml(f, indent='\t', newl='\n', addindent='\t', encoding='utf-8')
    return xml_file
