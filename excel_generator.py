import os
from collections import defaultdict, namedtuple, OrderedDict
from xml.dom import minidom

import xlsxwriter

Case = namedtuple('Case', ['path', 'summary', 'precond', 'steps', 'expected', 'importance', 'keywords', 'exectype'])

IMPORTANCE = {'1': '低', '2': '中', '3': '高'}
EXECTYPE = {'1': '手工', '2': '自动'}
CASES = 'cases'


def tree():
    return defaultdict(tree)


def add_case(t, doms, module_path):
    for dom in doms:
        if dom.nodeName == "testsuite":
            if type(dom.parentNode.parentNode) == minidom.Document:
                module_path = []
            key = dom.getAttribute('name').strip()
            add_case(t, dom.childNodes, module_path + [key])

        elif dom.nodeName == "testcase":
            path = "/".join(module_path)
            summary = get_format_content(dom.getAttribute('name'))
            precond = get_cdata_node(dom.getElementsByTagName('preconditions'))
            steps = get_cdata_node(dom.getElementsByTagName('actions'))
            expected = get_cdata_node(dom.getElementsByTagName('expectedresults'))
            importance = IMPORTANCE[get_cdata_node(dom.getElementsByTagName('importance'))]
            keywords = get_cdata_node(dom.getElementsByTagName('keyword'))
            exectype = EXECTYPE[get_cdata_node(dom.getElementsByTagName('execution_type'))]
            case = Case(path, summary, precond, steps, expected, importance, keywords, exectype)

            t.setdefault(CASES, OrderedDict())
            if type(t[CASES]) == list:
                t[CASES].append(case)
            elif type(t[CASES]) == OrderedDict:
                t[CASES] = [case]


def get_cdata_node(node_list):
    if node_list:
        node_name = node_list[0].nodeName
        if node_name == 'actions' or node_name == 'expectedresults':
            format_list = []
            for item in node_list:
                child_nodes = item.childNodes
                if child_nodes:
                    node = [n for n in child_nodes if n.nodeType == 4]
                    format_list.append(get_format_content(node[0].nodeValue) if node else "")
            content = '\n'.join(format_list)
            return content

        elif node_name == 'keyword':
            format_list = []
            for item in node_list:
                format_list.append(get_format_content(item.getAttribute('name')))
            content = '\n'.join(format_list)
            return content

        else:
            child_nodes = node_list[0].childNodes
            if child_nodes:
                node = [n for n in child_nodes if n.nodeType == 4]
                return get_format_content(node[0].nodeValue) if node else ""
            else:
                return ""
    else:
        return ""


def get_format_content(content):
    return content.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&").replace("&apos;", "'").replace(
        "&quot;", "\"").replace("&nbsp;", "").replace("&mdash;", "").replace("&lsquo;", "‘").replace(
        "&rsquo;", "’").replace("&ldquo;", "“").replace("&rdquo;", "”").replace('<p>', '').replace('</p>', '').strip()


def read_xml_and_build_cases(xml_file):
    paths = []
    xml_dom = minidom.parse(xml_file)
    doc_tree = xml_dom.documentElement
    title = doc_tree.getAttribute('name').strip()
    root_tree = OrderedDict()
    root_tree.setdefault(title, OrderedDict())
    t = root_tree[title]

    testsuites = doc_tree.childNodes
    add_case(t, testsuites, paths)

    return root_tree


def generate_excel(dir_path, title, cases):
    excel_file = os.path.join(dir_path, title + '.xlsx')
    wkbook = xlsxwriter.Workbook(excel_file)
    sheets = wkbook.add_worksheet(title)
    title_style = wkbook.add_format({'font_name': 'SimSun', 'font_size': 16, 'bold': True})
    content_style = wkbook.add_format({'font_name': 'SimSun', 'font_size': 14})

    title_tag = ['路径', '用例标题', '前置条件', '操作步骤', '期望结果', '重要度', 'keywords', '执行方式']
    for i in range(len(title_tag)):
        sheets.write(0, i, title_tag[i], title_style)

    case_list = cases["cases"]
    row = 1
    for case in case_list:
        col = 0
        for item in case:
            sheets.write(row, col, item, content_style)
            col += 1
        row += 1
    wkbook.close()
    return excel_file
