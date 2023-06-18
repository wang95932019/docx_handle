import docx.opc.exceptions
import docx
import base64
import re
import os
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.shape import CT_Picture
from docx.image.image import Image
from docx.parts.image import ImagePart
from PIL import Image
import config

fmt = {  # 自动编号form类型字典，w:numFmt 对应的格式化样式
    'japaneseCounting': ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'],
    'chineseCounting': ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'],
    'decimal': ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '1'],
    'decimalEnclosedCircle': ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '1'],
    'bullet': ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '1']
}

# 例如：[6, 7, 9, 13, 15, 8, 16, 2, 11, 17, 5, 0, 18, 1, 4, 10, 3, 14, 12]
numId_of_abstractId = []  # numId对应的abstractId数组

# 存放 ‘列表段落样式’ 的二维列表，子列表格式：[w:start,w:numFmt,w:lvlText]，例如：[[13, 'chineseCounting', '%1、']]
list_numXML = []  # numXML列表

result_key_list = {}


def log(*msg):
    print(msg)


def image_to_base64(image: Image) -> str:
    data = image.blob  # 二进制数据
    fmt = image.ext  # 后缀
    base64_str = base64.b64encode(data).decode('utf-8')
    return f'data:image/{fmt};base64,' + base64_str


def ini_document(doc):
    '''初始化docx文件，获取docx文件内的其他数据'''
    global numId_of_abstractId, list_numXML
    try:
        ct_numbering = doc.part.numbering_part._element
    except:
        return
    for num in ct_numbering.num_lst:
        # 获取numId和abstractNmuId的对应关系
        numId_of_abstractId.append(num.abstractNumId.val)
    ns = ct_numbering.nsmap
    xmlns = ns['w']
    prefix_name = '{' + xmlns + '}'
    for i in ct_numbering.findall(prefix_name + 'abstractNum'):
        # 获取每个abstractNumId里面的每个ilvl里的lvlText，numFmt，start
        tmp_tmp = []
        for k in i.findall(prefix_name + 'lvl'):
            tmp = []
            for j in list(k):
                if j.tag.replace(prefix_name, '').strip() == 'start':
                    tmp.append(int(j.get(prefix_name + 'val')))
                if j.tag.replace(prefix_name, '') == 'numFmt' or j.tag.replace(prefix_name, '') == 'lvlText':
                    tmp.append(j.get(prefix_name + 'val'))
                # tmp_tmp.append(tmp)
            # BUG修改循环次数太多导致，出现大量重复的列表，如：[1, 'decimal', '%1.']
            tmp_tmp.append(tmp)
        list_numXML.append(tmp_tmp)


def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph. *parent*
    would most commonly be a reference to a main Document object, but
    also works for a _Cell object, which itself can contain paragraphs and tables.
    """
    if isinstance(parent, docx.document.Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def replace_wrong_char(txt):
    '''
    替换掉可能存在问题的字符
    【1】, -> ，
    '''
    tmp_list = [[',', '，']]
    for i in tmp_list:
        txt = txt.replace(i[0], i[1])
    return txt


def del_before_text(txt, rules):
    '''
    删除正文之前的内容，避免其可能造成的影响
    '''
    tmp = re.search(rules[0][0], txt)
    start = re.search(rules[0][0], txt).start()
    return txt[start:]



def get_picture(document: docx.document.Document, img_list: list):
    '''获取Paragraph段内的图片'''
    result_list = []
    for i in range(len(img_list)):
        img: CT_Picture = img_list[i]
        embed = img.xpath('.//a:blip/@r:embed')[0]
        related_part: ImagePart = document.part.related_parts[embed]
        image: Image = related_part.image
        result_list.append(image)
    return result_list


def deal_to_create_text(num_Id, ilvl_id):
    '''根据自动序号numId和abstractId，创建自动序号文本'''
    return list_numXML[numId_of_abstractId[num_Id - 1]][ilvl_id]


def set_style_number_list_paragraph(paragraph, position: int):
    """
        自动编号。设置列表段落对应的样式与编号
            1. 通过获取的w:numId和w:ilvl，联合list_numXML，设定列表的样式
            2. 通过获取的w:start与形参position，设置列表序号
        paragraph：当前段落对象
        position：当前段落在当下列表中的位置，从1开始
    """

    numId = paragraph._element.pPr.numPr.numId.val
    ilvl = paragraph._element.pPr.numPr.ilvl.val
    replace_str_old = '%' + str(ilvl + 1)
    tmp_list = list_numXML[numId_of_abstractId[numId - 1]][ilvl]  # [1, 'decimal', '(%1)']
    num = tmp_list[0] + position - 1  # 当前段落在当下列表中的编号。编号 = w:start(开始编号) + position(位置) - 1

    # 将像 %1 这样的占位符，改为对应样式的数字如：在w:numFmt=decimal下：(%1) -> (1)
    if num <= 10:  # 检测列表序号是否大于10，如果大于需要进行对应处理，否则会报IndexError: list index out of range，即数组越界的错误
        replace_str_new = fmt[tmp_list[1]][num]
    else:
        replace_str_new = fmt[tmp_list[1]][10] + fmt[tmp_list[1]][num - 10]

    tmp_str = tmp_list[2]
    tmp_str = tmp_str.replace(replace_str_old, replace_str_new)
    return tmp_str


def check_not_table_name(text):
    '''检查是否没有表n-n-n，暂时废弃'''
    if re.match('\s?表\s?[1-9]\s?-[1-9]\s?-[1-9]', text) is None:
        return True
    return False


def format_table(table: Table):
    '''将表格原始xml数据转为二维数组'''
    data = []
    for row in table.rows:
        data.append([])
        for cell in row.cells:
            data[-1].append(cell.text)
    return data


def handle_docx(path):
    '''传入docx文件路径，返回格式化后的纯文本数据、表格数据和图片数据
    参数:
    path (文档相对路径)
    返回:
    code (状态编号)
    text (纯文本数据)
    table (全部表格)
    '''
    # try:
    #     doc = docx.Document(path)  # 文档路径
    #     ini_document(doc)
    #     tables_list_all = []
    #     images_list_all = []
    #     position = 0  # 存放当前列表段落，在当下列表中的位置
    #     text = ''
    #     for block in iter_block_items(doc):
    #         if isinstance(block, Paragraph):  # 判断block是不是一个段落（Paragraph），如果不是下面再判断是不是表格（Table）
    #             img_list = block._element.xpath('.//pic:pic')
    #             if (block.text != '' or len(img_list) != 0) and not re.search('\t\\d*$', block.text):       # not re.search('\t\\d*$', block.text)：不是目录
    #
    #                 # 判断是否属于同一个列表的指标：列表段落是否中断
    #                 if block._element.pPr is not None and block._element.pPr.numPr is not None:
    #                     if position != 0:
    #                         position += 1
    #                     else:
    #                         position = 1
    #                     auto_number = set_style_number_list_paragraph(block, position)
    #                     text += auto_number
    #                 else:
    #                     position = 0
    #
    #                 text += block.text.strip() + '\n\n'
    #                 img_list = block._element.xpath('.//pic:pic')
    #                 if len(img_list) != 0 or img_list:
    #                     images_list = get_picture(doc, img_list)
    #                     for image in images_list:
    #                         if image:
    #                             images_list_all.append(image)
    #                             text += f'[images#{len(images_list_all)}]\n'
    #                             # res = image_to_base64(image.blob, image.ext)
    #                             # print(res)
    #                             # ext = image.ext  # 后缀
    #                             # blob = image.blob  # 二进制内容
    #         elif isinstance(block, Table):
    #             tables_list_all.append(block)
    #             text += f'[tables#{len(tables_list_all)}]\n'
    #         else:
    #             log("Wrong")

    # except docx.opc.exceptions.PackageNotFoundError:
    doc = docx.Document(path)  # 文档路径
    ini_document(doc)
    tables_list_all = []
    images_list_all = []
    position = 0  # 存放当前列表段落，在当下列表中的位置
    text = ''
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):  # 判断block是不是一个段落（Paragraph），如果不是下面再判断是不是表格（Table）
            img_list = block._element.xpath('.//pic:pic')
            if (block.text != '' or len(img_list) != 0) and not re.search('\t\\d*$',
                                                                          block.text):  # not re.search('\t\\d*$', block.text)：不是目录

                # 判断是否属于同一个列表的指标：列表段落是否中断
                if block._element.pPr is not None and block._element.pPr.numPr is not None:
                    if position != 0:
                        position += 1
                    else:
                        position = 1
                    auto_number = set_style_number_list_paragraph(block, position)
                    text += auto_number
                else:
                    position = 0

                text += block.text.strip() + '\n\n'
                img_list = block._element.xpath('.//pic:pic')
                if len(img_list) != 0 or img_list:
                    images_list = get_picture(doc, img_list)
                    for image in images_list:
                        if image:
                            images_list_all.append(image)
                            text += f'[images#{len(images_list_all)}]\n'
                            # res = image_to_base64(image.blob, image.ext)
                            # print(res)
                            # ext = image.ext  # 后缀
                            # blob = image.blob  # 二进制内容
        elif isinstance(block, Table):
            tables_list_all.append(block)
            text += f'[tables#{len(tables_list_all)}]\n'
        else:
            log("Wrong")
    #     return 404, '文本读取失败', ['列表读取失败'], ['图片读取失败']
    return 200, text, tables_list_all, images_list_all


def handle_text(text):
    '''处理文本(删除前后换行和空格)'''
    return text.strip('\n').strip()


def check_key_num(txt, rule):
    '''返回标题的开始位置'''
    res = re.search(rule, txt)
    if res is not None:
        return res.start()
    return 999999


def check_key_name(rules: list, name):
    for i in range(len(rules)):
        # 添加正则，会增大匹配的范围，未必是正确的做法
        if re.search(rules[i][0], name):
            return i
        # if rules[i][0] == name:
        #     return i
    return -1


def handle_dfs(rules: list, txt: str, title: str = '', num_ceng: int = 0, new_rules: list = [], key_rules: list = [],
               key_num: int = -1, new_num: int = -1):
    '''
    rules: 不同层级的标题的匹配规则。层级从0开始
    txt：文本内容
    title：标题。title与txt是对应关系，第一次传入的应为 docx_txt和‘’。
    num_ceng：当下所处的层数
    new_rules：标题新的匹配规则。格式为[["key_name", rules]]
    key_rules：标题取到第几层的规则。格式为[["key_name", number]] number : 0 直接返回 1 解析一层 2 解析两层 n 解析n层
    key_num：  key_rules的第一个下标，默认为-1
    new_num：  new_rules的第一个下标，默认为-1
    '''

    # tmp_list_docx:dict = {}     # 字典用来存放处理结果，字典便于转为json

    # 下面if-elif的部分是用来，对 key_rules或new_rules 进行处理的，二者不可以同时生效，同时使用时会生效 key_rules
    if key_rules:  # 使用key_rules
        if key_num > -1:  # 如果 key_num 不是-1，就会取key_rules中规定的层数与num_ceng进行比较
            if num_ceng >= key_rules[key_num][1]:  # 当递归解析的层数已经到达，规定的层数时，就不会向下解析
                return txt.replace('\n\n', '\n')
        else:
            # 判断程序是否解析到了要使用key_rules的标题
            # 将key_rules中的标题与递归传递的标题进行对比。相等就返回title所在列表在key_rules中的下标，不相等就返回-1
            check_result = check_key_name(key_rules, title)
            if check_result > -1:
                if key_rules[check_result][1] == 0:  # 如果key_rules中设置的层数为0，之间返回结果
                    result_key_list[title] = txt.replace('\n\n', '\n')
                    tmp_list_key = txt.replace('\n\n', '\n')
                else:
                    result_key_list[title] = handle_dfs(rules[:key_rules[check_result][1]], txt, title, num_ceng=0,
                                                        key_rules=key_rules, key_num=check_result)
                    tmp_list_key = handle_dfs(rules[:key_rules[check_result][1]], txt, title, num_ceng=0,
                                              key_rules=key_rules, key_num=check_result)
                return tmp_list_key
    elif new_rules:  # 原理同上
        if new_num > -1:
            if num_ceng >= len(new_rules[new_num][1]) - 1:
                return txt.replace('\n\n', '\n')
        else:
            check_result = check_key_name(new_rules, title)
            if check_result > -1:
                result_key_list[title] = handle_dfs(new_rules[new_num][1], txt, title, num_ceng=0, new_rules=new_rules,
                                                    new_num=check_result)
                tmp_list_key = handle_dfs(new_rules[new_num][1], txt, title, num_ceng=0, new_rules=new_rules,
                                          new_num=check_result)
                return tmp_list_key

    tmp_list_docx = {}  # 字典用来存放处理结果，字典便于转为json
    tmp_list = []
    # 将while修改为 for,因为我不习惯使用 while
    for rule in rules:
        # 注意：此处作为，正则表达式的是每个rule中的第0条，如：'第[一二三四五六七八九十][一二三四五六七八九]?部分'
        tmp_list.append(check_key_num(txt, rule[0]))

    if sum(tmp_list) == len(rules) * 999999:  # 如果当前内容所有层级的标题都没有匹配
        if txt.find('\n\n') > -1:
            return txt.split('\n\n')
        return txt

    # 获取标题层级：获取匹配到的编号最小的（离开头最近的）标题的层级
    # 层级从0开始。即当匹配rules中的最高等级的标题时，num=0
    num = tmp_list.index(min(tmp_list))
    # res_tmp:list 中存放的是当前内容下的所有标题对应的内容。称之为内容列表
    # title_:list 中存放的是当前内容下的所有标题。称之为标题列表
    # 此处通过正则分割字符串，标准为 rules[num][1] 中num等级的标题，num从0开始
    res_tmp: list = re.split(rules[num][1], '\n' + txt + '\n')
    # 找到所有 num 等级的标题
    title_: list = re.findall(rules[num][2], '\n' + txt + '\n')
    # 如果当下标题处于第0层，移除列表中首元素
    # 目的：删除非正文的内容
    # if num_ceng == 0:
    #     res_tmp.pop(0)
    if len(title_) == 0:  # 如果当前标题下面没有子标题，就直接返回文本。如：培养目标
        return res_tmp

    # 遍历内容列表，如果内容为空，或者为\n，或者\n\n，就删除这个元素。
    # 为空不能删除！！！ 因为会有一个标题下面全部是图片的。删了和标题列表与内容列表就对应不上了，除非二者同时删除对应下标的内容。
    for i in range(len(res_tmp) - 1, -1, -1):  # 步长为-1，倒着遍历
        if res_tmp[i] == '' or res_tmp[i] == '\n' or res_tmp[i] == '\n\n':  # 如果内容为空，或者为\n，或者\n\n，就删除这个元素。
            res_tmp.pop(i)
    #         #title_.pop(i)

    if len(res_tmp) == 0:  # 如果标题下面没有内容，就将标题视为内容。如：素质要求下面的各条要求
        return title_
    # 如果标题列表的长度比内容列表小1，就在标题列表的开头中插入当前标题
    # 目的：为了应对标题下面不是标题而是之间出现文本
    if len(title_) + 1 == len(res_tmp):
        title_.insert(0, title)

    # 遍历内容列表与标题列表
    for i, j in zip(res_tmp, title_):
        title_new = handle_text(j)  # handle_text():text.strip('\n').strip()
        # 递归：num_ceng + 1
        tmp_list_docx[title_new] = handle_dfs(rules, handle_text(i), title_new, num_ceng=num_ceng + 1,
                                              key_rules=key_rules, new_rules=new_rules)
    return tmp_list_docx


def transpose_2d(data):
    """转置二维数组"""
    transposed = list(map(list, zip(*data)))
    return transposed


def handle_tables(tables_list, table_rules_list):
    result_list = []
    for rule in table_rules_list:
        tmp_list = []
        for table in tables_list:
            find = False
            tmp_num_list = []
            if rule[0] == 0:
                list_ = table
            else:
                list_ = transpose_2d(table)
            for row in list_:
                if sum(map(row.count, rule[1])) >= len(rule[1]):
                    find = True
                    tmp_num_list = list(map(row.index, rule[1]))
                    tmp_list = []
                # print(find, len(row))
                if find:
                    if rule[2] == 0:
                        # print(row)
                        # print(tmp_num_list)
                        tmp_row = []
                        for num in tmp_num_list:
                            tmp_row.append(row[num])
                        tmp_list.append(tmp_row)
                    else:
                        if len(tmp_list) <= rule[2]:
                            tmp_row = []
                            for num in tmp_num_list:
                                tmp_row.append(row[num])
                            tmp_list.append(tmp_row)
                        else:
                            find = False
                            break
            if tmp_list:
                result_list.append(tmp_list)
                tmp_list = []
    return result_list


def handle_path(path):
    file_list = os.listdir(path)
    for file in file_list:
        if file[-5:] == '.docx' and file[0] != '~':
            cur_path = os.path.join(path, file)
            res_path = cur_path.replace('docx', 'json')


def main(path, rules, depth: int = 99999, key_rules: list = [], new_rules: list = [], table_key_rules: list = []):
    """
    path：文档路径
    rules：匹配标题规则，其中每一个列表，对应一种标题种类
    depth：深度，默认99999，即深度不限，直到所有深度都解析完成
    key_rules
    """

    # 大致流程：解析文档（列表编号）， -> 对文本字符串进行处理 -> 根据 rules 将text由str变换为json
    # 1. 解析文档（列表编号）
    result_code, docx_text, tables_list, images_list = handle_docx(
        path)  # 解析文档，通过word xml的方式获取文档内容，并进行简单的处理，其中docx_text是字符串，格式为 大数据技术专业\n\n2022级人才培养方案\n\n
    if result_code != 200:
        return result_code, docx_text, tables_list, images_list

    # 2.1 替换干扰字符（，）
    docx_text = replace_wrong_char(docx_text)

    # 2.2 删除正文之前的内容
    docx_text = del_before_text(docx_text, rules)

    # 3. 根据 rules 将text由str变换为json
    res_text_list = handle_dfs(rules, docx_text, key_rules=key_rules, new_rules=new_rules)
    res_tables_list = []
    res_images_list = []
    result_table_key_list = []
    for _ in tables_list:
        res_tables_list.append(format_table(_))
    for _ in images_list:
        res_images_list.append(image_to_base64(_))
    if table_key_rules:
        result_table_key_list = handle_tables(res_tables_list, table_key_rules)
    return result_code, res_text_list, res_tables_list, res_images_list, result_key_list, result_table_key_list


def test():
    # skz_rules = config.skz_rules
    # sz_rules = config.sz_rules
    # skz_new_rules = config.skz_new_rules
    #
    # path1 = "test_docx/大数据技术专业人培_培养目标和培养规格.docx"
    # path2 = "test_docx/大数据技术专业人培.docx"
    # path3 = "test_docx/2020级服装与服饰设计专业（三+二本科）人才培养方案.docx"
    # path4 = "test_docx/大数据技术专业人培_十and十一and十二.docx"
    # path5 = "test_word_XML/test.docx"
    # path6 = "test_docx/市场营销专业2022级人才培养方案.docx"
    # path7 = "test_docx/市场营销专业2022级人才培养方案_无附件.docx"
    # path8 = "test_docx/市场营销_十三and十四.docx"
    # result_code, text_res, res_tables_list, res_images_list, result_key_list, result_table_key_list = main(
    #     path8, sz_rules, 99999, [], [])
    # res = handle_tables(res_tables_list, [[0, ["课程类型", "课程代码", "课程名称", "学分", "总学时"], 0]])
    # print(text_res)

    skjh_rules = [['2022-2023|学期授课总要求', '2022-2023|学期授课总要求\n', '(2022-2023.*?|学期授课总要求.*?)\n']]
    # skjh_rules = []
    path9 = "./授课计划/153.刘迎晓+全院选修+模特表演+22232学期授课进度计划.docx"
    result_code, text_res, res_tables_list, res_images_list, result_key_list, result_table_key_list = main(
        path9, skjh_rules, 99999, [], [])
    print(text_res)

    # tables_list =
    #
    # handle_tables(tables_list, table_rules_list)


if __name__ == '__main__':
    test()
