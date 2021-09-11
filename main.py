# -*- coding: UTF-8 -*-

import json
import os
import win32clipboard as w

# 默认分隔符
separator = '\t'

# 需处理的json文件路径
file_url = './1.json'

# 输出文件路径
file_output = './res.txt'


# 读入文件
def index():
    open_json = open(file_url, 'rb')
    load_json = json.loads(open_json.read())
    clear(load_json)
    open_json.close()


# 处理读入的数据，提取成一个列表(list)
def clear(jsons):
    count = 0
    for _index in jsons['result']['allSubjectList']:  # count是所有学科总数
        count += 1
    ll = []
    # jsons['result']['studentAnswerRecords'] 为一个字典(dict)
    for _index in jsons['result']['studentScoreDetailDTO']:
        l = [_index['userName'], _index['usrEduCode'], _index['allScore'], _index['classRank'],
             _index['schoolRank']]  # 姓名 学号 总分 班级排名 年级排名
        for subject in _index['scoreInfos']:  # 各科分数
            if subject['score'] == '未扫，不计排名':
                l.append('-1')
            else:
                l.append(subject['score'])
        ll.append(l)
    make_headline(jsons)
    sto(ll)


# 追加数据
def sto(ll):
    std_out = open(file_output, 'a', encoding='utf-8')
    for i in ll:
        for o in i:
            std_out.write(o + separator)
        std_out.write('\n')
    std_out.close()


# 表头
def make_headline(jsons):
    head_line = ''
    head_line = head_line + '姓名' + separator + '学号' + separator + '总分' + separator + '班级排名' + separator + '年级排名'
    for tmp in jsons['result']['allSubjectList']:
        head_line = head_line + separator + tmp['subjectName']
    # 表头处理完成，开始添加表头
    with open(file_output, 'a+', encoding='utf-8') as _f:
        _f.write(head_line + '\n')


if __name__ == '__main__':
    tmp = input(r'输入json文件名（默认为"1.json"）：')
    if tmp:
        file_url = './' + tmp
    if os.path.exists('res.txt'):  # 如果存在res.txt，则先删除，防止多次重复输出
        os.remove('res.txt')
    tmp=input('输入分割符（默认为制表符）：')
    if tmp:
        separator=tmp
    index()
    with open('res.txt', 'r', encoding='utf-8') as f:
        # 复制到剪贴板
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardText(f.read())
        w.CloseClipboard()
    print('\n已完成输出。输出内容已复制到剪贴板，在Excel中按【Ctrl+V】粘贴即可。不建议直接查看res.txt，可能产生排版错误等问题！')
    pause = input(' ')
