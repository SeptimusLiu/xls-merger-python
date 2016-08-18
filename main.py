# -*- coding:gbk -*-

import os
import re
import sys
reload(sys)
sys.setdefaultencoding("gbk")
import time
import xlrd
import xlwt


def main():
    file_maps = {}
    if not os.path.exists('src'):
        print('�Ҳ���Դ�ļ���')
        raw_input()
        sys.exit(1)

    start_time = time.time()
    src_cnt = 0
    res_cnt = 0

    for name in os.listdir("src"):
        if not re.search(r'.+\.xls[x]?', name):
            continue

        print('���ڼ��ر�%s...' % name)
        data = xlrd.open_workbook("src/" + name)
        sheets = data.sheet_names()
        for sh in sheets:
            table = data.sheet_by_name(sh)
            nrows = table.nrows
            if sh not in file_maps:
                file_maps[sh] = []

            for i in range(nrows):
                file_maps[sh].append(table.row_values(i))
        src_cnt += 1

    if not os.path.exists('res'):
        os.mkdir('res')

    for sh, data in file_maps.iteritems():
        print('���ںϲ���%s...' % sh.encode('gbk'))
        wb = xlwt.Workbook()
        sh_out = wb.add_sheet(sh)
        line = 0
        for i, row in enumerate(data):
            try:
                if not (row[0] == '���' and line == 0):
                    int(row[0])
                    if type(row[1]) == int or type(row[1]) == float:
                        raise
            except:
                continue
            # if row[0]
            for j, v in enumerate(row):
                sh_out.write(line, j, v)
            line += 1

        wb.save('res/%s.xls' % sh)
        res_cnt += 1
        print('��%s�ϲ����' % sh.encode('gbk'))

    end_time = time.time()
    print('�ϲ�������ɣ��ܼƴ������ļ�����%d���ϲ����ļ�����%d����ʱ��%d�롣' % (src_cnt, res_cnt, int(end_time - start_time)))
    print('��������˳�')
    raw_input()


if __name__ == '__main__':
    try:
        main()
    except Exception, err:
        print('�������г���%s' % err)
        raw_input()
