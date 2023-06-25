# -*- coding: utf-8 -*-
###########################################################################
#    Copyright 2023 奇客罗方智能科技 https://www.geekercloud.com
#    ITGeeker.net <alanljj@gmail.com>
############################################################################
import glob
import os
import win32com.client


def separate_file_info_by_ffp(ffp):
    dirname = os.path.dirname(ffp)
    filename, file_extension = os.path.splitext(ffp)
    # basename = os.path.basename(ffp)
    basename_no_ext = os.path.splitext(os.path.basename(ffp))[0]
    print(basename_no_ext)
    print('dirname: ', dirname)
    # print('filename: ', filename)
    print('basename_no_ext: ', basename_no_ext)
    # print('basename: ', basename)
    print('file_extension: ', file_extension)
    return dirname, basename_no_ext, file_extension


def convert_doc2docx_by_win32com(val_list):
    # print('val_list: %s' % val_list)
    doc_list = []
    for val in val_list:
        ffp = os.path.join(val[1], val[0])
        # print('ffp: %s' % ffp)
        doc_list.append(ffp)

    word = win32com.client.Dispatch("Word.Application")
    word.visible = 0
    doc2docx_l = []
    doc_err_l = []

    for i, doc in enumerate(doc_list):
        in_file = os.path.abspath(doc)
        # print('in_file: %s' % in_file)
        dirname, basename_no_ext, file_extension = separate_file_info_by_ffp(in_file)
        new_docx_f = os.path.join(dirname, basename_no_ext + '-converted' + '.docx')
        # print('new_docx_f: %s' % new_docx_f)
        try:
            wb = word.Documents.Open(in_file)
            wb.SaveAs2(new_docx_f, FileFormat=16)  # file format for docx
            wb.Close()
            doc2docx_l.append(new_docx_f)
            print('/*-/*-/*-/*-/*-/*-/*-/*-len(doc2docx_l): ', len(doc2docx_l))
        except Exception as err:
            print('err@try wb = word.Documents.Open(in_file): ', err)
            doc_err_l.append(in_file)
    word.Quit()

    print('doc2docx_l: %s' % doc2docx_l)
    print('doc_err_l: %s' % doc_err_l)
    return doc2docx_l, doc_err_l
