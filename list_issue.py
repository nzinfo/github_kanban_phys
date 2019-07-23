# -*- coding: utf-8 -*-

import sys
import os
import shutil
import collections

from github import Github
import qrcode
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm, Inches, Pt
import jinja2
from jinja2.utils import Markup

"""
    take action name as sys.argv[1]
    - export, export from github and generate to printed document
    - stage,  merge issues to the others
"""

issues_qr_code_path = "./.knowns"
export_path = "./export"

repo_name = "LambdaIM/lambdamint"
issue_url_template = "https://github.com/"+repo_name+"/issues/{0}"

GitIssue = collections.namedtuple('Person', 'n title body image')

# FIXME: 从环境变量读取
g = Github("<username>", "<password>")

repo = g.get_repo(repo_name)
open_issues = repo.get_issues(state='open')

action  = sys.argv[1]

# ensure all path
if not os.path.isdir(issues_qr_code_path):
    os.makedirs(os.path.abspath(issues_qr_code_path))

if action == "export" and os.path.isdir(export_path):
    shutil.rmtree(export_path)

if not os.path.isdir(export_path):
    os.makedirs(os.path.abspath(export_path))

if action == "stage":
    import shutil

    source = os.listdir(export_path)
    destination = issues_qr_code_path
    for f in source:
        if f.endswith(".png"):
            shutil.move(f, destination)
    exit(0) # return

# do export
def to_fix_size(s, slen, padding="..."):
    """
        如果是 中文 + 2
        其他 +1
    """ 
    rs = u''
    rs_i = 0
    s = s.replace('\r\n', '\n')
    for c in unicode(s):
        if ord(c) >= 0x4E00 and ord(c) <= 0x9FFF:
            rs_i += 2
        else:
            rs_i += 1
        if rs_i >= slen - len(padding):
            return rs + padding
        rs += c
    return rs
        
def docx_export(issues):
    tpl = DocxTemplate('issues_tpl.docx')
    issueEmpty = GitIssue('', '', '', '')

    context = {
        'issues': []
    }

    for i in range(0, len(issues), 3):
        issue1 = issues[i]
        issue2 = issueEmpty
        issue3 = issueEmpty
        if i + 1 < len(issues):
            issue2 = issues[i+1]
        if i + 2 < len(issues):
            issue3 = issues[i+2]
        
        item = {
            'n1': issue1.n,
            'title1': to_fix_size(issue1.title, 29),
            'body1': RichText(to_fix_size(issue1.body, 200), size=14),
            'image1': InlineImage(tpl, issue1.image, height=Inches(0.64)) if issue1.image else '',
            'n2': issue2.n,
            'title2': to_fix_size(issue2.title, 29),
            'body2': RichText(to_fix_size(issue2.body, 200), size=14),
            'image2': InlineImage(tpl, issue2.image, height=Inches(0.64)) if issue2.image else '',
            'n3': issue3.n,
            'title3': to_fix_size(issue3.title, 29),
            'body3': RichText(to_fix_size(issue3.body, 200), size=14),
            'image3': InlineImage(tpl, issue3.image, height=Inches(0.64)) if issue3.image else '',
        }
        context['issues'].append(item)
    
    # testing that it works also when autoescape has been forced to True
    jinja_env = jinja2.Environment(autoescape=True)
    tpl.render(context, jinja_env)
    tpl.save('inline_image.docx')

issues = []

for issue in open_issues:
    issue_number, title, body = issue.number, issue.title, issue.body
    qr_fname_stage = os.path.join(issues_qr_code_path, "%d.png" % issue_number)
    qr_fname_export = os.path.join(export_path, "%d.png" % issue_number)
    if False:
        if os.path.isfile(qr_fname_stage) or os.path.isfile(qr_fname_export):
            continue # has exported
    # 生成 qr code
    img = qrcode.make(issue_url_template.format(issue_number))
    img.save(qr_fname_export)
    # print(issue_number, body)
    issues.append(GitIssue(issue_number, title, body, qr_fname_export))

docx_export(issues)


