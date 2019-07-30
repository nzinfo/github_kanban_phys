# -*- coding: utf-8 -*-
"""
    从 Excel 文件导入 Issue

    # Issues

      取 第一个 Sheet 作为模板， 包括

      分组， 编号， 标题， 里程碑， 模块/项目(列表)，描述

    # 导入逻辑

        1. 读取全部的 milestone，检测当前项目是否存在
            
            - 如不存在，创建之

        2. 读取全部的项目， 检测当前项目是否存在
            
            - 如不存在，创建之

        3. 读取全部的 issue，解析现存的 issue 的编号

            - 如存在，略过
            - 如不存在，创建之
    
    # 创建

        全部创建的工作进入队列，统一执行队列
"""
import sys
import re

import xlrd
from github import Github, GithubObject


class GithubExecutor(object):

    def __init__(self, repo_name, user_name, passwd):
        g = Github(user_name, passwd)
        
        repo = g.get_repo(repo_name)
        
        self.repo = repo
        
        # milestones
        self.milestones = {}
        milestones = repo.get_milestones()
        for m in milestones:
            self.milestones[m.title] = m

        # projects
        self.projects = {}
        for p in repo.get_projects():
            self.projects[p.name] = p
        
        # issues
        existed_issues = repo.get_issues(state='all')
        
        self.issues = {}
        for issue in existed_issues:
            _, title, _ = issue.number, issue.title, issue.body
            m = re.search(r'\[(.*)\] (.*)', title)
            if m:
                k = m.groups()[0]
                self.issues[k] = issue
            # parse

    def ensureMilestone(self, milestone_name):
        if milestone_name not in self.milestones:
            return self.repo.create_milestone(title=milestone_name)
        else:
            return self.milestones[milestone_name]

    
    def ensureProject(self, project_name):
        if project_name not in self.projects:
            return self.repo.create_project(project_name, body="")
        else:
            return self.projects[project_name]

    def exist(self, issue):
        m = re.search(r'\[(.*)\] (.*)', issue.getTitle())
        k = ''
        if m:
            k = m.groups()[0]
        return k in self.issues
    
    def newIssue(self, issue):
        self.repo.create_issue(
            title= issue.getTitle(),
            body= u"\n".join(issue.description),
            milestone=issue.milestone
        )

class Issue(object):
    def __init__(self):
        self.category = ""
        self.issue_no = 1
        self.title = ""
        self.milestone = ""
        self.projects = []
        self.description = []   # array of lines
    
    def getTitle(self):
        return ('[%s-%04d] %s' % (self.category, int(self.issue_no), self.title))

class RowFeeder(object):

    def __init__(self):
        self.head = None

        # build a issue
        self.issue = Issue()
    
    def feed(self, row):
        """
            当构成了一个 task 时， 返回
        """
        if self.head is None:
            self.head = row     # just skip 1st line
            return None, False

        category, sno, title, milestone, project, desc = row[0:6]
        category = category.value.strip()
        sno = str(sno.value).strip()
        project = str(project.value).strip()

        res_issue = None
        if self.issue.category and self.issue.issue_no and category and sno and (self.issue.category != category or self.issue.issue_no != sno):
            """
                全部的字段均有数据， 含义是
                
                - 之前已经存在一个 Issue
                - 下一行 是 新 Issue 的第一行
                - 下一行 与 上一个 Issue 不是同一个 （ 如果 category，issue_no 相同 视为补充）
            """
            res_issue = self.issue
            self.issue = Issue()

        # 以下属性只能设置一次
        if not self.issue.category:
            self.issue.category = category
        if not self.issue.issue_no:
            self.issue.issue_no = sno
        if not self.issue.title:
            self.issue.title = title.value
        if not self.issue.milestone:
            self.issue.milestone = milestone.value

        if project and project not in self.issue.projects :
            self.issue.projects.append(project)

        self.issue.description.append(desc.value)

        return res_issue, res_issue is not None

    def lastIssue(self):
        if self.issue.issue_no:
            return self.issue
        else:
            return None

def main():

    repo_name = sys.argv[1]
    f_name = sys.argv[2]

    if False:
        s = "[UI-0012 xxkdjflsdjf"
        r = re.search(r'\[(.*)\] (.*)', s)
        print(r.groups()[0])
        
    executor = GithubExecutor(repo_name, 'nzinfo', '^Coreseek2010$')
    workbook = xlrd.open_workbook(f_name)
    
    # 根据sheet索引或者名称获取sheet内容
    sheet = workbook.sheet_by_index(0) # sheet索引从0开始
    
    tasks = []
    data_feeder = RowFeeder()
    for row in sheet.get_rows():
        task, has_task = data_feeder.feed(row)
        if has_task:
            tasks.append(task)
    # deal last
    if data_feeder.lastIssue():
        tasks.append(data_feeder.lastIssue())
    
    # update the tasks    
    for task in tasks:
        if task.milestone:
            task.milestone = executor.ensureMilestone(task.milestone)
        else:
            task.milestone = GithubObject.NotSet

        for p in task.projects:
            executor.ensureProject(p)
        # print task.getTitle()
        executor.newIssue(task)



if __name__ == '__main__':

    main()

# end of file
