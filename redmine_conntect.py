

from redminelib import Redmine
import xlsxwriter


def get_member_from_redmine(redmine_url, redmine_key):
    redmine = Redmine(url=redmine_url, key=redmine_key)
    projects = redmine.project.all()
    # print(projects)
    project_list = []
    for project in projects:
        project_data = {"name": project.name}
        user_list = []
        for value in list(project.memberships.values()):

            # print(value)
            # 只取user的姓名和组名，group不取
            if "user" in value.keys():
                group = value["roles"][0]["name"]
                name = value["user"]["name"]
                _dict = {group: name}
                user_list.append(_dict)
        project_data["member_info"] = user_list
        project_list.append(project_data)
    print(project_list)


def write_excel():
    workbook = xlsxwriter.Workbook("./test.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "project_name")
    workbook.close()


url = 'http://demo.redmineup.com'
key = "9a13f31770b80767a57d753961acbd3a18eb1370"
get_member_from_redmine(url, key)
# write_excel()
