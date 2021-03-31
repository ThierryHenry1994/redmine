from redminelib import Redmine
import xlsxwriter


def handle_user_info(user_info):
    result = []
    for users in user_info:
        user_variety = users[0]
        user_name = users[1]
        old_variety = ""
        for s in result:
            if user_variety == s[0]:
                s[1].append(user_name)
                old_variety = True
            else:
                pass
        if not old_variety:
            result.append([user_variety, [user_name]])

    return result


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
                user_list.append((group, name))
                # print(user_list)
        user_list = handle_user_info(user_list)
        project_data["member_info"] = user_list
        project_list.append(project_data)
    return project_list


def write_excel(excel_list):
    for sub in range(len(excel_list)):

        excel_name = "./"+excel_list[sub]["name"]+".xlsx"
        workbook = xlsxwriter.Workbook(excel_name)
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "project_name")
        worksheet.write(0, 1, "member")
        count = 0
        for info in excel_list[sub]["member_info"]:

            print("======", info)
            worksheet.write(count+1, 0, info[0])
            worksheet.write(count + 1, 1, ','.join(info[1]))
            count +=1
        workbook.close()


url = 'http://demo.redmineup.com'
key = "9a13f31770b80767a57d753961acbd3a18eb1370"
p_list = get_member_from_redmine(url, key)
write_excel(p_list)
