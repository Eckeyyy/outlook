import win32com.client as win32
import pandas as pd
import pathlib, os, random


def send_email(subject='', body='', to_address='', temlate_list=[], tem_permission='', add_path_list=[],
               add_permission='', sent_permission=''):
    image_path = str(pathlib.Path('image002.gif').absolute())
    image_path2 = str(pathlib.Path('image003.jpg').absolute())
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.htmlBody = body
    mail.To = to_address
    sent_email = get_sent_on_behalf_name(configs)
    if sent_email != None:
        mail.SentOnBehalfOfName = sent_email
    else:
        pass
    att = mail.Attachments.Add(image_path)
    att1 = mail.Attachments.Add(image_path2)
    att.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "img")
    att1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "img2")
    if len(temlate_list) != 0 and tem_permission == 'Y':
        for tem in temlate_list:
            mail.Attachments.Add(tem)
    else:
        pass
    if len(add_path_list) != 0 and add_permission == 'Y':
        for path in add_path_list:
            mail.Attachments.Add(path)
    else:
        pass
    if sent_permission == 'Y':
        mail.Send()
    else:
        mail.Display()


def get_reminder_config(configs):
    data = configs.loc['reminder'].dropna()
    str = ''
    for index, item in data.items():
        str = str + f'<div style="margin-left:0px">{item}</div>'
    return str


def get_body_config(configs):
    data = configs.loc['body'].dropna()
    body = ''
    for index, item in data.items():
        body = body + str(f"<div>{item}</div>")
    return body


def get_subject_config(configs):
    data = configs.loc['subject'].dropna()
    # items = configs.loc['subject_add'].dropna()
    subject = ''
    # active_subject = ''
    # for sub in items.items():
    #     active_subject += str(row[sub])
    for index, item in data.items():
        subject = subject + str(item)
    # print(active_subject)
    return subject


def get_table_header_config(configs):
    color_list = ["#a8bf8f", "#e0c7e3", "#bad077", "#ffa631", ]
    headers = list(configs.loc['table_header'].dropna())
    headers_html = ''
    for head in headers:
        random_color = random.choice(color_list)
        headers_html = headers_html + f'<th style=" box-sizing:border-box;background:{random_color};width:150px">{head}</th>'
    return headers_html


def get_signature(configs):
    signature = list(configs.loc['signature'].dropna())
    signature_html = ''
    for sign in signature:
        signature_html = signature_html + str(sign) + '<br/>'
    return str(signature_html)


def get_table_content(configs, row):
    headers = list(configs.loc['table_header'].dropna())
    datas = ''
    for head in headers:
        datas = datas + f'<td>{row[head]}</td>'
    return f'<tr>{datas}</tr>'


def get_sent_on_behalf_name(congfigs):
    name = congfigs.loc['sent_email'].dropna()[0]
    return str(name)


def get_html_body(msg, table_head, email_content, reminder, signature):
    table_head_class = """
    table thead tr th { 
                        box-sizing:border-box;
                        font-size:16px;
                        font-weight:400;
                        padding:0 5px;
                        height:30px;
                        }
    table tr td{
                padding:0 5px;
                }
    div {
            font-family: "Microsoft YaHei", Arial, sans-serif;
                 }
            """
    html_body = f'''
                <style>
                    {table_head_class}
                </style>
                   <div style="width:100%;height:100%；padding:0px 40px">
                        Dear Customer,
                             <div style='margin-left:0px;margin-top:10px'>
                                {msg}
                             </div>
                            <br/>
                            <br/>
                             <table border style='border-collapse: collapse;text-align:center;margin-left:20px'>
                                <thead>
                                    <tr>
                                        {table_head}
                                    </tr>
                                        {email_content}
                                </thead>
                            </table>
                            <br/>
                            <div style='margin-top:50px'>
                                {reminder}
                            </div>
                            <div style='margin-top:80px'>
                              {signature}
                                <div>
                                    <img src="cid:img" alt="Image" >
                                </div>
                                <div>
                                    <img src="cid:img2" alt="Image" >
                                </div>
                            </div>
                        </div>
                '''
    return html_body


def get_templates(configs):
    templates_path = str(configs.loc['template'].dropna()[0])
    templates_path_list = []
    for root, directories, files in os.walk(templates_path):
        for file in files:
            # 打印文件的绝对路径
            file_path = os.path.join(root, file)
            templates_path_list.append(file_path)
    return templates_path_list


def get_templates_permission(configs):
    add_permission = str(configs.loc['template_sent_or_not'].dropna()[0])
    return add_permission


def get_sent_permission(config):
    sent_permission = str(configs.loc['sent_or_not'].dropna()[0])
    return sent_permission


def get_files_permission(configs):
    add_permission = str(configs.loc['files_sent_or_not'].dropna()[0])
    # print(add_permission)
    return add_permission


def get_file_name_list(configs):
    file_name_list = list(configs.loc['files_name'].dropna())
    if len(file_name_list) != 0:
        return file_name_list
    else:
        pass


def get_files(file_name):
    file_path = str(configs.loc['files_path'].dropna()[0])
    for root, directories, files in os.walk(file_path):
        for file in files:
            if str(file_name) == file.split('.')[0]:
                # print(os.path.join(root, file))
                return os.path.join(root, file)
            else:
                pass


if __name__ == "__main__":
    try:
        excel_file = "email.xlsx"
        df = pd.read_excel(excel_file, sheet_name='data')
        configs = pd.read_excel(excel_file, sheet_name='config', index_col=0)
        reminder = get_reminder_config(configs)
        msg = get_body_config(configs)
        signature = get_signature(configs)
        attachment_folder = 'files'
        table_head = get_table_header_config(configs)
        email_content_dict = {}
        get_templates(configs)
        template_list = get_templates(configs)
        tem_permission = get_templates_permission(configs)
        add_permission = get_files_permission(configs)
        sent_permission = get_sent_permission(configs)
        # print(sent_permission)
        file_name_list = get_file_name_list(configs)
        for index, row in df.iterrows():
            email_address = row["email_address"].replace('；', ';')
            add_path_list = []
            # po_no = row["po_no"]
            for item in file_name_list:
                find_index = row[item]
                add_file_path = get_files(find_index)
                if add_file_path != None:
                    add_path_list.append(add_file_path)
            # print(add_path_list)
            table_content = get_table_content(configs, row)
            if email_address in email_content_dict:
                # Append the content for the same email address
                email_content_dict[email_address][0] += get_table_content(configs, row)
                email_content_dict[email_address][1] += add_path_list
            else:
                email_content_dict[email_address] = [table_content, add_path_list]
            subject = get_subject_config(configs)
        # print(email_content_dict)
        for index, (email_address, email_content) in enumerate(email_content_dict.items()):
            html_body = get_html_body(msg, table_head, email_content[0], reminder, signature)
            print(f'正在发送第{index + 1}封邮件，请稍等...')
            send_email(subject, html_body, email_address, template_list, tem_permission, email_content[1],
                       add_permission, sent_permission)
    except Exception as e:
        print('程序出错了>>_<<!错误代码====》',e)

input("邮件发送||生成完毕,点击回车退出========>")
