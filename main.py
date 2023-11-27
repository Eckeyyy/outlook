import win32com.client as win32
import pandas as pd
import pathlib, os, random


class Outook:
    def __init__(self, configs, data):
        self.configs = configs
        self.df = data

    def get_sheets(self, file):
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        return sheet_names

    def get_email_index(self):
        email_index = str(self.configs.loc['email_address_title'].dropna()[0])
        return email_index

    def get_reminder_config(self):
        data = self.configs.loc['reminder'].dropna()
        if len(data) != 0:
            str = ''
            for index, item in data.items():
                str = str + f'<div style="margin-left:0px">{item}</div>'
            return str
        else:
            return ''

    def get_body_config(self):
        data = self.configs.loc['body'].dropna()
        if len(data) == 0:
            print('请在config中编辑body内容,即邮件正文。')
        body = ''
        for index, item in data.items():
            body = body + str(f"<div>{item}</div>")
        return body

    def get_subject_config(self):
        sub_items = self.configs.loc['subject'].dropna()
        subject = ''
        # active_subject = ''
        # for sub in items.items():
        #     active_subject += str(row[sub])
        for index, item in sub_items.items():
            subject = subject + str(item)
        return subject

    def get_active_sub(self, row):
        active_sub_items = list(self.configs.loc['subject_add'].dropna())
        # print(active_sub_item)
        sub_nums = []
        for item in active_sub_items:
            sub_nums.append([str(row[item])])
        # print(sub_nums)
        return sub_nums

    def conbine_subject(self, static_subject, active_sub_content_list):
        active_sub_items = list(self.configs.loc['subject_add'].dropna())
        add_permission = str(self.configs.loc['subject_add_or_not'][0])
        subject_str = ''
        for index, title in enumerate(active_sub_items):
            subject_str += title + '#' + ';'.join(active_sub_content_list[index]) + '  '
        if add_permission == 'Y':
            return str(static_subject) + '--' + subject_str
        else:
            return static_subject

    def get_table_header_config(self):
        color_list = ["#a8bf8f", "#e0c7e3", "#bad077", "#ffa631", ]
        headers = list(self.configs.loc['table_header'].dropna())
        headers_html = ''
        for head in headers:
            random_color = random.choice(color_list)
            headers_html = headers_html + f'<th style=" box-sizing:border-box;background:{random_color};width:150px">{head}</th>'
        return headers_html

    def get_signature(self):
        signature = list(self.configs.loc['signature'].dropna())
        signature_html = ''
        for sign in signature:
            signature_html = signature_html + str(sign) + '<br/>'
        return str(signature_html)

    def get_table_content(self, row):
        headers = list(self.configs.loc['table_header'].dropna())
        if len(headers) == 0:
            print('请填写表头，需要与表格中的表头一致！')
        else:
            datas = ''
            for head in headers:
                datas = datas + f'<td>{row[head]}</td>'
            return f'<tr>{datas}</tr>'

    def get_sent_on_behalf_name(self):
        name_series = list(self.configs.loc['sent_email'].dropna())
        if not name_series:
            print('请填写发件人邮箱！')
        else:
            return str(name_series[0])

    def get_templates(self):
        templates_path = str(self.configs.loc['template'].dropna()[0])
        templates_path_list = []
        for root, directories, files in os.walk(templates_path):
            for file in files:
                # 打印文件的绝对路径
                file_path = os.path.join(root, file)
                templates_path_list.append(file_path)
        return templates_path_list

    def get_templates_permission(self):
        add_permission = str(self.configs.loc['template_sent_or_not'].dropna()[0])
        return add_permission

    def get_sent_permission(self):
        sent_permission = str(self.configs.loc['sent_or_not'].dropna()[0])
        return sent_permission

    def get_files_permission(self):
        add_permission = str(self.configs.loc['files_sent_or_not'].dropna()[0])
        # print(add_permission)
        return add_permission

    def get_file_name_list(self):
        file_name_list = list(self.configs.loc['files_name'].dropna())
        if len(file_name_list) != 0:
            return file_name_list
        else:
            pass

    def get_files(self, file_name):
        file_path = str(configs.loc['files_path'].dropna()[0])
        for root, directories, files in os.walk(file_path):
            for file in files:
                if str(file_name) == file.split('.')[0]:
                    # print(os.path.join(root, file))
                    return os.path.join(root, file)
                else:
                    pass

    def send_email(self, subject='', body='', to_address='', temlate_list=[], tem_permission='', add_path_list=[],
                   add_permission='', sent_permission=''):
        image_path = str(pathlib.Path('image002.gif').absolute())
        image_path2 = str(pathlib.Path('image003.jpg').absolute())
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.htmlBody = body
        mail.To = to_address
        sent_email = self.get_sent_on_behalf_name()
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

    def get_sent_with_table_permission(self):
        table_permission = list(configs.loc['sent_with_table'].dropna())[0]
        return str(table_permission)

    def get_html_body(self, msg, table_head, email_content, reminder, signature):
        table_permission = self.get_sent_with_table_permission()
        if table_permission == 'Y':
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
        else:
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
                                            <a herf="www.baidu.com" >百度</a>
                                              <div style='margin-left:0px;margin-top:10px'>
                                                 {msg}
                                              </div>
                                             <br/>
                                             <div style='margin-top:0px'>
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

    def send(self):
        try:
            reminder = self.get_reminder_config()
            msg = self.get_body_config()
            signature = self.get_signature()
            table_head = self.get_table_header_config()
            email_content_dict = {}
            template_list = self.get_templates()
            tem_permission = self.get_templates_permission()
            add_permission = self.get_files_permission()
            sent_permission = self.get_sent_permission()
            # print(sent_permission)
            file_name_list = self.get_file_name_list()
            subject = self.get_subject_config()
            emai_addr = self.get_email_index()
            for index, row in self.df.iterrows():
                email_address = row[emai_addr].replace('；', ';')
                add_path_list = []
                # po_no = row["po_no"]
                for item in file_name_list:
                    find_index = row[item]
                    add_file_path = self.get_files(find_index)
                    if add_file_path != None:
                        add_path_list.append(add_file_path)
                # print(add_path_list)
                acitve_subject_list = self.get_active_sub(row)
                # print(acitve_subject_list)
                table_content = self.get_table_content(row)
                if email_address in email_content_dict:
                    # Append the content for the same email address
                    email_content_dict[email_address][0] += self.get_table_content(row)
                    email_content_dict[email_address][1] += add_path_list
                    email_content_dict[email_address][2][0] += acitve_subject_list[0]
                    email_content_dict[email_address][2][1] += acitve_subject_list[1]
                else:
                    email_content_dict[email_address] = [table_content, add_path_list, acitve_subject_list]
            # print(email_content_dict['1055159845@qq.com'][2])
            for index, (email_address, email_content) in enumerate(email_content_dict.items()):
                html_body = self.get_html_body(msg, table_head, email_content[0], reminder, signature)
                final_subject = self.conbine_subject(subject, email_content[2])
                print(f'正在发送第{index + 1}封邮件，请稍等...')
                self.send_email(final_subject, html_body, email_address, template_list, tem_permission,
                                email_content[1],
                                add_permission, sent_permission)
            input("邮件发送||生成完毕,点击回车退出========>")
        except Exception as e:
            print('程序出错了>>_<<!错误代码====》', e)


if __name__ == "__main__":

    config__sheet_name = input('请输入你的自定义邮件主题sheet名===>回车继续 o^^o如果不输入则使用默认配置config:')
    config_name = 'config'
    if config__sheet_name.strip():
        config_name = config__sheet_name.strip()
    else:
        pass
    excel_file = "email.xlsx"
    # author=====>Ecke EXW009 ##############################################
    print('powered by python library--pywin32,pandas,os,pathlib,random')
    df = pd.read_excel(excel_file, sheet_name='data')
    configs = pd.read_excel(excel_file, sheet_name=config_name, index_col=0)
    outlook = Outook(configs, df)
    outlook.send()
