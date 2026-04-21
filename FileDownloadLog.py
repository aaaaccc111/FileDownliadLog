import os
import requests
from msal import ConfidentialClientApplication
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pyodbc
import datetime
import pytz
from dotenv import load_dotenv

load_dotenv()

supervisor_cache = {}

server = os.getenv("DB_SERVER")
database = os.getenv("DB_DATABASE")
username = os.getenv("DB_USERNAME")
password = os.getenv("DB_PASSWORD")
tenant_id = os.getenv("TENANT_ID")
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
sender_email = os.getenv("SENDER_EMAIL")
sender_password = os.getenv("SENDER_PASSWORD")


def get_access_token(client_id, client_secret, tenant_id):
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://manage.office.com/.default"]
    app = ConfidentialClientApplication(client_id=client_id, client_credential=client_secret, authority=authority)
    result = app.acquire_token_silent(scopes=scopes, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=scopes)
    if "access_token" in result:
        return result["access_token"]
    raise Exception(f'Token 取得失敗: {result.get("error_description")}')


def send_api_request(url, access_token):
    headers = {'Authorization': 'Bearer ' + access_token}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()
    raise Exception(f'API Request 失敗: {response.status_code}')


def send_email(sender_email, sender_password, recipient_email_list, subject, content):
    if not recipient_email_list: return
    message = MIMEMultipart()
    message["From"] = sender_email
    message['To'] = ', '.join(recipient_email_list)
    message["Subject"] = subject
    message.attach(MIMEText(content, "html"))
    with smtplib.SMTP("smtp.office365.com", 587) as smtp_server:
        smtp_server.starttls()
        smtp_server.login(sender_email, sender_password)
        smtp_server.sendmail(sender_email, recipient_email_list, message.as_string())


def connect_to_sql_server():
    conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    return pyodbc.connect(conn_str)


def execute_query(conn, query, params=None):
    cursor = conn.cursor()
    cursor.execute(query, params) if params else cursor.execute(query)
    rows = cursor.fetchall()
    cursor.close()
    return rows


def adjust_timezone(creation_time):
    dt = datetime.datetime.strptime(creation_time, "%Y-%m-%dT%H:%M:%S")
    return pytz.timezone("UTC").localize(dt).astimezone(pytz.timezone("Asia/Taipei"))



# 查詢部門資料
def get_employee_info(email):
    conn = connect_to_sql_server()
    query = """
    SELECT TOP 1 EmployeeID, Position, SupervisorID, EmployeeName, EmployeeEmail
    FROM Employees WHERE LOWER(EmployeeEmail) = ? AND LeaveDate IS NULL
    ORDER BY EmployeeID DESC
    """
    res = execute_query(conn, query, (email.lower().strip(),))
    conn.close()
    if res:
        return {"EmployeeID": res[0][0], "Position": res[0][1], "SupervisorID": res[0][2], "EmployeeName": res[0][3], "EmployeeEmail": res[0][4]}
    return None


def get_supervisor_info(s_id):
    if s_id in supervisor_cache: return supervisor_cache[s_id]
    conn = connect_to_sql_server()
    query = "SELECT EmployeeEmail, Position, EmployeeName FROM Employees WHERE EmployeeID = ? AND LeaveDate IS NULL"
    res = execute_query(conn, query, (s_id,))
    conn.close()
    if res:
        info = {"EmployeeEmail": res[0][0], "Position": res[0][1], "EmployeeName": res[0][2]}
        supervisor_cache[s_id] = info
        return info
    return None

def main():
    print("start")
    access_token = get_access_token(client_id, client_secret, tenant_id)
    today = datetime.datetime.now().date()
    current_year = today.year

    personal_records = {}   # 用於個人信
    manager_summaries = {}  # 用於主管信


    for i in range(1, 2): ##0是抓今天，1是抓昨天，以此類推 例如2,3抓後天 1,3是抓昨天跟前天
        target_date = today - datetime.timedelta(days=i)
        start_time, end_time = f"{target_date}T00:00:00", f"{target_date}T23:59:59"
        print(f"正在抓取 {target_date} 所有人員的下載檔案紀錄...")

        endpoint = f'https://manage.office.com/api/v1.0/{tenant_id}/activity/feed/subscriptions/content?contentType=Audit.SharePoint&startTime={start_time}&endTime={end_time}'

        try:
            content_list = send_api_request(endpoint, access_token)
            for chunk in content_list:
                details = send_api_request(chunk['contentUri'], access_token)
                for r in details:
                    if r.get('Operation') == 'FileDownloaded':
                        u_email = r.get('UserId', '').lower().strip()
                        staff = get_employee_info(u_email)

                        if staff:
                            adj_time = adjust_timezone(r['CreationTime'])
                            file_info = (adj_time, r['ObjectId'])

                            personal_records.setdefault(u_email, []).append(file_info)

                            m_id = staff['SupervisorID']
                            if m_id:
                                m_info = get_supervisor_info(m_id)
                                if m_info:
                                    m_pos = m_info['Position'] or ""
                                    m_email = m_info['EmployeeEmail'].lower().strip()

                                    if u_email == m_email:
                                        continue

                                    #如果下載原是經理或處長 他上級不收主管信
                                    staff_pos = staff.get('Position') or ""
                                    if "經理" in staff_pos or "處長" in staff_pos:
                                        continue


                                    high_level_keys = ["處長", "COO", "副總", "總經理", "CEO"]
                                    is_high_level = any(key in m_pos for key in high_level_keys)

                                    #判斷是否為主管職位，如果是主管職位就不發主管信給他，因為他自己就會收到個人信了
                                    if is_high_level:
                                        pass

                                    # 建立清單
                                    if m_id not in manager_summaries:
                                        manager_summaries[m_id] = {}
                                    manager_summaries[m_id].setdefault(u_email, []).append(file_info)
        except Exception as e:
            print(f"抓取錯誤: {e}")

    print("\n--- 發送信件 ---")

    # 個人信
    for email, files in personal_records.items():
        emp_rows_html = ""
        for timestamp, file_path in sorted(files, key=lambda x: x[0], reverse=True):
            formatted_time = timestamp.strftime("%Y-%m-%d %H:%M:%S")
            emp_rows_html += f"""
                <tr>
                    <td style="padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666; font-size: 13px; white-space: nowrap;">{formatted_time}</td>
                    <td style="padding: 12px; border-bottom: 1px solid #eeeeee; color: #333333; font-size: 13px; word-break: break-all;">{file_path}</td>
                </tr>
            """

        employee_content = f"""
        <div style="background-color: #f6f8fa; padding: 20px; font-family: 'Microsoft JhengHei', Arial, sans-serif;">
            <table align="center" border="0" cellpadding="0" cellspacing="0" width="650" style="background-color: #ffffff; border-collapse: collapse; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                <tr><td style="background-color: #2c3e50; padding: 20px; color: #ffffff; text-align: center;"><h1 style="margin: 0; font-size: 18px; letter-spacing: 2px;">SharePoint 檔案下載記錄</h1></td></tr>
                <tr>
                    <td style="padding: 30px 20px;">
                        <p style="font-size: 15px; color: #333333;">您好，系統紀錄到您上週有自 SharePoint 下載檔案的紀錄，請確認：</p>
                        <table width="100%" style="border: 1px solid #eeeeee; border-collapse: collapse;">
                            <thead>
                                <tr style="background-color: #f9f9f9;"><th style="padding: 12px; border-bottom: 2px solid #2c3e50; text-align: left; color: #2c3e50; font-size: 14px;">下載時間</th><th style="padding: 12px; border-bottom: 2px solid #2c3e50; text-align: left; color: #2c3e50; font-size: 14px;">檔案路徑 / 內容</th></tr></thead>
                            <tbody>{emp_rows_html}</tbody>
                        </table>
                        <div style="margin-top: 30px; padding: 15px; background-color: #fff9f4; border-left: 4px solid #d35400; border-radius: 4px;">
                            <p style="margin: 0; font-size: 13px; color: #d35400; text-align: center;">提醒您，檔案下載僅限公司內部使用，嚴禁外流及儲存於外部裝置。 </p>
                        </div>
                    </td>
                </tr>
                <tr><td style="padding: 20px; background-color: #fafafa; border-top: 1px solid #eeeeee; text-align: center; color: #999999; font-size: 12px;">此為系統自動發送郵件，請勿直接回覆。<br>© {current_year} Hyena Inc. Information Department</td></tr>
            </table>
        </div>
        """
        print(f"寄送個人信: {email}")
        send_email(sender_email, sender_password, [email], "【系統通知】SharePoint 檔案下載紀錄", employee_content)

    # 主管信
    for m_id, staff_dict in manager_summaries.items():
        m_info = get_supervisor_info(m_id)
        all_manager_rows_html = ""

        #計數器
        staff_count = 0

        for s_email, files in sorted(staff_dict.items()):
            staff_name = s_email.split('@')[0]

            row_bg_color = "#ffffff" if staff_count % 2 == 0 else "#f8f9fa"

            sorted_files = sorted(files, key=lambda x: x[0], reverse=True)

            # 取得該員工最後一筆紀錄的索引，用來判斷是否要畫加粗底線
            num_files = len(sorted_files)

            for index, (timestamp, file_path) in enumerate(sorted_files):
                formatted_time = timestamp.strftime("%Y-%m-%d %H:%M:%S")

                # 3. 每個人最後一筆 底線外框加粗
                bottom_style = "2px solid #adb5bd" if index == num_files - 1 else "1px solid #eeeeee"


                all_manager_rows_html += f"""
                    <tr style="background-color: {row_bg_color};">
                        <td style="padding: 12px; border-bottom: {bottom_style}; background-color: {row_bg_color}; color: #333333; font-size: 13px; font-weight: bold;">{staff_name}</td>
                        <td style="padding: 12px; border-bottom: {bottom_style}; background-color: {row_bg_color}; color: #666666; font-size: 13px; white-space: nowrap;">{formatted_time}</td>
                        <td style="padding: 12px; border-bottom: {bottom_style}; background-color: {row_bg_color}; color: #333333; font-size: 13px; word-break: break-all;">{file_path}</td>
                    </tr>
                """


            staff_count += 1

        manager_content = f"""
        <div style="background-color: #f6f8fa; padding: 20px; font-family: 'Microsoft JhengHei', Arial, sans-serif;">
            <table align="center" border="0" cellpadding="0" cellspacing="0" width="800" style="background-color: #ffffff; border-collapse: collapse; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                <tr><td style="background-color: #2c3e50; padding: 20px; color: #ffffff; text-align: center;"><h1 style="margin: 0; font-size: 18px; letter-spacing: 2px;">所屬員工 SharePoint 檔案下載記錄彙整</h1></td></tr>
                <tr>
                    <td style="padding: 30px 20px;">
                        <p style="font-size: 15px; color: #333333;">主管您好，以下為部門員工上週自 SharePoint 下載檔案紀錄，請確認：</p>
                        <table width="100%" style="border-collapse: collapse; border: 1px solid #eeeeee;">
                            <thead>
                                <tr style="background-color: #eeeeee;">
                                    <th style="padding: 12px; border-bottom: 2px solid #2c3e50; text-align: left; color: #2c3e50; font-size: 14px; width: 15%;">員工</th>
                                    <th style="padding: 12px; border-bottom: 2px solid #2c3e50; text-align: left; color: #2c3e50; font-size: 14px; width: 25%;">下載時間</th>
                                    <th style="padding: 12px; border-bottom: 2px solid #2c3e50; text-align: left; color: #2c3e50; font-size: 14px; width: 60%;">檔案路徑 / 內容</th>
                                </tr>
                            </thead>
                            <tbody>{all_manager_rows_html}</tbody>
                        </table>
                        <div style="margin-top: 30px; padding: 15px; background-color: #fff9f4; border-left: 4px solid #d35400; border-radius: 4px;">
                            <p style="margin: 0; font-size: 13px; color: #d35400; text-align: center;">請主管留意員工檔案下載情形，如有異常下載，請通知相關單位做處理，以保障公司資訊安全。 </p>
                        </div>
                    </td>
                </tr>
                <tr><td style="padding: 20px; background-color: #fafafa; border-top: 1px solid #eeeeee; text-align: center; color: #999999; font-size: 12px;">此為系統自動發送郵件，請勿直接回覆。<br>© {current_year} Hyena Inc. Information Department</td></tr>
            </table>
        </div>
        """
        print(f"寄送主管信: {m_info['EmployeeEmail']}")
        send_email(sender_email, sender_password, [m_info['EmployeeEmail']], "【系統通知】所屬員工 SharePoint 檔案下載紀錄", manager_content)

    print("\nend")

if __name__ == '__main__':
    main()