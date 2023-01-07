
from time import sleep
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText

from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import pandas
from webdriver_manager.chrome import ChromeDriverManager

RESERVE_FILE_EXT = "_reserve"


def save_table(card_number: str, password: str, url: str) -> str:
    # headless mode
    option = Options()
    option.add_argument('--headless')
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=option)

    # normal mode
    # driver = webdriver.Chrome(Service(ChromeDriverManager().install()))

    driver.get(url)
    user_card_no = driver.find_element(by=By.NAME, value="usercardno")
    user_card_no.send_keys(card_number)
    user_password = driver.find_element(by=By.NAME, value="userpasswd")
    user_password.send_keys(password)

    form = driver.find_element(by=By.NAME, value="InForm")
    form.submit()

    sleep(5)
    handle_array = driver.window_handles
    driver.switch_to.window(handle_array[-1])
    is_successful = extend_rentals(driver)
    zip_name = f'{card_number}.zip'
    dfs = pandas.read_html(driver.page_source)
    df = dfs[4]
    df.to_pickle(zip_name)

    # reserve df
    zip_name2 = f"{zip_name.split('.')[0]}{RESERVE_FILE_EXT}.zip"
    is_reserve = (str(dfs[5]).find("シリーズ予約組替") != -1)
    if is_successful & is_reserve:
        reserve_index = 6
    else:
        reserve_index = 5
    df2 = pandas.read_html(driver.page_source)[reserve_index]
    df2.to_pickle(zip_name2)

    return is_successful


def extend_rentals(driver: WebDriver) -> None:
    date_today = str(datetime.today().date())

    is_loop = True
    is_successful = False
    while is_loop:
        df = pandas.read_html(driver.page_source)[4]
        if all(isinstance(_col, int) for _col in df.columns):
            break
        df.columns = [col.split()[0] for col in df.columns]
        df = df.loc[:, ['貸出更新', 'タイトル', '返却期限日']]
        df = df.dropna(how='any')
        df['返却期限日'] = pandas.to_datetime(df['返却期限日'])
        are_deadline = [str(_date.date()) ==
                        date_today for _date in df['返却期限日'].to_list()]
        are_updatable = [_update ==
                         "再貸出" for _update in df['貸出更新'].to_list()]
        are_update_now = [_d & _u for _d, _u in zip(
            are_deadline, are_updatable) if _u]

        buttons = driver.find_elements(
            by=By.XPATH, value="//button[@value='再貸出']")

        is_successful = True
        if any(are_update_now):
            buttons[are_update_now.index(True)].click()
            pass
            sleep(2)

            extend_button = driver.find_element(
                by=By.NAME, value="chkLKOUSIN")
            extend_button.click()
            pass
            sleep(2)
        else:
            is_loop = False
    return is_successful


# def get_reserve_df(driver: WebDriver):
#     reserve_button = driver.find_element(
#         by=By.XPATH, value="//a[@href='#ContentRsv']")
#     reserve_button.click()
#     table = driver.find_element(by=By.XPATH,
#                                 value="//table[tbody[tr[th[span[contains(text(), 'No.']]]]]")

#     pass


def refine_table(file_name: str) -> pandas.DataFrame:
    df = pandas.read_pickle(file_name)
    df.columns = [col.split()[0] for col in df.columns]
    # df2 = df.loc[:, ['貸出更新', 'タイトル', '貸出日', '返却期限日']]
    df = df.loc[:, ['貸出更新', 'タイトル', '返却期限日']]
    df['返却期限日'] = pandas.to_datetime(df['返却期限日'])
    df['ID'] = [file_name.replace('.zip', '')] * len(df)
    # df2 = df2.reindex(columns=['タイトル', '返却期限日''ID', ])
    return df.dropna(how='any')


def refine_table2(file_name: str) -> pandas.DataFrame:
    df = pandas.read_pickle(file_name)
    if (str(df).find("予約データはありません") == -1) & (str(df).find("予約取消データはありません。") == -1):
        df = df.loc[:, ["タイトル", "状況", "取り置き期限"]]
        df['ID'] = [file_name.split('_')[0]] * len(df)
    else:
        df = pandas.DataFrame()
    return df


def generate_earliest_date(zips: list, res_extend: dict) -> str:
    df = pandas.DataFrame()
    for zip in zips:
        if res_extend[zip]:
            df = pandas.concat([df, refine_table(f"{zip}.zip")])
    return min(df['返却期限日'])


def availablility_reserve(zips: list) -> bool:
    df = pandas.DataFrame()
    for zip in zips:
        df = pandas.concat([df, refine_table2(f"{zip}{RESERVE_FILE_EXT}.zip")])
    if not df.empty:
        return any(True for _ in df.loc[:, "状況"].to_list() if "準備できました" in _)
    else:
        return False


def create_email_message(zips: list, res_extend: dict) -> str:
    df = pandas.DataFrame()
    df2 = pandas.DataFrame()
    for zip in zips:
        if res_extend[zip]:
            df = pandas.concat([df, refine_table(f"{zip}.zip")])

        df2 = pandas.concat(
            [df2, refine_table2(f"{zip}{RESERVE_FILE_EXT}.zip")])

    df = df.sort_values(by='タイトル')
    df = df.sort_values(by='返却期限日')
    if '取り置き期限' in df2.columns:
        df2 = df2.sort_values(by='取り置き期限')
    are_extentable = [_update != "再貸出" for _update in df['貸出更新'].to_list()]
    df = df.loc[:, ['タイトル', '返却期限日', 'ID']]

    message = df.to_html(index=False)
    lines = []
    current_year = str(datetime.today().date())[0:4]
    for _line in message.split('\n'):
        if f"<td>{current_year}" in _line:
            if are_extentable.pop(0):
                _line = _line.replace(f'<td>{current_year}',
                                      f'<td style="text-decoration: underline;">{current_year}')
        lines.append(_line)

    message = "\n".join(lines)
    message = message.replace(
        '<table border="1" class="dataframe">',
        '<table border="1" width="100%" cellpadding="0" cellspacing="0" style="width: 100%; max-width: 600px;">')

    message2 = df2.to_html(index=False)
    message2 = message2.replace(
        '<table border="1" class="dataframe">',
        '<table border="1" width="100%" cellpadding="0" cellspacing="0" style="width: 100%; max-width: 600px;">')

    message = f'''<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>図書館</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<style>
body{{font-family: "Helvetica Neue", "Helvetica", "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Arial", "Yu Gothic", "Meiryo", sans-serif;}}
</style></head><body>
<p>下線つき日付は最終期限（すでに延長済）</p>
{message}
<p>予約リスト</p>
{message2}
</body></html>
'''
    with open('message.html', 'w', encoding='utf-8') as f:
        f.write(message)
    return message


def send_email(to_email, subject, message, smtp_server, smtp_port_number, smtp_user_name, smtp_password, from_email) -> None:
    msg = MIMEText(message, "html", "utf-8")
    msg["Subject"] = subject
    msg["To"] = to_email
    msg["From"] = from_email

    server = smtplib.SMTP(smtp_server, smtp_port_number)
    server.starttls()
    server.login(smtp_user_name, smtp_password)
    server.send_message(msg)

    server.quit()


def run(condition_name="standard"):
    from config import card_numbers
    from config import to_emails, from_email, title
    from config import smtp_server, smtp_port_number, smtp_user_name, smtp_password
    from config import url
    print(datetime.today())

    keys = list(card_numbers.keys())

    res_extend = {}
    for key in card_numbers.keys():
        print(key)
        res_extend[key] = save_table(key, card_numbers[key], url)

    message = create_email_message(keys, res_extend)
    return_date = generate_earliest_date(keys, res_extend)
    is_available_reserve = availablility_reserve(keys)

    if condition_name == "standard":
        condition = (return_date - timedelta(days=3) <
                     datetime.today()) or (is_available_reserve)
    else:
        condition = True
    if condition:
        for to_email in to_emails:
            send_email(to_email, title, message, smtp_server,
                       smtp_port_number, smtp_user_name, smtp_password, from_email)


def run_and_email_if_errors_exists():
    from config import from_email, index_email_for_error, to_emails
    from config import smtp_server, smtp_port_number, smtp_user_name, smtp_password
    try:
        run()
    except Exception as e:
        send_email(to_emails[index_email_for_error], "Errors during Library cruising", str(e), smtp_server,
                   smtp_port_number, smtp_user_name, smtp_password, from_email)


if __name__ == "__main__":
    # run(condition_name="test")
    run_and_email_if_errors_exists()
