# coding=gbk
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import json


# 判断页面是否加载出来
def isPresent(driver):
    temp = 1
    try:
        driver.find_elements_by_css_selector(
            'div.line-around.layout-box.mod-pagination > a:nth-child(2) > div > select > option')
    except:
        temp = 0
    return temp


# 微博水军引导攻击
def water_army(keyword, username, password, driver, i):
    # cookie登录
    driver.set_window_size(452, 790)
    # 加载驱动，使用浏览器打开指定网址
    driver.get("https://m.weibo.cn/")
    # 判断是否已有cookie文件
    f1 = open('cookie.txt')
    cookie = f1.read()
    if cookie == "":  # 如果cookie未存入cookie.txt，第一次登录
        # 清空浏览器原有cookie
        driver.delete_all_cookies()
        # 使用用户名密码及验证码登录
        # 加载驱动，使用浏览器打开指定网址
        driver.set_window_size(452, 790)
        driver.get("https://passport.weibo.cn/signin/login")
        print("开始自动登陆，若出现验证码手动验证")
        time.sleep(3)

        elem = driver.find_element_by_xpath("//*[@id='loginName']")
        elem.send_keys(username)
        elem = driver.find_element_by_xpath("//*[@id='loginPassword']")
        elem.send_keys(password)
        elem = driver.find_element_by_xpath("//*[@id='loginAction']")
        elem.send_keys(Keys.ENTER)
        print("暂停20秒，用于验证码验证")
        time.sleep(20)

        # 第一次登录把cookie写入文件
        cookies = driver.get_cookies()
        f1 = open('cookie.txt', 'w')
        f1.write(json.dumps(cookies))
        f1.close()
    else:

        cookie = json.loads(cookie)
        for c in cookie:
            driver.add_cookie(c)
        # 刷新页面
        driver.refresh()

    while 1:  # 循环条件为1必定成立
        result = isPresent(driver)
        # 解决输入验证码无法跳转的问题
        driver.get('https://m.weibo.cn/')
        print('判断页面1成功 0失败  结果是=%d' % result)
        if result == 1:
            elems = driver.find_elements_by_css_selector(
                'div.line-around.layout-box.mod-pagination > a:nth-child(2) > div > select > option')
            # return elems #如果封装函数，返回页面
            break
        else:
            print('页面还没加载出来呢')
            time.sleep(20)

    time.sleep(2)
    # 首先定位到右上角的编写微博的div标签
    driver.find_element_by_xpath("//div[@class='lite-iconf lite-iconf-releas']").click()
    time.sleep(2)
    weibo_context = keyword + " 水军引导攻击 " + str(i)
    string_begin = 0
    send_weibo(driver, weibo_context, string_begin)
    print(weibo_context)


# 模拟人进行发微博操作
def send_weibo(driver, weibo_string, string_begin):
    # 由于发出的微博中含有“#”，而输入“#”界面会自动进入选话题界面，然后报异常，
    # 所以在出现异常后返回上一个发微博的界面
    try:
        # 然后定位到输入微博内容的span标签下的textarea标签
        driver.find_element_by_xpath("//span[@class='m-wz-def']/textarea").send_keys(
            weibo_string[string_begin:])  # 输入string_begin之后字符串中的所有内容
        time.sleep(2)
        #  最后定位到右上角发送微博的a标签
        driver.find_element_by_xpath(
            "//div[@class='m-box m-flex-grow1 m-box-model m-fd-row m-aln-center m-justify-end m-flex-base0']/a").click()
        time.sleep(2)
        print("发送一条微博成功")
    except:  # 返回上一个界面，并将微博内容的字符串减去刚才的“#”字符
        string_begin = string_begin + 1
        driver.back()
        time.sleep(2)
        send_weibo(driver, weibo_string, string_begin)


if __name__ == '__main__':
    username = "15586430583"  # 你的微博登录名
    password = "yutao19981119"  # 你的密码
    driver = webdriver.Chrome()  # 你的chromedriver的地址
    temp_filename = "中考"
    keywords = ["#" + temp_filename + "#"]  # 此处可以设置多个话题，#必须要加上
    weibo_num = 100  # 水军微博的数量
    # 载入第一个表格
    for keyword in keywords:
        for i in range(weibo_num):
            water_army(keyword, username, password, driver, i)
    print("所有水军微博发送完毕")
