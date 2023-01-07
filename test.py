import os
from tkinter import *
import time
import datetime
import threading
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import pickle
from selenium.webdriver.support.select import Select

def save_variable(v, filename):
    f = open(filename, 'wb')

    pickle.dump(v,f)

    f.close()

    return filename

def load_variavle(filename):
    f = open(filename, 'rb')

    r = pickle.load(f)

    f.close()

    return r

root = Tk()
root.title('(ww)商品部专用导出程序')

# 商品资料--------1
def shangpinziliao():
    print("启动商品资料线程")
    options = webdriver.ChromeOptions()
    #不弹出下载窗口,设置下载地址
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': 'Z:\\商品采购部\\4.pq表\\1.商品资料\\'}
    options.add_experimental_option('prefs', prefs)

    if var17.get() == 1:
        options.add_argument('--headless')

    driver1 = webdriver.Chrome(options=options)
    driver1.set_page_load_timeout(8)
    try:
        driver1.get('https://ww.erp321.com/jst.aspx')
    except:
        driver1.execute_script('window.stop()')

    driver1.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver1.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver1.maximize_window()
    driver1.get('https://ww.erp321.com/jst.aspx')
    time.sleep(1)
    try:
        driver1.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass

    i = 0
    while True:
        try:
            driver1.find_element_by_id("confirm_close").click()
            break
        except:
            print(BaseException)
            i = i + 1
            if i > 2:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(1)

    time.sleep(3)

    driver1.get(
        'https://www.erp321.com/app/item/item_sku/item_sku_im.aspx')
    time.sleep(3)
    driver1.find_element_by_xpath(
        '//*[@id="searchWhere"]/div[1]/div[1]/ul/li[14]/span[2]').click()  # 点击清空
    time.sleep(3)
    hover = driver1.find_element_by_xpath('//*[@id="Export_Btn"]/span[2]')
    #鼠标放在导出上悬停
    ActionChains(driver1).move_to_element(hover).perform()
    time.sleep(1)

    while True:
        try:
            driver1.find_element_by_css_selector('#Export_Btn > div > div:nth-child(1)').click()
            time.sleep(6)
            # print("商品资料导出成功！")
            Text.insert("end", "商品维护导出成功！\n")
            break
        except:
            time.sleep(3)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    # print("等待8秒开始关闭")
    time.sleep(800)
    driver1.close()


#商品及库存管理--------8
def shangpinkucunguanli():
    print("启动商品资料线程")
    options = webdriver.ChromeOptions()
    #不弹出下载窗口,设置下载地址
    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': 'C:\\Users\\Administrator\\Desktop\\'}
    options.add_experimental_option('prefs', prefs)

    if var17.get() == 1:
        options.add_argument('--headless')

    driver8 = webdriver.Chrome(options=options)
    driver8.set_page_load_timeout(8)
    try:
        driver8.get('https://ww.erp321.com/jst.aspx')
    except:
        driver8.execute_script('window.stop()')

    driver8.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver8.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver8.maximize_window()
    driver8.get('https://ww.erp321.com/jst.aspx')
    time.sleep(1)
    try:
        driver8.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass

    i = 0
    while True:
        try:
            driver8.find_element_by_id("confirm_close").click()
            break
        except:
            print(BaseException)
            i = i + 1
            if i > 2:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(1)

    time.sleep(3)
    driver8.get('https://src.erp321.com/erp-web-group/goods/goodsInventoryManagement')
    time.sleep(3)
    driver8.find_element_by_xpath(
        '//*[@id="root"]/div/div/div[1]/div[1]/div[3]/div[2]/div/div/div[1]/div[1]/div/div[1]/div/div/div/span[2]').click()  # 点击清空
    time.sleep(3)
    hover = driver8.find_element_by_xpath(
        '//*[@id="root"]/div/div/div[1]/div[1]/div[3]/div[2]/div/div/div[2]/div[1]/div[1]/button')
    # 鼠标放在导出上悬停
    ActionChains(driver8).move_to_element(hover).perform()
    time.sleep(1)
    while True:
        try:
            driver8.find_element_by_xpath('/html/body/div[6]/div/div/ul/li[6]/span/span').click()
            time.sleep(6)
            # print("商品资料导出成功！")
            Text.insert("end", "商品及库存管理导出成功！\n")
            break
        except:
            time.sleep(3)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    # print("等待8秒开始关闭")
    time.sleep(800)
    driver8.close()


# 商品主题分析---------5
def shangpinzhutifenxi():
    options = webdriver.ChromeOptions()

    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': 'Z:\\商品采购部\\4.pq表\\5.商品主题分析\\'}

    options.add_experimental_option('prefs', prefs)
    if var17.get() == 1:
        options.add_argument('--headless')
    print("商品主题分析线程")
    driver5 = webdriver.Chrome(options=options)
    driver5.set_page_load_timeout(15)
    try:
        driver5.get('https://ww.erp321.com/jst.aspx')
    except:
        driver5.execute_script('window.stop()')

    driver5.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver5.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver5.maximize_window()
    driver5.get('https://ww.erp321.com/jst.aspx')
    time.sleep(2)
    try:
        driver5.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass
    i = 0
    while True:
        try:
            driver5.find_element_by_id("confirm_close").click()
            break
        except:

            i = i + 1
            if i > 2:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(0.5)

    # 点击搜索按钮
    driver5.find_element_by_xpath('//*[@id="side"]/div[2]').click()
    driver5.find_element_by_xpath('//*[@id="side-search-key"]').click()
    driver5.find_element_by_xpath('//*[@id="side-search-key"]').send_keys("商品主题分析")
    driver5.find_element_by_xpath('//*[@id="side"]/div[4]/div[1]/div[3]/div[2]/div/span').click()
    time.sleep(1)
    # 报表

    ifram0 = driver5.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    driver5.switch_to.frame(ifram0)
    print(ifram0)
    time.sleep(3)

    # 跳转到旧版
    # driver5.find_element_by_xpath('//*[@id="transferOldVBtn"]').click()
    # time.sleep(3)
    #网页内嵌，需要先定位ifram再转换
    ifram1 = driver5.find_element_by_xpath('//*[@id="s_filter_frame"]')
    driver5.switch_to.frame(ifram1)
    print(ifram1)

    time.sleep(3)
    driver5.find_element_by_xpath('//*[@id="button_bar"]/span[1]').click()
    time.sleep(1)
    driver5.find_element_by_xpath('//*[@id="search_list"]/div[7]').click()
    time.sleep(11)


    driver5.switch_to.default_content()
    ifram0 = driver5.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    # print("iframe" + str(ifram0))
    driver5.switch_to.frame(ifram0)

    ifram_tab0 = driver5.find_element_by_xpath('//*[@id="tab0"]')
    driver5.switch_to.frame(ifram_tab0)
    time.sleep(3)
    while True:
        try:
            # 导出
            driver5.find_element_by_xpath('//*[@id="form1"]/div[4]/span[12]/span/span').click()
            time.sleep(60)
            Text.insert("end", "商品主题分析导出成功！\n")
            break
        except:
            time.sleep(0.5)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    # print("等待3秒开始关闭")
    time.sleep(20000)
    driver5.close()


# 采购订单信息 --------2
def caigoudanxinxi():
    # 采购单信息------------3
    print("启动采购订单信息线程")
    options = webdriver.ChromeOptions()

    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': 'Z:\\商品采购部\\4.pq表\\2.采购订单信息\\'}

    options.add_experimental_option('prefs', prefs)
    if var17.get() == 1:
        options.add_argument('--headless')

    driver3 = webdriver.Chrome(options=options)
    driver3.set_page_load_timeout(15)
    try:
        driver3.get('https://ww.erp321.com/jst.aspx')
    except:
        driver3.execute_script('window.stop()')

    driver3.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver3.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver3.maximize_window()
    driver3.get('https://ww.erp321.com/jst.aspx')
    time.sleep(2)
    try:
        driver3.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass
    # information=driver.find_element_by_class_name("side-menu-cell side-menu-cell-usual").click()
    i = 0
    while True:
        try:
            confirm_body = driver3.find_element_by_id("confirm_close").click()
            break
        except:
            i = i + 1
            if i > 3:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(1)

    # 到导出数据的页面
    # 点击搜索按钮

    driver3.find_element_by_xpath('//*[@id="side"]/div[2]').click()
    driver3.find_element_by_xpath('//*[@id="side-search-key"]').click()
    driver3.find_element_by_xpath('//*[@id="side-search-key"]').send_keys("采购主题分析")
    driver3.find_element_by_xpath('//*[@id="side"]/div[4]/div[1]/div[3]/div[2]/div/span').click()

    # 第一个
    iframe1 = driver3.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    driver3.switch_to.frame(iframe1)
    # print(iframe1)
    # 第二个

    driver3.switch_to.frame("s_filter_frame")
    driver3.find_element_by_xpath('//*[@id="button_bar"]/span[1]').click()
    time.sleep(0.5)
    driver3.find_element_by_xpath('//*[@id="search_list"]/div[1]').click()
    driverSize = driver3.get_window_size()
    time.sleep(2)
    while True:
        try:
            driver3.maximize_window()
            print(driverSize)
            driver3.switch_to.default_content()
            iframe = driver3.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
            driver3.switch_to.frame(iframe)
            driver3.find_element_by_xpath('//*[@id="tabs"]/div[1]/div[13]').click()
            break
        except:
            time.sleep(2)
            print("等待")
            
    # 恢复窗口大小
    time.sleep(1)
    driver3.set_window_size(driverSize['width'], driverSize['height'], windowHandle='current')
    #所有帧上移除并在页面切换焦点，重新开始定位
    driver3.switch_to.default_content()
    ifram0 = driver3.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    # print("iframe" + str(ifram0))
    driver3.switch_to.frame(ifram0)
    ifram_tab0 = driver3.find_element_by_xpath('//*[@id="tab1"]')
    driver3.switch_to.frame(ifram_tab0)

    # # 导出
    while True:
        try:
            time.sleep(5)
            driver3.find_element_by_xpath(
                '//*[@id="form1"]/div[4]/span[12]').click()  # //*[@id="form1"]/div[4]/span[12]
            # print("采购单概况导出成功！")
            Text.insert("end", "采购订单导出成功！\n")
            break
        except:
            time.sleep(1)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    # print("等待8秒开始关闭")
    time.sleep(200)
    driver3.close()


# 进出仓明细----------3
def jinchucangmingxi():
    # 进出仓明细------------4

    print("启动进出仓明细线程")

    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': 'Z:\\商品采购部\\4.pq表\\3.进出仓明细\\'}

    if var17.get() == 1:
        options.add_argument('--headless')

    options.add_experimental_option('prefs', prefs)

    driver4 = webdriver.Chrome(options=options)
    driver4.set_page_load_timeout(15)
    try:
        driver4.get('https://ww.erp321.com/jst.aspx')
    except:
        driver4.execute_script('window.stop()')

    driver4.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver4.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver4.maximize_window()
    driver4.get('https://ww.erp321.com/jst.aspx')
    time.sleep(2)
    try:
        driver4.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass
    i = 0
    while True:
        try:
            driver4.find_element_by_id("confirm_close").click()
            break
        except:

            i = i + 1
            if i > 3:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(1)

    # 点击搜索按钮
    driver4.find_element_by_xpath('//*[@id="side"]/div[2]').click()
    driver4.find_element_by_xpath('//*[@id="side-search-key"]').click()
    driver4.find_element_by_xpath('//*[@id="side-search-key"]').send_keys("采购主题分析")
    driver4.find_element_by_xpath('//*[@id="side"]/div[4]/div[1]/div[3]/div[2]/div/span').click()
    time.sleep(2)
    # # 第一个
    iframe1 = driver4.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    driver4.switch_to.frame(iframe1)
    print(iframe1)
    # # 第二个
    # driver.switch_to.frame("s_filter_frame")
    iframe2 = driver4.find_element_by_xpath('//*[@id="s_filter_frame"]')
    driver4.switch_to.frame(iframe2)

    # 点击查询池
    driver4.find_element_by_xpath('//*[@id="button_bar"]/span[1]').click()
    #
    time.sleep(0.5)
    driver4.find_element_by_xpath('//*[@id="search_list"]/div[2]').click()
    driver4.get_window_size()

    # 点击进出仓明细
    driver4.switch_to.default_content()
    ifram0 = driver4.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    driver4.switch_to.frame(ifram0)

    driver4.find_element_by_xpath('//*[@id="tabs"]/div[1]/div[9]').click()
    time.sleep(4)

    ifram_tab0 = driver4.find_element_by_xpath('//*[@id="tab1"]')
    driver4.switch_to.frame(ifram_tab0)
    # 导出
    while True:
        try:
            time.sleep(5)
            driver4.find_element_by_xpath('//*[@id="form1"]/div[4]/span[12]/span/span').click()
            # print("进出仓明细导出成功！")
            Text.insert("end", "进出仓明细导出成功！\n")
            break

        except:
            time.sleep(1)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    time.sleep(1000)
    driver4.close()


# 销售明细分析（财务）--------11
def xiaoshouximingfenxicaiwu():
    # 销售主题分析------------5
    print("启动销售主题分析财务线程")
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': 'Z:\\商品采购部\\4.pq表\\11.销售明细分析（财务）\\'}
    if var17.get() == 1:
        options.add_argument('--headless')

    options.add_experimental_option('prefs', prefs)

    driver11 = webdriver.Chrome(options=options)
    driver11.set_page_load_timeout(15)
    try:
        driver11.get('https://ww.erp321.com/jst.aspx')
    except:
        driver11.execute_script('window.stop()')

    driver11.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver11.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver11.maximize_window()
    driver11.get('https://ww.erp321.com/jst.aspx')

    # 商品资料
    i = 0
    while True:
        try:
            driver11.find_element_by_id("confirm_close").click()
            break
        except:
            # print("未检测到弹窗")
            i = i + 1
            if i > 2:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(0.5)
    # 点击搜索按钮、
    driver11.find_element_by_xpath('//*[@id="side"]/div[2]').click()
    driver11.find_element_by_xpath('//*[@id="side-search-key"]').click()
    driver11.find_element_by_xpath('//*[@id="side-search-key"]').send_keys("销售主题分析(财务)")
    driver11.find_element_by_xpath('//*[@id="side"]/div[4]/div[1]/div[3]/div[2]/div/span').click()

    time.sleep(3)

    ifram0 = driver11.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    # print(ifram0)
    driver11.switch_to.frame(ifram0)
    #
    ifram1 = driver11.find_element_by_xpath('//*[@id="s_filter_frame"]')
    driver11.switch_to.frame(ifram1)
    # print(ifram1)

    driver11.find_element_by_css_selector('#button_bar > span:nth-child(1)').click()  ##button_bar > span:nth-child(1)
    time.sleep(3)

    driver11.find_element_by_css_selector('#search_list > div').click()
    time.sleep(5)

    # # 销售订单明细

    driver11.switch_to.default_content()
    iframe = driver11.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    driver11.switch_to.frame(iframe)
    time.sleep(6)
    driverSize = driver11.get_window_size()
    driver11.maximize_window()
    time.sleep(6)
    driver11.find_element_by_css_selector('#tabs > div.tabbar_bar.big.top.tabbar_skin_tab > div:nth-child(13)').click()
    time.sleep(6)
    driver11.switch_to.default_content()
    ifram0 = driver11.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    # print("iframe"+str(ifram0))
    driver11.switch_to.frame(ifram0)

    ifram_tab0 = driver11.find_element_by_xpath('//*[@id="tab1"]')
    driver11.switch_to.frame(ifram_tab0)
    driver11.set_window_size(driverSize['width'], driverSize['height'], windowHandle='current')
    time.sleep(6)
    while True:
        try:
            time.sleep(5)
            driver11.find_element_by_css_selector('#form1 > div.rpt.btn_bar > span.btn_line.h30 > span > span').click()
            # print("销售主题分析财务导出成功！")
            Text.insert("end", "销售主题分析(财务)导出成功！\n")
            break
        except:
            time.sleep(1)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")
    # print("等待8秒开始关闭")
    time.sleep(20000)
    driver11.close()


# 各仓库在途库存数据------4
def shangpinkucunfencang():
    print("启动商品库存分仓线程")
    options = webdriver.ChromeOptions()

    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': 'Z:\\商品采购部\\4.pq表\\4.各仓库在途库存数据\\'}

    options.add_experimental_option('prefs', prefs)
    if var17.get() == 1:
        options.add_argument('--headless')
    driver4 = webdriver.Chrome(options=options)
    driver4.set_page_load_timeout(15)
    try:
        driver4.get('https://ww.erp321.com/jst.aspx')
    except:
        driver4.execute_script('window.stop()')

    driver4.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver4.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver4.maximize_window()
    driver4.get('https://ww.erp321.com/jst.aspx')

    # 商品资料
    i = 0
    while True:
        try:
            driver4.find_element_by_id("confirm_close").click()
            break
        except:
            # print("未检测到弹窗")
            i = i + 1
            if i > 2:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(0.5)
    # 报表
    time.sleep(2)
    try:
        driver4.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass
    driver4.find_element_by_xpath('//*[@id="side"]/div[2]').click()
    driver4.find_element_by_xpath('//*[@id="side-search-key"]').click()
    driver4.find_element_by_xpath('//*[@id="side-search-key"]').send_keys("商品库存（分仓）")
    driver4.find_element_by_xpath('//*[@id="side"]/div[4]/div[1]/div[3]/div[2]/div/span').click()  # 商品库存
    # 报表
    time.sleep(2)

    driver4.switch_to.default_content()  # 回到主框架
    # 用xpath定位
    ifram0 = driver4.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    # print(ifram0)
    driver4.switch_to.frame(ifram0)
    #  排除次品
    cip = driver4.find_element_by_class_name('_cbb_p')
    ActionChains(driver4).move_to_element(cip).perform()
    driver4.find_element_by_css_selector('#wms_co_id_p > span > div > div > div:nth-child(2) > label').click()
    time.sleep(2)
    driver4.find_element_by_css_selector('#wms_co_id_p > span > div > div > div:nth-child(3) > label').click()
    time.sleep(2)
    driver4.find_element_by_css_selector('#wms_co_id_p > span > div > div > div:nth-child(5) > label').click()
    time.sleep(3)
    driver4.find_element_by_css_selector(
        '#form1 > table.full.fixed.padding_left_bar > tbody > tr.top_toolbar > td > ul > li:nth-child(11) > span.btn_search').click()
    time.sleep(3)
    hover = driver4.find_element_by_xpath('//*[@id="Tool_New_Btn"]/span[2]')
    time.sleep(3)
    ActionChains(driver4).move_to_element(hover).perform()
    time.sleep(2)

    while True:
        try:
            driver4.find_element_by_xpath('//*[@id="Tool_New_Btn"]/div/div[1]').click()
            # print("商品库存分仓导出成功！")
            Text.insert("end", "各仓库在途库存数据导出成功！\n")
            break
        except:
            time.sleep(3)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    # print("等待8秒开始关闭")
    # driver6.set_window_size(driverSize['width'], driverSize['height'])
    time.sleep(100)
    driver4.close()


# 箱及仓位库存(本仓)----------13
def jihuacaigoujianyiBen():
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': 'Z:\\商品采购部\\4.pq表\\13.箱及仓位库存(本仓)\\'}
    if var17.get() == 1:
        options.add_argument('--headless')
    options.add_experimental_option('prefs', prefs)

    driver13 = webdriver.Chrome(options=options)

    driver13.get('https://ww.erp321.com/jst.aspx')

    driver13.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver13.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver13.maximize_window()
    driver13.get('https://ww.erp321.com/jst.aspx')

    i = 0
    while True:
        try:
            confirm_body = driver13.find_element_by_id("confirm_close").click()
            break
        except:
            print("未检测到弹窗")
            i = i + 1
            if i > 2:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(0.5)
    # information=driver.find_element_by_xpath("//div[@class='side-menu']/div[3]").click()
    # information=driver.find_element_by_xpath("//div[@class='side-menu-cell side-menu-cell-usual']/div[2]").click()
    time.sleep(2)

    # 点击左上角
    driver13.find_element_by_css_selector('#side > div.side-logo.bounceIn.animated').click()
    time.sleep(1)
    # 点击
    driver13.find_element_by_xpath('//*[@id="side-search-key"]').click()
    driver13.find_element_by_xpath('//*[@id="side-search-key"]').send_keys("箱及仓位库存")
    driver13.find_element_by_xpath('//*[@id="side"]/div[4]/div[1]/div[3]/div[2]/div').click()

    #
    iframe = driver13.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    driver13.switch_to.frame(iframe)

    # 点击改变商家
    time.sleep(3)
    driver13.find_element_by_xpath('//*[@id="plb_bar"]').click()
    time.sleep(1)
    driver13.find_element_by_xpath('//*[@id="plb10962175_10962175"]').click()
    time.sleep(0.4)
    driver13.find_element_by_xpath('//*[@id="form1"]/div[4]/div[3]/span[2]').click()
    print("点击完成")
    driver13.maximize_window()
    time.sleep(3)
    #刷新商品名
    driver13.find_element_by_css_selector('#btnRefreshSku').click()
    time.sleep(2)
    driver13.switch_to_default_content()
    time.sleep(2)
    driver13.find_element_by_id('msg_close').click()
    # 清空
    iframe1 = driver13.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    driver13.switch_to.frame(iframe1)
    driver13.find_element_by_css_selector(
        '#form1 > div.padding_left_bar.full > table > tbody > tr.top_toolbar > td > ul > li:nth-child(12) > span > span.btn_search').click()
    time.sleep(3)
    while True:
        try:
            move = driver13.find_element_by_css_selector('#_jt_toolbar_right')
            ActionChains(driver13).move_to_element(move).perform()  # 鼠标悬停
            driver13.find_element_by_css_selector('#btnExport').click()
            Text.insert("end", "箱及仓位库存(本仓)导出成功！\n")
            break
        except:
            time.sleep(0.5)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    time.sleep(2022)
    driver13.close()


# 未发货订单------------6
def weifahuodingdan():
    print("启动未发货订单线程")
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': 'Z:\\商品采购部\\4.pq表\\6.未发货订单统计\\'}
    options.add_experimental_option('prefs', prefs)
    if var17.get() == 1:
        options.add_argument('--headless')

    driver6 = webdriver.Chrome(options=options)
    driver6.set_page_load_timeout(15)
    try:
        driver6.get('https://ww.erp321.com/jst.aspx')
    except:
        driver6.execute_script('window.stop()')
    driver6.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver6.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver6.maximize_window()
    driver6.get('https://ww.erp321.com/jst.aspx')
    time.sleep(1)
    try:
        driver6.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass
    # 商品资料
    # information=driver.find_element_by_class_name("side-menu-cell side-menu-cell-usual").click()
    i = 0
    while True:
        try:
            confirm_body = driver6.find_element_by_id("confirm_close").click()
            break
        except:
            # print("未检测到弹窗")
            i = i + 1
            if i > 2:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(0.5)
    # information=driver.find_element_by_xpath("//div[@class='side-menu']/div[3]").click()
    # information=driver.find_element_by_xpath("//div[@class='side-menu-cell side-menu-cell-usual']/div[2]").click()

    # 点击搜索按钮
    driver6.find_element_by_xpath('//*[@id="side"]/div[2]').click()
    driver6.find_element_by_xpath('//*[@id="side-search-key"]').click()
    driver6.find_element_by_xpath('//*[@id="side-search-key"]').send_keys("未发货订单统计")
    driver6.find_element_by_xpath('//*[@id="side"]/div[4]/div[1]/div[3]/div[2]/div/span').click()

    ifram0 = driver6.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    # print(ifram0)
    driver6.switch_to.frame(ifram0)

    ifram1 = driver6.find_element_by_xpath('//*[@id="s_filter_frame"]')
    driver6.switch_to.frame(ifram1)
    # print(ifram1)

    # 点击查询池
    driver6.find_element_by_xpath('//*[@id="button_bar"]/span[1]').click()
    time.sleep(2)
    driver6.find_element_by_xpath('//*[@id="search_list"]/div[11]').click()
    time.sleep(5)

    # 点击明细（订单商品）
    driver6.switch_to.default_content()
    ifram0 = driver6.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    # print("iframe" + str(ifram0))
    driver6.switch_to.frame(ifram0)
    time.sleep(1)
    driver6.find_element_by_xpath('//*[@id="tabs"]/div[1]/div[3]').click()
    time.sleep(8)

    ifram_tab1 = driver6.find_element_by_xpath('//*[@id="tab1"]')
    driver6.switch_to.frame(ifram_tab1)
    # 导出
    while True:
        try:

            time.sleep(3)
            comfirmdel = driver6.find_element_by_css_selector(
                '#form1 > div.rpt.btn_bar > span.btn_line.h30 > span > span')
            #滚动条滚动到可以导出的位置
            driver6.execute_script("arguments[0].click();", comfirmdel)
            # print("采购单概况导出成功！")
            Text.insert("end", "发货订单导出成功！\n")
            break
        except:
            time.sleep(3)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    # print("等待8秒开始关闭")
    time.sleep(300)
    driver6.close()


# 箱及仓位库存(卓尚分仓)---------10
def xiangjicangweikucunFen():
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': 'Z:\\商品采购部\\4.pq表\\10.箱及仓位库存(卓尚分仓)\\'}
    if var17.get() == 1:
        options.add_argument('--headless')
    options.add_experimental_option('prefs', prefs)

    driver10 = webdriver.Chrome(options=options)

    driver10.get('https://ww.erp321.com/jst.aspx')

    driver10.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver10.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver10.maximize_window()
    driver10.get('https://ww.erp321.com/jst.aspx')

    i = 0
    while True:
        try:
            driver10.find_element_by_id("confirm_close").click()
            break
        except:
            print("未检测到弹窗")
            i = i + 1
            if i > 2:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(0.5)
    # information=driver.find_element_by_xpath("//div[@class='side-menu']/div[3]").click()
    # information=driver.find_element_by_xpath("//div[@class='side-menu-cell side-menu-cell-usual']/div[2]").click()
    time.sleep(2)
    try:
        driver10.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass

    # 点击左上角
    driver10.find_element_by_xpath('//*[@id="side"]/div[2]').click()
    time.sleep(1)
    # 点击
    driver10.find_element_by_xpath('//*[@id="side-search-key"]').click()
    driver10.find_element_by_xpath('//*[@id="side-search-key"]').send_keys("箱及仓位库存")
    driver10.find_element_by_xpath('//*[@id="side"]/div[4]/div[1]/div[3]/div[2]/div').click()

    iframe = driver10.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    driver10.switch_to.frame(iframe)

    # 点击改变商家
    time.sleep(3)
    driver10.find_element_by_xpath('//*[@id="plb_bar"]').click()
    time.sleep(1)
    driver10.find_element_by_css_selector('#plb10962175_10989645 > div.plb_item_ap').click()  #
    time.sleep(0.4)
    driver10.find_element_by_xpath('//*[@id="form1"]/div[4]/div[3]/span[2]').click()
    print("点击完成")
    driver10.maximize_window()
    time.sleep(3)
    #刷新商品名
    driver10.find_element_by_css_selector('#btnRefreshSku').click()
    time.sleep(2)
    driver10.switch_to_default_content()
    time.sleep(1)
    driver10.find_element_by_id('msg_close').click()
    iframe1 = driver10.find_element_by_xpath('//*[@id="frame_list"]/iframe[2]')
    driver10.switch_to.frame(iframe1)
    time.sleep(1)
    # 清空
    driver10.find_element_by_css_selector(
        '#form1 > div.padding_left_bar.full > table > tbody > tr.top_toolbar > td > ul > li:nth-child(12) > span > span.btn_search_reset').click()
    time.sleep(3)
    # 导出
    # driver9.refresh()

    while True:
        try:
            # button = driver10.find_element_by_xpath("/html/body/form/div[6]/div[2]/div[1]/table/tbody/tr/td[2]/span[1]")
            # driver10.execute_script("arguments[0].click()", button)
            driver10.find_element_by_id('btnExport').click()
            # print('导出成功')
            Text.insert("end", "箱及仓位库存(卓尚分仓)导出成功！\n")
            break
        except:
            time.sleep(0.5)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    time.sleep(3300)
    driver10.close()


# 入库汇总--------14
def rukuhuizong():
    print("启动(财务结算)入库汇总查询线程")
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': 'Z:\\商品采购部\\4.pq表\\14.(财务结算)入库汇总查询\\'}

    options.add_experimental_option('prefs', prefs)
    if var17.get() == 1:
        options.add_argument('--headless')

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(8)
    try:
        driver.get('https://ww.erp321.com/jst.aspx')
    except:
        driver.execute_script('window.stop()')

    driver.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver.maximize_window()
    driver.get('https://ww.erp321.com/jst.aspx')
    try:
        driver.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass

    # 商品资料----------1
    # information=driver.find_element_by_class_name("side-menu-cell side-menu-cell-usual").click()
    i = 0
    while True:
        try:
            driver.find_element_by_id("confirm_close").click()
            break
        except:
            print(BaseException)
            i = i + 1
            if i > 2:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(1)

    time.sleep(3)

    # 进入入库汇总查询界面
    driver.get('https://www.erp321.com/app/FMS/fmscommon/inoutsum/inoutsumhistory.aspx?method=search')
    time.sleep(8)
    # 更改时间
    driver.find_element_by_css_selector('#date1').click()
    time.sleep(1)
    a= Select(driver.find_element_by_css_selector(
        '#ui-datepicker-div > div > div > select.ui-datepicker-year'))  # ui-datepicker-div > div > div > select.ui-datepicker-year
    #ActionChains(driver).click(a).perform()
    # print('1')
    time.sleep(2)
    a.select_by_value('2018')
    time.sleep(1)
    #driver.find_element_by_xpath('//*[@id="ui-datepicker-div"]/div/div/select[1]/option[11]').click()
    driver.find_element_by_xpath(
        '//*[@id="ui-datepicker-div"]/table/tbody/tr[2]/td[4]').click()
    # print('2')
    time.sleep(3)

    # 点击搜索按钮
    a = driver.find_element_by_css_selector(
        '#trcondition > td > table > tbody > tr > td:nth-child(5) > span:nth-child(1)')
    a.click()  # trcondition > td > table > tbody > tr > td:nth-child(5) > span:nth-child(1)
    time.sleep(3)

    # 导出
    while True:
        try:
            driver.find_element_by_css_selector(
                "#trcondition > td > table > tbody > tr > td:nth-child(5) > span:nth-child(3) > span > span").click()
            time.sleep(6)
            # print("商品资料导出成功！")
            Text.insert("end", "(财务结算)入库汇总查询导出成功！\n")
            break
        except:
            time.sleep(1)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    # print("等待8秒开始关闭")
    time.sleep(40)
    driver.close()


# 预付款汇总查询---------16
def yufukuanhuizong():
    options = webdriver.ChromeOptions()

    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': 'Z:\\商品采购部\\4.pq表\\16.(财务结算)预付款汇总查询\\'}

    options.add_experimental_option('prefs', prefs)
    if var17.get() == 1:
        options.add_argument('--headless')

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(15)
    try:
        driver.get('https://ww.erp321.com/jst.aspx')
    except:
        driver.execute_script('window.stop()')

    driver.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver.maximize_window()
    driver.get('https://ww.erp321.com/jst.aspx')
    try:
        driver.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass
    i = 0
    while True:
        try:
            driver.find_element_by_id("confirm_close").click()
            break
        except:

            i = i + 1
            if i > 2:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(0.5)

    time.sleep(5)
    # 进入预付款汇总查询界面
    driver.get('https://www.erp321.com/app/FMS/fmscommon/prepaysum/prepaysumhistory.aspx?method=search')
    time.sleep(5)

    # 更改时间
    driver.find_element_by_css_selector('#date1').click()
    time.sleep(1)
    a = Select(driver.find_element_by_css_selector('#ui-datepicker-div > div > div > select.ui-datepicker-year'))
    time.sleep(1)
    a.select_by_value('2018')
    time.sleep(1)
    driver.find_element_by_css_selector(
        '#ui-datepicker-div > table > tbody > tr:nth-child(1) > td:nth-child(7) > a').click()
    time.sleep(1)

    # 点击搜索按钮
    driver.find_element_by_css_selector(
        '#trcondition > td > table > tbody > tr > td:nth-child(5) > span:nth-child(1)').click()
    time.sleep(3)

    # 导出
    while True:
        try:
            driver.find_element_by_css_selector(
                "#trcondition > td > table > tbody > tr > td:nth-child(5) > span:nth-child(3) > span > span").click()
            time.sleep(6)
            # print("商品资料导出成功！")
            Text.insert("end", "(财务结算)预付款汇总查询导出成功\n")
            break
        except:
            time.sleep(1)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    # print("等待3秒开始关闭")
    time.sleep(20000)
    driver.close()


# 应付付款 --------15
def yingfufukuanhuizong():
    # 采购单信息------------3
    # print("启动采购订单信息线程")
    options = webdriver.ChromeOptions()

    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': 'Z:\\商品采购部\\4.pq表\\15.(财务结算)应付付款汇总查询\\'}

    options.add_experimental_option('prefs', prefs)
    if var17.get() == 1:
        options.add_argument('--headless')

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(15)
    try:
        driver.get('https://ww.erp321.com/jst.aspx')
    except:
        driver.execute_script('window.stop()')

    driver.delete_all_cookies()

    for cookie in savedCookies:
        # k代表着add_cookie的参数cookie_dict中的键名，这次我们要传入这5个键
        for k in {'name', 'value', 'domain', 'path', 'expiry'}:
            # cookie.keys()属于'dict_keys'类，通过list将它转化为列表
            if k not in list(cookie.keys()):
                # saveCookies中的第一个元素，由于记录的是登录前的状态，所以它没有'expiry'的键名，我们给它增加
                if k == 'expiry':
                    t = time.time()
                    cookie[k] = int(t)  # 时间戳s
        # 将每一次遍历的cookie中的这五个键名和键值添加到cookie
        driver.add_cookie({k: cookie[k] for k in {'name', 'value', 'domain', 'path', 'expiry'}})
    driver.maximize_window()
    driver.get('https://ww.erp321.com/jst.aspx')
    try:
        driver.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass

    # information=driver.find_element_by_class_name("side-menu-cell side-menu-cell-usual").click()
    i = 0
    while True:
        try:
            driver.find_element_by_id("confirm_close").click()
            break
        except:
            i = i + 1
            if i > 3:
                print("超过3秒未找到弹窗" + str(i))
                break
            time.sleep(1)

    time.sleep(3)
    # 进入预付款汇总查询界面
    driver.get('https://www.erp321.com/app/FMS/fmscommon/copeandpaysum/copeandpaysum.aspx?method=search&ts___=1669619950303&am___=LoadDataToJSON')
    time.sleep(8)

    # 更改时间
    driver.find_element_by_css_selector('#date1').click()
    time.sleep(1)
    a=Select(driver.find_element_by_css_selector('#ui-datepicker-div > div > div > select.ui-datepicker-year'))
    time.sleep(1)
    a.select_by_value('2018')
    time.sleep(1)
    driver.find_element_by_css_selector(
        '#ui-datepicker-div > table > tbody > tr:nth-child(1) > td:nth-child(7) > a').click()
    time.sleep(1)

    # 点击搜索按钮
    driver.find_element_by_css_selector(
        '#trcondition > td > table > tbody > tr > td:nth-child(5) > span:nth-child(1)').click()
    time.sleep(3)

    # 导出
    while True:
        try:
            driver.find_element_by_css_selector(
                "#trcondition > td > table > tbody > tr > td:nth-child(5) > span:nth-child(3) > span > span").click()
            time.sleep(6)
            # print("商品资料导出成功！")
            Text.insert("end", "(财务结算)应付付款汇总查询导出成功\n")
            break
        except:
            time.sleep(1)
            # print("等待导出按钮")
            Text.insert("end", "等待导出按钮\n")

    # print("等待8秒开始关闭")
    time.sleep(50)
    driver.close()

#获取屏幕分辨率，屏幕的宽和高
sw = root.winfo_screenwidth()
sh = root.winfo_screenheight()
x = (sw - 500) / 2
y = (sh - 500) / 2

root.geometry(f'500x500+{int(x)}+{int(y)}')  # 这里的乘号不是 * ，而是小写英文字母 x


def tes():
    for i in range(14):
        tt = threading.Thread(target=timedelay)
        tt.setDaemon = True
        tt.start()


def timedelay():
    Text.insert("end", f"开始延迟5秒\n")
    time.sleep(50)
    Text.insert("end", f"延迟完成\n")
    Text.see("end")


def check():
    Text.insert("end", "开始更新cookie\n")
    t1 = threading.Thread(target=checkcookie)
    t1.setDaemon = True
    t1.start()


def checkcookie():
    global savedCookies
    chrome_options = webdriver.ChromeOptions()
    #禁止图片和css的加载
    prefs = {"profile.managed_default_content_settings.images": 2,
             'permissions.default.stylesheet': 2}
    chrome_options.add_experimental_option("prefs", prefs)
    driver_cookie = webdriver.Chrome(options=chrome_options)
    driver_cookie.get('https://ww.erp321.com/jst.aspx')
    time.sleep(3)
    login_id = driver_cookie.find_element_by_id("login_id")

    login_id.send_keys("19906618313")
    login_password = driver_cookie.find_element_by_id("password").send_keys("1a2b3c19906618313")
    driver_cookie.find_element_by_class_name("login-sub").click()
    time.sleep(2)
    try:
        driver_cookie.find_element_by_css_selector('body > div.newHome-dialog > div > img').click()
    except:
        pass
    savedCookies = driver_cookie.get_cookies()
    # print(savedCookies)
    timestamp_now = int(time.time())
    save_variable(timestamp_now, 'timestamp.zs')
    save_variable(savedCookies, 'variable.zs')
    Text.insert("end", "cookie更新完成\n")
    time.sleep(50)
    driver_cookie.close()


def button_start():
    # Text.insert("end",var1.get())
    # Text.see("end")

    if var1.get() == 1:
        t1 = threading.Thread(target=shangpinziliao)
        t1.setDaemon = True
        t1.start()
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<商品维护>>\n")
        # Text.insert("end",dt_ms + "：延迟等待下个线程\n")
        Text.see("end")

    if var2.get() == 1:
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<采购订单信息>>\n")

        t2 = threading.Thread(target=caigoudanxinxi)
        t2.setDaemon = True
        t2.start()
        # Text.insert("end", "延迟等待下个线程\n")

    if var3.get() == 1:
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<进出仓明细>>\n")
        t3 = threading.Thread(target=jinchucangmingxi)
        t3.setDaemon = True
        t3.start()
        # Text.insert("end", "延迟等待下个线程\n")
        Text.see("end")

    if var4.get() == 1:
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<各仓库在途库存数据>>\n")
        t4 = threading.Thread(target=shangpinkucunfencang)
        t4.setDaemon = True
        t4.start()
        Text.see("end")

    if var5.get() == 1:
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<商品主题分析>>\n")
        t5 = threading.Thread(target=shangpinzhutifenxi)
        t5.setDaemon = True
        t5.start()
        Text.see("end")

    if var6.get() == 1:
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<未发货订单>>\n")

        #print("导出<<商品主题分析>>\n")
        t6 = threading.Thread(target=weifahuodingdan)
        t6.setDaemon = True
        t6.start()
        Text.see("end")

    # if var8.get() == 1:
    #     dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
    #     Text.insert("end", dt_ms + "开始导出<<商品及库存管理>>\n")
    #
    #     #print("导出<<商品及库存管理>>\n")
    #     t6 = threading.Thread(target=shangpinkucunguanli)
    #     t6.setDaemon = True
    #     t6.start()
    #     Text.see("end")


    if var13.get() == 1:
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<箱及仓位库存(本仓)>>\n")
        Text.see("end")

        t8 = threading.Thread(target=jihuacaigoujianyiBen)
        t8.setDaemon = True
        t8.start()

    if var11.get() == 1:
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<销售明细分析(财务)>>\n")

        Text.see("end")
        t9 = threading.Thread(target=xiaoshouximingfenxicaiwu)
        t9.setDaemon = True
        t9.start()

    if var10.get() == 1:
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<箱及仓位库存(卓尚分仓)>>\n")
        Text.see("end")
        t10 = threading.Thread(target=xiangjicangweikucunFen)
        t10.setDaemon = True
        t10.start()

    if var14.get() == 1:
        t14 = threading.Thread(target=rukuhuizong)
        t14.setDaemon = True
        t14.start()
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<入库汇总查询数据>>\n")
        Text.see("end")

    if var15.get() == 1:
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<应付付款汇总查询数据>>\n")

        t15 = threading.Thread(target=yingfufukuanhuizong)
        t15.setDaemon = True
        t15.start()
        # Text.insert("end", "延迟等待下个线程\n")

    if var16.get() == 1:
        dt_ms = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S：')
        Text.insert("end", dt_ms + "开始导出<<预付款汇总查询数据>>\n")
        t16 = threading.Thread(target=yufukuanhuizong)
        t16.setDaemon = True
        t16.start()
        Text.see("end")


def button_end():
    os.system('taskkill /im chrome.exe /F')
    os.system('taskkill /im chromedriver.exe /F')
    os.system('taskkill /im webdriver.exe /F')

    Text.insert("end", "已关闭所有线程\n")

    Text.see("end")


var1 = IntVar()
var2 = IntVar()
var3 = IntVar()
var4 = IntVar()
var5 = IntVar()
var6 = IntVar()
#var8 = IntVar()
var10 = IntVar()
var11 = IntVar()
var13 = IntVar()
var14 = IntVar()
var15 = IntVar()
var16 = IntVar()
var17 = IntVar()

Checkbutton1 = Checkbutton(text='1.商品维护', variable=var1, onvalue=1, offvalue=0, bg='white')
Checkbutton2 = Checkbutton(text='2.采购订单信息', variable=var2, onvalue=1, offvalue=0, bg='white')
Checkbutton3 = Checkbutton(text='3.进出仓明细', variable=var3, onvalue=1, offvalue=0, bg='white')
Checkbutton4 = Checkbutton(text='4.各仓库在途库存数据', variable=var4, onvalue=1, offvalue=0, bg='white')
Checkbutton5 = Checkbutton(text='5.商品主题分析', variable=var5, onvalue=1, offvalue=0, bg='white')
Checkbutton6 = Checkbutton(text='6.未发货订单', variable=var6, onvalue=1, offvalue=0, bg='white')
#Checkbutton8 = Checkbutton(text='8.商品及库存管理', variable=var8, onvalue=1, offvalue=0, bg='white')
Checkbutton10 = Checkbutton(text='10.箱及仓位库存(卓尚分仓)', variable=var10, onvalue=1, offvalue=0, bg='white')
Checkbutton11 = Checkbutton(text='11.销售明细分析(财务)', variable=var11, onvalue=1, offvalue=0, bg='white')
Checkbutton13 = Checkbutton(text='13.箱及仓位库存(本仓)', variable=var13, onvalue=1, offvalue=0, bg='white')
Checkbutton14 = Checkbutton(text='14.(财务结算)入库汇总查询', variable=var14, onvalue=1, offvalue=0, bg='white')
Checkbutton15 = Checkbutton(text='15.(财务结算)应付付款汇总查询', variable=var15, onvalue=1, offvalue=0, bg='white')
Checkbutton16 = Checkbutton(text='16.(财务结算)预付款汇总查询', variable=var16, onvalue=1, offvalue=0, bg='white')

Checkbutton1.grid(row=1, column=0, sticky=W, padx=30, pady=2)
Checkbutton2.grid(row=2, column=0, sticky=W, padx=30, pady=2)
Checkbutton3.grid(row=3, column=0, sticky=W, padx=30, pady=2)
Checkbutton4.grid(row=4, column=0, sticky=W, padx=30, pady=2)
Checkbutton5.grid(row=5, column=0, sticky=W, padx=30, pady=2)
Checkbutton6.grid(row=6, column=0, sticky=W, padx=30, pady=2)
#Checkbutton8.grid(row=7, column=0, sticky=W, padx=30, pady=2)
Checkbutton10.grid(row=7, column=0, sticky=W, padx=30, pady=2)
Checkbutton11.grid(row=8, column=0, sticky=W, padx=30, pady=2)
Checkbutton13.grid(row=9, column=0, sticky=W, padx=30, pady=2)
Checkbutton14.grid(row=10, column=0, sticky=W, padx=30, pady=2)
Checkbutton15.grid(row=11, column=0, sticky=W, padx=30, pady=2)
Checkbutton16.grid(row=12, column=0, sticky=W, padx=30, pady=2)


def allselect():
    var1.set(1)
    var2.set(1)
    var3.set(1)
    var4.set(1)
    var5.set(1)
    var6.set(1)
    var13.set(1)
    var11.set(1)
    var10.set(1)
    var14.set(1)
    var15.set(1)
    var16.set(1)
    pass
    # 全选


def alldeselect():
    var1.set(0)
    var2.set(0)
    var3.set(0)
    var4.set(0)
    var5.set(0)
    var6.set(0)
    var13.set(0)
    var11.set(0)
    var10.set(0)
    var14.set(0)
    var15.set(0)
    var16.set(0)
    # 全选


def portion():
    var1.set(1)
    var2.set(1)
    var3.set(1)
    var4.set(1)
    var5.set(1)
    var6.set(1)
    pass


def portion_two():
    var14.set(1)
    var15.set(1)
    var16.set(1)
    pass


def portion_three():
    var4.set(1)
    var5.set(1)
    var6.set(1)
    pass


def portion_four():
    var4.set(1)
    var1.set(1)
    pass


button_select = Button(root, command=allselect)
button_select['text'] = "全选"
button_select.grid(row=1, column=1, sticky=W)

button_select = Button(root, command=alldeselect)
button_select['text'] = "全不选"
button_select.grid(row=3, column=1, sticky=W)

# 部分选择
button_select = Button(root, command=portion)
button_select['text'] = "选择1~6"
button_select.grid(row=3, column=3, sticky=W)

# 部分选择
button_select = Button(root, command=portion_two)
button_select['text'] = "选择14~16"
button_select.grid(row=5, column=3, sticky=W)

# 部分选择
button_select = Button(root, command=portion_three)
button_select['text'] = "下单(4~6)"
button_select.grid(row=7, column=3, sticky=W)

# 部分选择
button_select = Button(root, command=portion_four)
button_select['text'] = "对库存(1和4)"
button_select.grid(row=9, column=3, sticky=W)

button_start = Button(root, command=button_start)
button_start['text'] = "导出"
button_start.grid(row=5, column=1, sticky=W)

button_end = Button(root, command=button_end)
button_end['text'] = "结束"
button_end.grid(row=7, column=1, sticky=W)

button_select = Button(root, command=check)
button_select['text'] = "更新cookie"
button_select.grid(row=1, column=3, sticky=W)

Checkbutton11 = Checkbutton(text='无窗口运行', variable=var17, onvalue=1, offvalue=0, bg='white')
Checkbutton11.grid(row=1, column=2, sticky=W)
# Checkbutton1.grid(row=1,column=0,sticky=W,padx=30, pady=2)


Text = Text(root)
Text.insert("end", "选择你要导出的数据,并点击导出\n")
Text['height'] = 18
Text['width'] = 69

Text.grid(row=13, columnspan=5, sticky=W, padx=5, pady=5)

try:
    timestamp_previous = load_variavle('timestamp.zs')
    results = load_variavle('variable.zs')
    Text.insert("end", "检测登录状态: 有效\n")
except:
    timestamp_previous = 0

timestamp_now = int(time.time())

if timestamp_now - timestamp_previous > 259200:
    # Text.insert("end", var1.get())
    # print_che()
    Text.insert("end", "检测登录状态: 失效!\n正在更新登录状态\n")
    t1 = threading.Thread(target=checkcookie)
    t1.setDaemon = True
    t1.start()
    t1.join()

try:
    savedCookies = load_variavle('variable.zs')
except:
    Text.insert("end", "检测登录状态: 失效!\n正在更新登录状态\n")
    t1 = threading.Thread(target=checkcookie)
    t1.setDaemon = True
    t1.start()
    t1.join()

print("111")

root.mainloop()
