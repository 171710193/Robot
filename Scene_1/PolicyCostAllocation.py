# -*- coding: utf-8 -*-
#---------------------------------------*
# Description: 保单成本资金拨付
# Author: zcw
# Date:   2020-11-16
# Name:   PolicyCostAllocation.py
#---------------------------------------*
import os
import time
import datetime as dt
import xlwings as xw
import pyscreeze
import pyautogui
import uiautomation
import win32api
import win32con
import win32gui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 盘符
PATH = 'E:\\'
# 临时文件路径
TEMP_PATH = r'E:\flow1_download'

class PolicyCostAllocation:

    def __init__(self):
        ...

    # 获取取数日期
    def _get_last_week_date(self):
        """
            :return: beginDate is yesterday, lastDate is yesterday of last week
        """
        beginDate = dt.datetime.today() + dt.timedelta(days=-1)
        lastDate = dt.datetime.today() + dt.timedelta(days=-8)
        return beginDate, lastDate

    # 导出新车险销售费用数据表
    def _cx_table_export(self, driver):
        while not self._waitelement_by_id(driver, 'treeDemo_293_a', 10): ...
        # 选择
        driver.execute_script("document.getElementById('treeDemo_293_a').click();")
        time.sleep(1)
        driver.execute_script("document.getElementById('treeDemo_294_a').click();")
        time.sleep(1)
        driver.execute_script("document.getElementById('treeDemo_295_a').click();")
        time.sleep(2)
        iframe = driver.find_element_by_xpath('//div[@class="panel panel-htop easyui-fluid"]//iframe')
        driver.switch_to.frame(iframe)
        driver = self._list_field_remove_by_range(driver, 'last_cloumn_1')
        time.sleep(1)
        driver.switch_to.default_content()
        time.sleep(2)
        # 进入数据清单目录
        driver.execute_script("document.getElementsByClassName('panel panel-htop easyui-fluid')[0].getElementsByTagName('iframe')[0]\
            .contentDocument.getElementById('com.ibm.bi.authoring.treeBtn').click();")
        time.sleep(1)
        # 切换到保单成本所在的 iframe 中
        driver.switch_to.frame(iframe)
        time.sleep(1)
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[1])
        time.sleep(1)
        while not self._waitelement_by_id(driver, 'hal__dom__uniqueID__56', 10): ...
        # 展开销售清单列表
        driver.find_element_by_id("hal__dom__uniqueID__56").find_elements_by_tag_name('img')[1].click()
        time.sleep(1)
        # 生成新的清单数据表
        data_list_ele = driver.find_element_by_id('hal__dom__uniqueID__56').find_elements_by_tag_name('div')[0]
        # 险种代码
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__394', 'xz_code')
        # 客户类别
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__438', 'client_type')
        # 核保日期
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__399', 'underwriting')
        # 费用确认日期
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__486', 'expense_confirmation')
        # 保费
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__404', 'premium')
        # 手续费金额
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__461', 'sxf')
        # 展业费金额
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__463', 'zyf')
        # 业绩提奖金额
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__465', 'yjte')
        # 技术服务费金额
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__471', 'jsfw')
        # 财务资源金额
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__467', 'cwzy')
        # 总费用金额
        self._create_field(driver, data_list_ele, 'hal__dom__uniqueID__459', 'zf')
        # 设置公式
        self._data_style_edit(driver)
        # 点击下载 excel
        driver.switch_to.parent_frame()
        time.sleep(1)
        driver.find_element_by_id('com.ibm.bi.authoring.runMenuPluginContainer').click()
        while not self._waitelement_by_id(driver, 'view100_item103', 5): ...
        time.sleep(1)
        driver.find_element_by_id('view100_item103').click()
        # 等待进入设置页面
        while pyscreeze.locateOnScreen(r'D:\flow_1_resource\newcar.png', minSearchTime=20) is None: ...
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[1])
        while not self._waitelement_by_id(driver, 'dv22__tblDateTextBox__txtInput', 5): ...
        beginDate, endDate = self._get_last_week_date()
        # 时间筛选
        bd = str(beginDate.year) +'-'+ str(beginDate.month) +'-'+ str(beginDate.day)
        ed = str(endDate.year) +'-'+ str(endDate.month) +'-'+ str(endDate.day)
        driver.find_element_by_id('dv22__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv26__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv32__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv36__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv42__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv46__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv52__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv56__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv62__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv66__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv72__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv72__tblDateTextBox__txtInput').send_keys(bd)
        driver.find_element_by_id('dv76__tblDateTextBox__txtInput').clear()
        finish_btn = driver.find_element_by_xpath('//button[text()= "完成"]')
        driver.execute_script("arguments[0].scrollIntoView();", finish_btn)
        finish_btn.click()
        self._window_waiting()
        pyautogui.press('backspace')
        save = win32gui.FindWindow('#32770', '导出为WPS PDF')
        file_path = r'D:\新车险销售费用数据清单(费用确认时间{}).xlsx'.format(str(beginDate)[5:].replace('-', '.') +'-'+ str(endDate)[5:].replace('-', '.'))
        if not os.path.exists(file_path): os.makedirs(file_path)
        edit = uiautomation.ControlFromHandle(save).EditControl(searchDepth=10, Name='文件名:').NativeWindowHandle
        win32api.SendMessage(edit, win32con.WM_SETTEXT, None, file_path)
        pyautogui.hotkey('alt', 's')
        return file_path

    # 导出新财产险销售费用数据表
    def _ccx_table_export(self, driver):
        driver.switch_to.default_content()
        # 选择财产险保单成本
        driver.execute_script("document.getElementById('treeDemo_302_a').click()")
        time.sleep(1)
        # 选择新财产保险销售费用清单
        driver.execute_script("document.getElementById('treeDemo_303_a').click();")
        time.sleep(2)
        iframe = driver.find_element_by_xpath('//div[@class="panel panel-htop easyui-fluid"]//iframe')
        driver.switch_to.frame(iframe)
        self._list_field_remove_by_range(driver, 'last_cloumn_2')





    # 表格处理
    def _excel_dispose(self, app, file_path, table_type):
        """
            :param app: WPS 进程实例
            :param file_path: 表格路径
            :param table_type: 表格类型
            【1.新车险费用数据表  2.新财产险费用数据表 3. 4. 5.】
            :return: None
        """
        if table_type == 1:
            wb1 = app.books.open(file_path)
            sh_1 = wb1.sheets['页面1_1']
            sh_1.api.activate
            # 设置除数
            sh_1.range('O7').value = 10000
            sh_1.range('O7').api.copy
            end_row_index = 0
            # 选取到汇总行
            for i in range(10, sh_1.used_range.rows.count):
                if sh_1.range('B{}'.format(i)).value == '整体 - 汇总':
                    end_row_index = i
                    break
            # 选定需要重新计算的列范围
            sh_1.range('G10:M{}'.format(end_row_index)).api.select
            time.sleep(1)
            # 选择性粘贴
            self._hot_keys(1, 'shift', 'f10')
            self._press('s', 1)
            self._press('v', 1)
            self._press('i', 1)
            self._press('enter', 1)
            # 设置单元格格式
            self._hot_keys(1, 'ctrl', '1')
            self._press('c', 1)
            self._press('down', 0.5, 3)
            self._press('tab', 0.5, 2)
            self._press('up', 1)
            self._press('enter', 1)
            # 创建数据透视图
            sh_1.range('B9:M{}'.format(end_row_index - 1)).api.copy
            self._hot_keys(1, 'alt', 'd', 'p')
            self._press('enter', 1)
            # 分公司
            self._web_element_location('fgs', 5, 5, 1)
            time.sleep(1)
            # 保费
            self._web_element_location('bf', 5, 5, 1)
            time.sleep(1)
            pyautogui.scroll(-500)
            # 总费用金额
            self._web_element_location('zfy', 5, 5, 1)
            pyautogui.click(x=1310, y=280, clicks=1)
            pyautogui.scroll(500)
            # 设置单元格格式
            sh_1.range('B1:C1').api.select
            pyautogui.hotkey('ctrl', 'shift', 'down')
            pyautogui.hotkey('ctrl', 'shift', 'down')
            pyautogui.hotkey('ctrl', 'shift', 'down')
            self._hot_keys(1, 'shift', 'f10')
            self._press('f', 1)
            self._press('c', 1)
            self._press('down', 0.5, 3)
            self._press('tab', 0.5, 2)
            self._press('up', 1)
            self._press('enter', 1)
            pyautogui.click(x=230, y=400, clicks=1)
            time.sleep(1)
            # 添加字段带筛选区
            self._add_screen_field('xzdm')
            self._add_screen_field('khlb')
            self._add_screen_field('hbrq')
            self._add_screen_field('fyqrrq')
            self._hot_keys(1, 'ctrl', 's')
        elif table_type == 2:
            ...
        elif table_type == 3: #
            ...
        elif table_type == 4: #
            ...
        elif table_type ==5: #
            ...

    # 清单字段范围选取删除
    def _list_field_remove_by_range(self, driver, last_field_name):
        # 进入到清单编辑页面
        while not self._waitelement_by_id(driver, 'com.ibm.bi.classicviewer.editBtn', 10): ...
        driver.find_element_by_id('com.ibm.bi.classicviewer.editBtn').click()
        # 等待弹窗出现
        while not self._waitelement_by_id(driver, 'ok', 10): ...
        self._press('enter', 2)
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[1])
        time.sleep(2)
        while not self._waitelement_by_id(driver, 'idLayoutView', 10): ...
        # 获取清单第一行字段集合
        td_list = driver.find_element_by_id("idLayoutView").find_elements_by_tag_name('tbody')[3].\
            find_elements_by_tag_name('tr')[0].find_elements_by_tag_name('td')
        # 删除选中字段
        td_list[1].click()
        pyautogui.keyDown('shift')
        time.sleep(1)
        # 定位到要删除的最后一段元素
        self._web_element_location(last_field_name, 5, 10)
        pyautogui.keyUp('shift')
        driver.find_element_by_id('btnMore').click()
        time.sleep(1)
        while not self._waitelement_by_id(driver, 'mnuLayoutDelete', 10): ...
        driver.find_element_by_id('mnuLayoutDelete').find_elements_by_tag_name('div')[0].click()
        return driver

    # 表格透视列选项定位
    def _add_screen_field(self, file_name):
        box = pyscreeze.locateOnScreen(r'{}flow_1_resource\{}.png'.format(PATH, file_name), minSearchTime=10)
        pyautogui.rightClick(box.left + 10, box.top + 5)
        self._hot_keys(1, 'shift', 'f10')
        self._press('down', 1)
        self._press('enter', 1)

    # 网页元素坐标定位
    def _web_element_location(self, file_name, xOffset, yOffset, click_count=1):
        box = pyscreeze.locateOnScreen(r'D:\flow_1_resource\{}.png'.format(file_name), minSearchTime=10)
        pyautogui.click(x = box.left + xOffset, y = box.top + yOffset, clicks = click_count)

    # 构建模板
    def _create_field(self, driver, father_ele, ele_id, file_name):
        # 滚动至要添加到清单模板中的字段
        driver.execute_script("arguments[0].scrollIntoView();", father_ele.find_element_by_id(ele_id))
        time.sleep(1)
        # 为避免图像识别失败，将鼠标移动到其他区别避免覆盖图片
        pyautogui.moveTo(900, 400)
        time.sleep(1)
        self._web_element_location(file_name, 30, 10, 2)
        time.sleep(1)

    # 数据样式编辑
    def _data_style_edit(self, driver):
        # 获取新创建的模板清单字段列表
        td_list = \
            driver.find_element_by_id("idLayoutView").find_elements_by_tag_name('tbody')[3].find_elements_by_tag_name(
                'tr')[0].find_elements_by_tag_name('td')
        # 设置核保日期公式
        self._set_query_formula(driver, td_list[3], "to_char([保单成本].[新车险销售费用数据清单].[核保日期],'yy/mm/dd')")
        td_list = \
            driver.find_element_by_id("idLayoutView").find_elements_by_tag_name('tbody')[3].find_elements_by_tag_name(
                'tr')[0].find_elements_by_tag_name('td')
        # 设置费用确认日期
        self._set_query_formula(driver, td_list[4], "to_char([保单成本].[新车险销售费用数据清单].[费用确认日期],'yy/mm/dd')")

    # 设置公式
    def _set_query_formula(self, driver, select_ele, formula_str):
        # 选中要设置公式的字段
        select_ele.click()
        time.sleep(1)
        # 展开工具栏更多按钮
        driver.find_element_by_id('btnMore').click()
        time.sleep(1)
        # 进入表达式编辑窗口
        driver.find_element_by_id('mnuLayoutPopup').find_element_by_id('hal__dom__uniqueID__234').click()
        time.sleep(1)
        # 设置查询公式
        driver.find_element_by_id('_5GF_taText').clear()
        time.sleep(1)
        driver.find_element_by_id('_5GF_taText').send_keys(formula_str)
        time.sleep(1)
        driver.find_element_by_id('_5GF_btnOK').click()
        time.sleep(1)

    # 快捷键
    def _hot_keys(self, interval, *args):
        pyautogui.hotkey(*args)
        if interval > 0: time.sleep(interval)

    # 按键输入
    def _press(self, key_code, interval, presses = 1):
        pyautogui.press(keys=key_code,presses=presses)
        if interval > 0: time.sleep(interval)

    # 窗口等待
    def _window_waiting(self, title):
        while win32gui.FindWindow('#32770', title) == 0: time.sleep(1)

    # 根据元素ID等待元素
    def _waitelement_by_id(self, driver, id, seconds):
        try:
            WebDriverWait(driver, seconds).until(EC.presence_of_all_elements_located((By.ID, id)))
        except Exception:
            return False
        return True

    # 登入
    def _login(self, robot):
        chromeOptions = webdriver.ChromeOptions()
        prefs = {"download.default_directory": TEMP_PATH}
        chromeOptions.add_experimental_option("prefs", prefs)
        chromeOptions.add_argument('--args --disable-web-security --user-data-dir')
        driver = webdriver.Chrome(options=chromeOptions)
        driver.get('http://mis.sinosafe.com.cn/pages/index.html')
        driver.maximize_window()
        time.sleep(2)
        # 输入账号
        driver.execute_script("document.getElementById('username').value = 'liuxuanchen@sinosafe.com.cn';")
        time.sleep(1)
        # 输入密码
        driver.execute_script("document.getElementById('password').value = 'Habx-8888';")
        time.sleep(1)
        # 登入
        driver.execute_script("document.getElementsByName('submit')[0].click();")
        while not robot._waitelement_by_id(driver, 'loading', 5): ...
        return driver

    # 退出浏览器
    def _quit_browser(self, driver):
        driver.quit()

if __name__ == "__main__":
    try:
        robot = PolicyCostAllocation()
        # 系统登入
        chromDriver = robot._login(robot)
        # 创建 wps进程
        app = xw.App(visible=True, add_book=False)
        # 导出新车险销售费用数据表
        file_path = robot._cx_table_export(chromDriver)
        time.sleep(1)
        # 处理新车销售费用数据
        robot._excel_dispose(app, file_path, 1)
        time.sleep(1)
        # 导出新财产险销售数据费用表
        robot._ccx_table_export(chromDriver)
    except Exception as e:
        print(e)
    finally:
        print('end')
        #robot._quit_browser(driver)