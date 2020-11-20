# -*- coding: utf-8 -*-
#---------------------------------------*
# Description: 保单成本资金拨付
# Author: zcw
# Date:   2020-11-16
# Name:   PolicyCostAllocation.py
#---------------------------------------*
import os
import time
import pyscreeze
import pyautogui
import datetime as dt
import xlwings as xw
import uiautomation
import win32api
import win32con
import win32gui
import win32clipboard as wc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 盘符
PATH = 'E:\\'
# 临时文件路径
TEMP_PATH = r'E:\flow1_download'

class PolicyCostAllocation(object):

    def __init__(self):
        ...

    @property
    def formula(self):
        return self.formulaTxt

    @formula.setter
    def formula(self, formulaTxt):
        self.formulaTxt = formulaTxt

    # 获取取数日期
    def _get_access_date(self):
        """
            :return: beginDate is yesterday, endDate is yesterday of last week,
                efd is the day before endDate
        """
        beginDate = dt.datetime.today() + dt.timedelta(days=-1)
        endDate = dt.datetime.today() + dt.timedelta(days=-8)
        bd = str(beginDate.year) + '-' + str(beginDate.month) + '-' + str(beginDate.day)
        ed = str(endDate.year) + '-' + str(endDate.month) + '-' + str(endDate.day)
        efd = str(endDate.year) + '-' + str(endDate.month) + '-' + str(endDate.day - 1)
        return bd, ed, efd

    # 导出新车险销售费用数据表
    def _cx_table_export(self, driver):
        while not self._waitelement_by_id(driver, 'treeDemo_293_a', 10): ...
        # 选择保单成本报表
        driver.execute_script("document.getElementById('treeDemo_293_a').click();")
        time.sleep(1)
        # 选择车险保单成本
        driver.execute_script("document.getElementById('treeDemo_294_a').click();")
        time.sleep(1)
        # 选择新车险销售费用报表
        driver.execute_script("document.getElementById('treeDemo_295_a').click();")
        time.sleep(2)
        iframe = driver.find_element_by_xpath('//div[@class="panel panel-htop easyui-fluid"]//iframe')
        driver.switch_to.frame(iframe)
        self.formula = "to_char([保单成本].[新车险销售费用数据清单].[{}],'yy/mm/dd')"
        # 生成表格模板
        self._builder_table_module(driver, iframe, 'hal__dom__uniqueID__56', 'last_cloumn_1',\
            ['险种代码', '客户类别3', '核保日期', '费用确认日期', '保费', '手续费金额', '展业费金额', '业绩提奖金额',\
                '技术服务费金额', '财务资源金额', '总费用金额'], {'核保日期': 3, '费用确认日期': 4})
        # 点击下载 excel
        driver.switch_to.parent_frame()
        while not self._waitelement_by_id(driver, 'com.ibm.bi.authoring.runMenuPluginContainer', 5): ...
        driver.find_element_by_id('com.ibm.bi.authoring.runMenuPluginContainer').click()
        while not self._waitelement_by_id(driver, 'view100_item103', 5): ...
        driver.find_element_by_id('view100_item103').click()
        # 设置报表时间数据范围
        while pyscreeze.locateOnScreen(r'D:\flow_1_resource\xinche.png', minSearchTime=20) is None: ...
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[1])
        while not self._waitelement_by_id(driver, 'dv22__tblDateTextBox__txtInput', 5): ...
        bd, ed, efd = self._get_access_date()
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
        driver.find_element_by_id('dv76__tblDateTextBox__txtInput').send_keys(ed)
        # 点击完成自动执行下载
        finishBtn = driver.find_element_by_xpath('//button[text()= "完成"]')
        driver.execute_script("arguments[0].scrollIntoView();", finishBtn)
        finishBtn.click()
        fileDir = r'D:\{}年保单成本额度表'.format(dt.datetime.now().year)
        fileName = r'新车险销售费用数据清单(费用确认时间{}).xlsx'.format(\
            bd[5:].replace('-', '.') + '-' + efd[5:].replace('-', '.'))
        return self._download_table(fileDir, fileName)

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
        self.formula = "to_char([财产险保单成本].[新财产险销售费用清单].[{}],'yy/mm/dd')"
        # 生成表格模板
        self._builder_table_module(driver, iframe, 'hal__dom__uniqueID__50', 'last_cloumn_2',\
            ['险种代码', '保单号', '核保日期', '实收日期', '保费', '手续费金额', '展业费金额', '业绩提奖金额',\
                '技术服务费金额', '财务资源金额', '总费用金额'], {'核保日期': 3, '实收日期': 4})
        # 点击下载 excel
        driver.switch_to.parent_frame()
        while not self._waitelement_by_id(driver, 'com.ibm.bi.authoring.runMenuPluginContainer', 5): ...
        driver.find_element_by_id('com.ibm.bi.authoring.runMenuPluginContainer').click()
        while not self._waitelement_by_id(driver, 'view100_item103', 5): ...
        driver.find_element_by_id('view100_item103').click()
        # 设置报表时间数据范围
        while pyscreeze.locateOnScreen(r'D:\flow_1_resource\xincaichan.png', minSearchTime=20) is None: ...
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[1])
        while not self._waitelement_by_id(driver, 'dv22__tblDateTextBox__txtInput', 5): ...
        bd, ed, efd = self._get_access_date()
        driver.find_element_by_id('dv22__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv26__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv32__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv36__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv42__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv42__tblDateTextBox__txtInput').send_keys(bd)
        driver.find_element_by_id('dv46__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv46__tblDateTextBox__txtInput').send_keys(ed)
        driver.find_element_by_id('dv52__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv56__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv62__tblDateTextBox__txtInput').clear()
        driver.find_element_by_id('dv66__tblDateTextBox__txtInput').clear()
        # 点击完成自动执行下载
        finishBtn = driver.find_element_by_xpath('//button[text()= "完成"]')
        driver.execute_script("arguments[0].scrollIntoView();", finishBtn)
        finishBtn.click()
        fileDir = r'D:\{}年保单成本额度表'.format(dt.datetime.now().year)
        fileName = r'新财产险销售费用清单(实收日期{}).xlsx'.format( \
            bd[5:].replace('-', '.') + '-' + efd[5:].replace('-', '.'))
        return self._download_table(fileDir, fileName)

    # 表格处理
    def _excel_dispose(self, app, file_path, table_type):
        """
            :param app: WPS 进程实例
            :param file_path: 表格绝对路径
            :param table_type: 表格数据类型【1.新车险费用数据表  2.新财产险费用数据表  3. 4. 5.】
            :return: None
        """
        # 等待表格下载完毕
        while not os.path.exists(file_path): time.sleep(1)
        # 根据报表对应的类型进行处理
        if table_type in [1, 2]:
            wb = app.books.open(file_path)
            sh_1 = wb.sheets['页面1_1']
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
            sh_2 = wb.sheets['Sheet1']
            sh_2.api.activate
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
            sh_2.range('B1:C1').api.select
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
            if table_type == 1:
                self._add_screen_field('xzdm')
                self._add_screen_field('khlb')
                self._add_screen_field('hbrq')
                self._add_screen_field('fyqrrq')
            else:
                self._add_screen_field('xzdm')
                self._add_screen_field('hbrq')
                self._add_screen_field('ssrq')
                # 对核保日期进行筛选
                sh_2.range('B2').api.select
                self._hot_keys(1, 'alt', 'down')
                self._press('tab', 0.1, 3)
                # 勾选更多选项
                self._press('space', 1)
                # 切换到日期选择框
                self._press('tab', 0.1, 4)
                self._date_filtering(2019, 10, 1)
        elif table_type == 3:
            ...
        elif table_type == 4:
            ...
        elif table_type == 5:
            ...
        # 保存退出
        wb.save()
        app.quit()

    # 构建表格模板
    def _builder_table_module(self, driver, parent_Frame, data_Item_Id, last_field_name, field_name_list, field_map):
        # 进入清单编辑页面
        while not self._waitelement_by_id(driver, 'com.ibm.bi.classicviewer.editBtn', 10): ...
        driver.find_element_by_id('com.ibm.bi.classicviewer.editBtn').click()
        # 等待错误弹窗并关闭
        while not self._waitelement_by_id(driver, 'ok', 10): ...
        self._press('enter', 2)
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[1])
        time.sleep(2)
        while not self._waitelement_by_id(driver, 'idLayoutView', 10): ...
        # 获取清单第首行字段集合
        td_list = driver.find_element_by_id("idLayoutView").find_elements_by_tag_name('tbody')[3].\
            find_elements_by_tag_name('tr')[0].find_elements_by_tag_name('td')
        # 删除选中范围字段
        td_list[1].click()
        # 页面移动到屏幕最右边
        self._press('right', 0.1, 5)
        time.sleep(1)
        # 定位到要删除的最后一段元素
        pyautogui.keyDown('shift')
        self._web_element_location(last_field_name, 5, 10)
        pyautogui.keyUp('shift')
        driver.find_element_by_id('btnMore').click()
        time.sleep(1)
        while not self._waitelement_by_id(driver, 'mnuLayoutDelete', 10): ...
        driver.find_element_by_id('mnuLayoutDelete').find_elements_by_tag_name('div')[0].click()
        time.sleep(2)
        # 等待删除完毕
        while True:
            td_list = driver.find_element_by_id("idLayoutView").find_elements_by_tag_name('tbody')[3]. \
                find_elements_by_tag_name('tr')[0].find_elements_by_tag_name('td')
            if len(td_list) == 1: break
            time.sleep(2)
        driver.switch_to.default_content()
        time.sleep(2)
        # 进入数据清单目录
        driver.execute_script(
            "document.getElementsByClassName('panel panel-htop easyui-fluid')[0].getElementsByTagName('iframe')[0].contentDocument.getElementById('com.ibm.bi.authoring.treeBtn').click();")
        time.sleep(1)
        driver.switch_to.frame(parent_Frame)
        time.sleep(1)
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[1])
        while not self._waitelement_by_id(driver, data_Item_Id, 10): ...
        # 展开销售清单列表
        driver.find_element_by_id(data_Item_Id).find_elements_by_tag_name('img')[1].click()
        time.sleep(1)
        # 生成新的清单数据表
        data_list_element = driver.find_element_by_id(data_Item_Id).find_elements_by_tag_name('div')[0]
        # 添加字段
        self._add_field(driver, data_list_element, field_name_list)
        # 设置公式
        self._data_style_edit(driver, field_map)

    # 表格透视列选项定位
    def _add_screen_field(self, file_name):
        box = pyscreeze.locateOnScreen(r'{}flow_1_resource\{}.png'.format(PATH, file_name), minSearchTime=10)
        pyautogui.rightClick(box.left + 10, box.top + 5)
        self._press('down', 1)
        self._press('enter', 1)

    # 网页元素坐标定位
    def _web_element_location(self, file_name, xOffset, yOffset, click_count = 1):
        box = pyscreeze.locateOnScreen(r'D:\flow_1_resource\{}.png'.format(file_name), minSearchTime=10)
        pyautogui.click(x = box.left + xOffset, y = box.top + yOffset, clicks = click_count)

    # 添加字段
    def _add_field(self, driver, data_list_element, field_name_list):
        for t in field_name_list:
            # 滚动至要添加到清单模板中的字段
            driver.execute_script(
                "arguments[0].scrollIntoView();", \
                data_list_element.find_element_by_xpath('//span[text()="{}"]'.format(t)))
            time.sleep(1)
            # 为避免图像识别失败，将鼠标移动到其他区别避免覆盖图片
            pyautogui.moveTo(900, 400)
            time.sleep(1)
            # 添加字段到模板中
            self._web_element_location(t, 30, 10, 2)
            time.sleep(1)

    # 数据样式编辑
    def _data_style_edit(self, driver, fieldMap):
        for k in fieldMap:
            # 获取新创建的模板清单字段列表, 每成功设置一个字段公式后该元素会刷新
            td_list = \
                driver.find_element_by_id("idLayoutView").find_elements_by_tag_name('tbody')[3].find_elements_by_tag_name(\
                    'tr')[0].find_elements_by_tag_name('td')
            # 设置查询公式
            self._set_query_formula(driver, td_list[fieldMap.get(k)], self.formula.format(k))
            time.sleep(2)

    # 设置公式
    def _set_query_formula(self, driver, select_ele, formula_str):
        # 选中要设置公式的字段
        select_ele.click()
        time.sleep(1)
        # 展开更多按钮
        driver.find_element_by_id('btnMore').click()
        while not self._waitelement_by_id(driver, 'mnuDoEditQueryExpression', 5): ...
        # 进入表达式编辑窗口
        driver.find_element_by_id('mnuDoEditQueryExpression').find_elements_by_tag_name('div')[0].click()
        time.sleep(1)
        # 设置查询公式
        driver.find_element_by_id('_5GF_taText').clear()
        time.sleep(1)
        driver.find_element_by_id('_5GF_taText').send_keys(formula_str)
        time.sleep(1)
        driver.find_element_by_id('_5GF_btnOK').click()

    # 日期过滤
    def _date_filtering(self, year, month, day):
        threshold = dt.date(year, month, day)
        year_prefix = '{}'.format(dt.datetime.now().year)[0:2]
        while True:
            self._press('down')
            self._hot_keys('ctrl', 'c')
            time.sleep(0.5)
            wc.OpenClipboard()
            dateValue = wc.GetClipboardData(win32con.CF_UNICODETEXT)
            if dt.datetime.strptime(year_prefix + dateValue, '%Y/%m/%d').date() < threshold:
                pyautogui.press('space')
                time.sleep(0.5)
                wc.CloseClipboard()
            else:
                time.sleep(0.5)
                self._press(key_code='tab', presses=2)
                self._press('enter')
                break

    # 快捷键
    def _hot_keys(self, interval = 0, *args):
        pyautogui.hotkey(*args)
        if interval > 0: time.sleep(interval)

    # 按键输入
    def _press(self, key_code, interval = 0, presses = 1):
        pyautogui.press(keys=key_code,presses=presses)
        if interval > 0: time.sleep(interval)

    # 表格下载
    def _download_table(self, fileDir, fileName):
        self._window_waiting('另存为')
        pyautogui.press('backspace')
        save = win32gui.FindWindow('#32770', '另存为')
        if not os.path.exists(fileDir): os.makedirs(fileDir)
        edit = uiautomation.ControlFromHandle(save).EditControl(searchDepth=10, Name='文件名:').NativeWindowHandle
        filePath = fileDir + os.path.sep + fileName
        win32api.SendMessage(edit, win32con.WM_SETTEXT, None, filePath)
        pyautogui.hotkey('alt', 's')
        return filePath

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
        # 等待主页面加载完毕
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