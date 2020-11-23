# -*- coding: utf-8 -*-
#---------------------------------------*
# Description: 场景1 保单成本资金拨付
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
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from xlwings.constants import DeleteShiftDirection

# 文本保存磁盘符
PATH = 'D:\\'

class PolicyCostAllocation(object):

    # This is a global public instance.
    _web_driver = None

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
                efd is the day before endDate.
        """
        beginDate = dt.datetime.today() + dt.timedelta(days=-8)
        endDate = dt.datetime.today() + dt.timedelta(days=-1)
        bd  = str(beginDate.year) + '-' + str(beginDate.month) + '-' + str(beginDate.day)
        ed  = str(endDate.year) + '-' + str(endDate.month) + '-' + str(endDate.day)
        efd = str(endDate.year) + '-' + str(endDate.month) + '-' + str(endDate.day - 1)
        return bd, ed, efd

    # 导出新车险销售数据费用表
    def _xcx_table_export(self, driver):
        """
            # ---------------------------------------#
            | MenuTreeLevel：                         |
            |  →保单成本报表 id: 293                    |
            |   ↓车险保单成本 id: 294                   |
            |     ↓新车险销售费用报表 id: 295            |
            #----------------------------------------#
        """
        self.formula = "to_char([保单成本].[新车险销售费用数据清单].[{}],'yy/mm/dd')"
        self._builder_module(driver, self._menu_level_entry(driver, 293, 294, 295), {'hal__dom__uniqueID__56': None},\
            ['分公司', '险种代码', '客户类别3', '核保日期', '费用确认日期', '保费', '手续费金额', '展业费金额', '业绩提奖金额',\
                '技术服务费金额', '财务资源金额', '总费用金额'], {'核保日期': 3, '费用确认日期': 4})
        self._goto_date_setting_page(driver, 'xinchexian')
        while not self._waitelement_by_id(driver, 'dv22__tblDateTextBox__txtInput', 5): ...
        for id in [22, 26, 32, 36, 42, 46, 52, 56, 62, 66, 72, 76]: driver.find_element_by_id('dv{}__tblDateTextBox__txtInput'.format(id)).clear()
        bd, ed, efd = self._get_access_date()
        driver.find_element_by_id('dv72__tblDateTextBox__txtInput').send_keys(bd)
        driver.find_element_by_id('dv76__tblDateTextBox__txtInput').send_keys(ed)
        fileDir = r'{}{}年保单成本额度表'.format(PATH, dt.datetime.now().year)
        fileName = r'新车险销售费用数据清单(费用确认时间{}).xlsx'.format(bd[5:].replace('-', '.') + '-' + efd[5:].replace('-', '.'))
        return self._download_table(driver, fileDir, fileName)

    # 导出新财产险销售费用数据表
    def _xccx_table_export(self, driver):
        """
            # ---------------------------------------#
            | MenuTreeLevel：                         |
            |  →保单成本报表 id: 293                    |
            |   ↓财产险保单成本 id: 302                 |
            |     ↓新财产保险销售费用清单 id: 303        |
            #----------------------------------------#
        """
        self.formula = "to_char([财产险保单成本].[新财产险销售费用清单].[{}],'yy/mm/dd')"
        self._builder_module(driver, self._menu_level_entry(driver, None, 302, 303), {'hal__dom__uniqueID__50': None},\
            ['分公司', '险种代码', '保单号', '核保日期', '实收日期', '保费', '手续费金额', '展业费金额', '业绩提奖金额',\
                '技术服务费金额', '财务资源金额', '总费用金额'], {'核保日期': 3, '实收日期': 4})
        self._goto_date_setting_page(driver, 'xincaichanxian')
        while not self._waitelement_by_id(driver, 'dv22__tblDateTextBox__txtInput', 5): ...
        for id in [22, 26, 32, 36, 42, 46, 52, 56, 62, 66]: driver.find_element_by_id('dv{}__tblDateTextBox__txtInput'.format(id)).clear()
        bd, ed, efd = self._get_access_date()
        driver.find_element_by_id('dv42__tblDateTextBox__txtInput').send_keys(bd)
        driver.find_element_by_id('dv46__tblDateTextBox__txtInput').send_keys(ed)
        fileDir = r'{}{}年保单成本额度表'.format(PATH, dt.datetime.now().year)
        fileName = r'新财产险销售费用清单(实收日期{}).xlsx'.format(bd[5:].replace('-', '.') + '-' + efd[5:].replace('-', '.'))
        return self._download_table(driver, fileDir, fileName)

    # 导出新人生险销售费用数据清单
    def _xrsx_table_export(self, driver):
        """
            # ---------------------------------------#
            | MenuTreeLevel：                        |
            |  →保单成本报表 id: 293                   |
            |   ↓人生险保单成本 id: 306                |
            |    ↓新人生险销售费用清单 id: 308          |
            #----------------------------------------#
        """
        self._builder_module(driver, self._menu_level_entry(driver, None, 306, 308), {'hal__dom__uniqueID__50', None},\
            ['分公司', '险种代码', '保单号', '核保日期', '实收日期', '保费', '手续费金额', '展业费金额', '业绩提奖金额',\
                '技术服务费金额', '财务资源金额', '总费用金额'])
        self._goto_date_setting_page(driver, 'xinrenshengxian')
        while not self._waitelement_by_id(driver, 'dv22__tblDateTextBox__txtInput', 5): ...
        for id in [22, 26, 32, 36, 42, 46, 52, 56, 62, 66]: driver.find_element_by_id('dv{}__tblDateTextBox__txtInput'.format(id)).clear()
        bd, ed, efd = self._get_access_date()
        driver.find_element_by_id('dv42__tblDateTextBox__txtInput').send_keys(bd)
        driver.find_element_by_id('dv46__tblDateTextBox__txtInput').send_keys(ed)
        fileDir = r'{}{}年保单成本额度表'.format(PATH, dt.datetime.now().year)
        fileName = r'新人生险销售费用清单(实收日期{}).xlsx'.format(bd[5:].replace('-', '.') + '-' + efd[5:].replace('-', '.'))
        return self._download_table(driver, fileDir, fileName)

    # 导出实收保费统计表
    def _ssbf_table_export(self, driver):
        """
            # ---------------------------------------#
            | MenuTreeLevel：                        |
            |  →公共报表 id: 1                        |
            |   ↓收付 id: 34                         |
            |    ↓实收保费统计表 id: 41                |
            #----------------------------------------#
        """
        # 选择模板
        ActionChains(driver).move_to_element(driver.find_element_by_xpath('//span[text()="我的模板"]')).perform()
        while not self._waitelement_by_id(driver, 'mytemplate', 5): ...
        driver.execute_script("document.getElementById('mytemplate').getElementsByTagName('div')[1].click();")
        driver.switch_to.frame(driver.find_element_by_xpath('//div[@class="panel panel-htop easyui-fluid"]//iframe'))
        # 等待模板列表加载完成
        while True:
            try:
                time.sleep(5)
                # 选择06计划财务部
                driver.find_element_by_xpath('//span[text()="06计划财务部"]').click()
                time.sleep(1)
                # 选择保单成本管理
                driver.find_element_by_xpath('//span[text()="保单成本管理"]').click()
                time.sleep(1)
                # 选择实收保费报表模板
                driver.find_element_by_xpath('//span[text()="实收保费统计表--xuanchen模板"]').click()
                break
            except Exception:
                ...
        # 等待日期设置页面加载完成
        while pyscreeze.locateOnScreen(r'D:\flow_1_resource\{}.png'.format('shishoubaofei'), minSearchTime=10) is None: ...
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[0])
        while not self._waitelement_by_id(driver, 'rt_NS_', 5): ...
        date_input_elements = driver.find_element_by_id('rt_NS_').find_elements_by_tag_name('tbody')[2]\
            .find_elements_by_class_name('clsSelectDateEditBox')
        for i in range(len(date_input_elements)): date_input_elements[i].clear()
        bd, ed, efd = self._get_access_date()
        date_input_elements[4].send_keys(bd)
        date_input_elements[5].send_keys(ed)
        d = dt.date(2019, 10, 1)
        date_input_elements[7].send_keys(str(d.year) + '-' + str(d.month) + '-' + str(d.day))
        fileDir = r'{}{}年保单成本额度表'.format(PATH, dt.datetime.now().year)
        fileName = r'实收保费统计表(实收日期{}, 承保确认时间2019.10.01之前).xlsx'.format(bd[5:].replace('-', '.') + '-' + efd[5:].replace('-', '.'))
        return self._download_table(driver, fileDir, fileName, False)

    # 表格处理
    def _excel_dispose(self, file_path, table_type):
        """
            :param file_path: 表格绝对路径
            :param table_type: 表格数据类型
            【1.新车险费用数据表  2.新财产险费用数据表 3.新人生险费用数据表 4.实收保费统计表】
            :return: None
        """
        # 确认表格下载完毕
        while not os.path.exists(file_path): time.sleep(1)
        # 获取表格处理器进程
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open(file_path)
        # 默认激活第一个表单
        sh_1 = wb.sheets['页面1_1']
        sh_1.api.activate
        # 根据报表对应的类型进行处理
        if table_type in [1, 2, 3]:
            if table_type == 3:
                # 删除空白列
                sh_1.range('A:A').delete()
                sh_1.range('M:M').delete()
                # 删除空白行
                sh_1.range('1:8').api.Delete(DeleteShiftDirection.xlShiftUp)
                # 合并其他页签数据
                sh_2 = wb.sheets['页面1_2']
                sh_2.api.activate
                # 复制要合并的数据
                sh_2.range('B10:M{}'.format(self._get_summary_row_index(sh_2))).api.copy
                # 从第一个页签的最后一行的下一行开始粘贴
                sh_1.api.activate
                sh_1.range('A{}'.format(sh_1.used_range.rows.count + 1)).api.select
                self._hot_keys(10, 'ctrl', 'v')
            # 设置除数
            sh_1.range('O7').value = 10000
            sh_1.range('O7').api.copy
            # 获取数据汇总行的索引
            end_row_index = self._get_summary_row_index(sh_1) if table_type != 3 else sh_1.used_range.rows.count
            # 选定需要重新计算的列范围
            style_range = 'G10:M{}' if table_type != 3 else 'F2:L{}'
            sh_1.range(style_range.format(end_row_index)).api.select
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
            self._press('down', 1, 3)
            self._press('tab', 1, 2)
            self._press('up', 1)
            self._press('enter', 1)
            # 创建数据透视图
            perspective_range = 'B9:M{}' if table_type != 3 else 'A1:L{}'
            sh_1.range(perspective_range.format(end_row_index - 1)).api.copy
            self._hot_keys(1, 'alt', 'd', 'p')
            self._press('enter', 1)
            sh_2 = wb.sheets['Sheet1']
            sh_2.api.activate
            # 分公司
            self._web_element_location('fgs', 5, 5)
            # 保费
            self._web_element_location('bf', 5, 5)
            pyautogui.scroll(-500)
            # 总费用金额
            self._web_element_location('zfy', 5, 5)
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
            self._press('down', 1, 3)
            self._press('tab', 1, 2)
            self._press('up', 1)
            self._press('enter', 1)
            pyautogui.click(x=230, y=400, clicks=1)
            time.sleep(1)
            # 添加筛选字段
            if table_type == 1:
                self._add_screen_field(['xzdm', 'khlb', 'hbrq', 'fyqrrq'])
            else:
                self._add_screen_field(['xzdm', 'hbrq', 'ssrq'])
                # 对核保日期进行筛选
                sh_2.range('B2').api.select
                self._hot_keys(1, 'alt', 'down')
                self._press('tab', 1, 3)
                # 勾选更多选项
                self._press('space', 1)
                # 切换到日期选择框
                self._press('tab', 1, 4)
                self._date_filtering(2019, 10, 1, table_type)
        elif table_type == 4:
            ...
        # 保存退出
        wb.save()
        app.quit()
        time.sleep(2)

    # 菜单切入
    def _menu_level_entry(self, driver, level_1, level_2, level_3):
        """
            :param level_1: 一级菜单
            :param level_2: 二级菜单
            :param level_3: 三级菜单
            :return: 页面框架
        """
        if level_1 is not None:
            while not self._waitelement_by_id(driver, 'treeDemo_{}_a'.format(level_1), 10): ...
            driver.execute_script("document.getElementById('treeDemo_{}_a').click();".format(level_1))
            time.sleep(1)
        if self._waitelement_by_id(driver, 'treeDemo_{}_a'.format(level_2), 10):
            driver.execute_script("document.getElementById('treeDemo_{}_a').click();".format(level_2))
            time.sleep(1)
        if self._waitelement_by_id(driver, 'treeDemo_{}_a'.format(level_3), 10):
            driver.execute_script("document.getElementById('treeDemo_{}_a').click();".format(level_3))
            time.sleep(2)
        iframe = driver.find_element_by_xpath('//div[@class="panel panel-htop easyui-fluid"]//iframe')
        driver.switch_to.frame(iframe)
        return iframe

    # 构建表格模板
    def _builder_module(self, driver, parent_Frame, data_Item, field_name_list, field_map = None):
        # 进入清单编辑页面
        while not self._waitelement_by_id(driver, 'com.ibm.bi.classicviewer.editBtn', 10): ...
        driver.find_element_by_id('com.ibm.bi.classicviewer.editBtn').click()
        # 等待错误弹窗并关闭
        while not self._waitelement_by_id(driver, 'ok', 20): ...
        self._press('enter', 2)
        while pyscreeze.locateOnScreen(r'D:\flow_1_resource\blue.png', minSearchTime=30) is None: ...
        time.sleep(1)
        self._press('enter')
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[1])
        while not self._waitelement_by_id(driver, 'idLayoutView', 10): ...
        # 获模板第一行所有元素
        top_line = driver.find_element_by_id("idLayoutView").find_elements_by_tag_name('tbody')[3].\
                find_elements_by_tag_name('tr')[0].find_elements_by_tag_name('td')
        # 选中模板的第一行第一列元素
        top_line[0].click()
        time.sleep(2)
        # 页面移动到第一行最后一列元素可见区域
        driver.execute_script("arguments[0].scrollIntoView();", top_line[-1])
        pyautogui.keyDown('shift')
        # 选中模板第一行最后一列元素
        pyautogui.click(x = 230 + 55 + top_line[-1].location_once_scrolled_into_view.get('x'), y = 230, clicks= 1)
        pyautogui.keyUp('shift')
        self._press('delete', 2)
        # 等待默认模板删除完毕
        while True:
            if driver.find_element_by_id("idLayoutView").find_elements_by_tag_name('tbody')[3]. \
                find_elements_by_tag_name('tr')[2].find_elements_by_tag_name('td')[0].get_attribute('class')\
                    != 'clsTemplateText listColumnBody':
                time.sleep(3)
            else:
                break
        driver.switch_to.default_content()
        time.sleep(2)
        # 进入数据清单目录
        driver.execute_script(
            "document.getElementsByClassName('panel panel-htop easyui-fluid')[0].getElementsByTagName('iframe')[0].contentDocument.getElementById('com.ibm.bi.authoring.treeBtn').click();")
        time.sleep(1)
        driver.switch_to.frame(parent_Frame)
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[1])
        key = list(data_Item.keys())[0]
        while not self._waitelement_by_id(driver, data_Item, 15): ...
        # 展开数据项字段列表
        root_item = driver.find_element_by_id(data_Item)
        root_item.find_elements_by_tag_name('img')[1].click()
        time.sleep(1)
        # 值不为 None 表示该数据项含有子数据项
        if data_Item.get(key) is not None: root_item = root_item.find_element_by_xpath('//span[text()="{}"]/..'.format(data_Item.get(key)))
        # 生成新的模板字段
        self._add_field(driver, driver.find_element_by_id(root_item).find_elements_by_tag_name('div')[0], field_name_list)
        # 设置字段公式
        if field_map is not None: self._data_style_edit(driver, field_map)

    # 进入日期设置页面
    def _goto_date_setting_page(self, driver, page_name):
        driver.switch_to.parent_frame()
        while not self._waitelement_by_id(driver, 'com.ibm.bi.authoring.runMenuPluginContainer', 5): ...
        driver.find_element_by_id('com.ibm.bi.authoring.runMenuPluginContainer').click()
        while not self._waitelement_by_id(driver, 'view100_item103', 5): ...
        driver.find_element_by_id('view100_item103').click()
        while pyscreeze.locateOnScreen(r'D:\flow_1_resource\{}.png'.format(page_name), minSearchTime=60) is None: ...
        driver.switch_to.frame(driver.find_elements_by_tag_name('iframe')[1])

    # 表格透视列选项定位
    def _add_screen_field(self, file_name_list):
        try:
            for file_name in file_name_list:
                box = pyscreeze.locateOnScreen(r'{}flow_1_resource\{}.png'.format(PATH, file_name), minSearchTime=10)
                pyautogui.rightClick(box.left + 10, box.top + 5)
                self._press('down', 1)
                self._press('enter', 1)
        except Exception as e:
            raise Exception('页面中无法定位到素材文件[{}]对应的像素：'.format(file_name) +e)

    # 网页元素坐标定位
    def _web_element_location(self, file_name, xOffset, yOffset, click_count = 1):
        try:
            box = pyscreeze.locateOnScreen(r'D:\flow_1_resource\{}.png'.format(file_name), minSearchTime=10)
            pyautogui.click(x=box.left + xOffset, y=box.top + yOffset, clicks=click_count)
            time.sleep(1)
        except Exception as e:
            raise Exception('页面中无法定位到素材文件[{}]对应的像素：'.format(file_name) +e)

    # 添加字段
    def _add_field(self, driver, data_list_element, field_name_list):
        for t in field_name_list:
            # 滚动至要添加到清单模板中的字段
            driver.execute_script("arguments[0].scrollIntoView();", data_list_element.find_element_by_xpath('//span[text()="{}"]'.format(t)))
            time.sleep(1)
            # 为避免图像识别失败，将鼠标移动到其他区别避免覆盖图片
            pyautogui.moveTo(900, 400)
            time.sleep(1)
            # 添加字段到模板中
            self._web_element_location(t.replace('/', ''), 30, 10, 2)
            time.sleep(1)

    # 数据样式编辑
    def _data_style_edit(self, driver, fieldMap):
        for k in fieldMap:
            # 获取新创建的模板清单字段列表, 每成功设置一个字段公式后该元素会刷新
            td_list = driver.find_element_by_id("idLayoutView").find_elements_by_tag_name('tbody')[3]\
                .find_elements_by_tag_name('tr')[0].find_elements_by_tag_name('td')
            self._set_query_formula(driver, td_list[fieldMap.get(k)], self.formula.format(k))
            time.sleep(2)

    # 设置公式
    def _set_query_formula(self, driver, field_element, formula_str):
        # 选中要设置公式的字段
        field_element.click()
        while not self._waitelement_by_id(driver, 'btnMore', 5): ...
        # 展开更多按钮
        driver.find_element_by_id('btnMore').click()
        while not self._waitelement_by_id(driver, 'mnuDoEditQueryExpression', 5): ...
        # 进入表达式编辑窗口
        driver.find_element_by_id('mnuDoEditQueryExpression').find_elements_by_tag_name('div')[0].click()
        while not self._waitelement_by_id(driver, '_5GF_taText', 5): ...
        # 设置查询公式
        driver.find_element_by_id('_5GF_taText').clear()
        driver.find_element_by_id('_5GF_taText').send_keys(formula_str)
        driver.find_element_by_id('_5GF_btnOK').click()

    # 日期过滤
    def _date_filtering(self, year, month, day, table_type):
        threshold = dt.date(year, month, day)
        prefix = str(dt.datetime.now().year)[0:2] if table_type != 3 else ''
        pattern = '%Y/%m/%d' if table_type != 3 else '%Y-%m-%d'
        while True:
            self._press('down')
            self._hot_keys('ctrl', 'c')
            time.sleep(0.5)
            wc.OpenClipboard()
            dateValue = wc.GetClipboardData(win32con.CF_UNICODETEXT)
            if dt.datetime.strptime(prefix + str(dateValue).split(' ')[0], pattern).date() < threshold:
                self._press('space', 0.5)
                wc.CloseClipboard()
            else:
                self._press(key_code='tab', interval=1, presses=2)
                self._press('enter')
                break

    # 获取数据汇总行索引
    def _get_summary_row_index(self, sheet):
        for i in range(10, sheet.used_range.rows.count):
            if sheet.range('B{}'.format(i)).value == '整体 - 汇总':
                return i

    # 快捷键
    def _hot_keys(self, interval = 0, *args):
        pyautogui.hotkey(*args)
        if interval > 0: time.sleep(interval)

    # 按键输入
    def _press(self, key_code, interval = 0, presses = 1):
        pyautogui.press(keys=key_code, presses=presses)
        if interval > 0: time.sleep(interval)

    # 表格下载
    def _download_table(self, driver, fileDir, fileName, inThisWindow = True) ->str:
        finishBtn = driver.find_element_by_xpath('//button[text()= "完成"]')
        driver.execute_script("arguments[0].scrollIntoView();", finishBtn)
        finishBtn.click()
        if not inThisWindow:
            while pyscreeze.locateOnScreen(r'D:\flow_1_resource\{}.png'.format('blue'), minSearchTime=10) is None: time.sleep(5)
            driver.switch_to.parent_frame()
            while not self._waitelement_by_id(driver, 'com.ibm.bi.classicviewer.runMenu', 5): ...
            driver.find_element_by_id('com.ibm.bi.classicviewer.runMenu').click()
            time.sleep(2)
            driver.find_element_by_class_name('commonMenuItems').find_elements_by_tag_name('a')[2].click()
            while len(driver.window_handles) != 2: time.sleep(1)
        self._window_waiting('另存为')
        self._press('backspace', 1)
        save = win32gui.FindWindow('#32770', '另存为')
        if not os.path.exists(fileDir): os.makedirs(fileDir)
        edit = uiautomation.ControlFromHandle(save).EditControl(searchDepth=10, Name='文件名:').NativeWindowHandle
        filePath = fileDir + os.path.sep + fileName
        win32api.SendMessage(edit, win32con.WM_SETTEXT, None, filePath)
        self._hot_keys(1, 'alt', 's')
        if len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
        driver.switch_to.default_content()
        driver.execute_script("document.getElementsByClassName('tabs-wrap')[0].getElementsByTagName('li')[1].getElementsByTagName('a')[1].click();")
        self._web_driver = driver
        return filePath

    # 窗口等待
    def _window_waiting(self, title):
        while win32gui.FindWindow('#32770', title) == 0: time.sleep(1)

    # 根据元素 id 等待元素
    def _waitelement_by_id(self, driver, id, seconds):
        try:
            WebDriverWait(driver, seconds).until(EC.presence_of_all_elements_located((By.ID, id)))
        except Exception:
            return False
        return True

    # 系统登入
    def _login_system(self, robot):
        chromeOptions = webdriver.ChromeOptions()
        chromeOptions.add_experimental_option("prefs", {"download.prompt_for_download": True})
        chromeOptions.add_argument('--args --disable-web-security --user-data-dir')
        driver = webdriver.Chrome(options=chromeOptions)
        driver.get('http://mis.sinosafe.com.cn/pages/index.html')
        driver.maximize_window()
        while not self._waitelement_by_id(driver, 'username', 5): ...
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
        self._web_driver = driver

    # 退出浏览器
    def _quit_browser(self, driver):
        driver.quit()

if __name__ == "__main__":
    try:
        robot = PolicyCostAllocation()
        # 系统登入
        robot._login_system(robot)
        # 导出新车险销售费用数据表 / 处理新车险销售费用数据表
        robot._excel_dispose(robot._xcx_table_export(robot._web_driver), 1)
        # 导出新财产险销售数据费用表 / 处理新财产险销售数据费用表
        robot._excel_dispose(robot._xccx_table_export(robot._web_driver), 2)
        # 导出新人生险销售费用表 / 处理新人生险销售费用表
        robot._excel_dispose(robot._xrsx_table_export(robot._web_driver), 3)
        # 导出实收保费统计表 / 处理实收统计表
        robot._excel_dispose(robot._ssbf_table_export(robot._web_driver), 4)
    except Exception as e:
        print(e)
    finally:
        print('end')
        #robot._quit_browser(driver)