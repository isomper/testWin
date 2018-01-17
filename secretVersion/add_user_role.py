#coding=utf-8
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class initDriver:
    u'''web驱动初始化'''
    def open_driver(self,ipAdd):
        driver = webdriver.Chrome()
        #driver = webdriver.Firefox()
        #窗口最大化
        driver.maximize_window()
        #打开IP地址对应的网页
        driver.get("https://" + ipAdd + "/fort/login")
        return driver    
        
class OptElment(object):
    def __init__(self,driver):
        #selenium驱动
        self.driver = driver
    
    u'''等待元素出现后再定位元素并点击EC
        parameter:
            - type:定位的类型，如id,name,tag name,class name,css,xpath等
            - value：页面的元素值
            - timeout:超时前等待的时间
        return：定位元素并点击该元素
    '''
    def find_element_wait_and_click_EC(self,type,value,timeout=10):
            if type == "id":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.ID, value))).click()
            elif type == "xpath":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.XPATH, value))).click()
            elif type == "name":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.NAME, value))).click()
            elif type == "tagname":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, value))).click()
            elif type == "classname":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, value))).click()
            elif type == "css":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, value))).click()
            elif type == "link":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.LINK_TEXT, value))).click()
            elif type == "plt":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, value))).click()
    
    u'''清空文本框的内容
        parameter:
            - type:定位的类型，如id,name,tag name,class name,css,xpath等
            - value：页面的元素值
            - timeout:超时前等待的时间
    '''
    def find_element_wait_and_clear_EC(self, type, value, timeout=10):
    
        if type == "id":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.ID, value))).clear()
        elif type == "xpath":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.XPATH, value))).clear()
        elif type == "name":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.NAME, value))).clear()
        elif type == "tagname":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, value))).clear()
        elif type == "classname":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, value))).clear()
        elif type == "css":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, value))).clear()
        elif type == "link":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.LINK_TEXT, value))).clear()
        elif type == "plt":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, value))).clear()
    
    u'''等待元素出现后再定位元素并点击EC
        parameter:
            - type:定位的类型，如id,name,tag name,class name,css,xpath等
            - value：页面的元素值
            - timeout:超时前等待的时间
        return：定位元素并点击该元素
    '''
    def find_element_wait_and_click_EC(self,type,value,timeout=10):
            if type == "id":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.ID, value))).click()
            elif type == "xpath":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.XPATH, value))).click()
            elif type == "name":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.NAME, value))).click()
            elif type == "tagname":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, value))).click()
            elif type == "classname":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, value))).click()
            elif type == "css":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, value))).click()
            elif type == "link":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.LINK_TEXT, value))).click()
            elif type == "plt":
                WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, value))).click()
    
    u'''定位元素
        parameter:
            - type:定位的类型，如id,name,tag name,class name,css,xpath等
            - value：页面的元素值
            - timeout:超时前等待的时间
    '''
    def find_element_EC(self, type, value, timeout=10):
    
        if type == "id":
            return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.ID, value)))
        elif type == "xpath":
            return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.XPATH, value)))
        elif type == "name":
            return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.NAME, value)))
        elif type == "tagname":
            return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, value)))
        elif type == "classname":
            return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, value)))
        elif type == "css":
            return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, value)))
        elif type == "link":
            return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.LINK_TEXT, value)))
        elif type == "plt":
            return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, value)))

    u'''填写文本框的内容
        parameter:
            - type:定位的类型，如id,name,tag name,class name,css,xpath等
            - value：页面的元素值
            - key:文本框内容
            - timeout:超时前等待的时间
    '''
    def find_element_sendkyes_EC(self, type, value, key, timeout=10):
    
        if type == "id":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.ID, value))).send_keys(key)
        elif type == "xpath":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.XPATH, value))).send_keys(key)
        elif type == "name":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.NAME, value))).send_keys(key)
        elif type == "tagname":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, value))).send_keys(key)
        elif type == "classname":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, value))).send_keys(key)
        elif type == "css":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, value))).send_keys(key)
        elif type == "link":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.LINK_TEXT, value))).send_keys(key)
        elif type == "plt":
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, value))).send_keys(key)
    
    def select_element_by_visible_text(self,selem,text):
        return Select(selem).select_by_visible_text(text)
    
    u'''定位到topFrame'''
    def switch_to_top(self):
        self.driver.switch_to_default_content()
        #如果content1的frame加载完毕，定位到content1的frame
        self.driver.switch_to_frame("content1")
        self.driver.switch_to_frame("topFrame")
    
    u'''定位到mainFrame'''
    def switch_to_main(self):
        self.driver.switch_to_default_content()
        self.driver.switch_to_frame("content1")
        self.driver.switch_to_frame("mainFrame")
    
    u'''返回上级frame'''
    def switch_to_content(self):
        self.driver.switch_to_default_content()
    
    u'''从一个frame跳转到其他frame
        Parameters:
            - frameName:要跳转到的frame的名字      
    '''
    def from_frame_to_otherFrame(self,frameName):
        self.switch_to_content()
        
        if frameName == "mainFrame":
            #定位到mainFrame
            self.switch_to_main()
            
        elif frameName == "topFrame":
            #定位到topFrame            
            self.switch_to_top()
    
    u'''菜单选择
        Parameters:
            - levelText1：1级菜单文本
            - levelText2：2级菜单文本 
            - levelText3：3级菜单文本           
    '''
    def select_menu(self,levelText1,levelText2='no',levelText3='no'):
        #进入topframe
        self.from_frame_to_otherFrame("topFrame")
        
        #点击一级菜单
        WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.LINK_TEXT,levelText1))).click()
        
        #如果有2级菜单，再点击2级菜单
        if levelText2 != 'no':
            WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.LINK_TEXT,levelText2))).click()
        
        #如果有3级菜单，根据名称点击3级菜单
        if levelText3 != 'no':
            frameElem.from_frame_to_otherFrame("leftFrame")
            WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.LINK_TEXT,levelText3))).click()
    
    u'''点击弹出框上的确定按钮'''
    def click_login_msg_button(self):
        #确定按钮
        self.driver.switch_to_default_content()
        OKBTN = "//div[@id='aui_buttons']/button[1]"
        return self.find_element_wait_and_click_EC('xpath',OKBTN)

    u'''输入用户名和密码登录'''
    def login(self):
        self.optEle.find_element_sendkyes_EC('id','username','isomper')
        self.optEle.find_element_sendkyes_EC('id','pwd','1')
        time.sleep(1)
        self.optEle.find_element_wait_and_click_EC('id','do_login')

u'''添加角色'''
class AddRole(object):
    def __init__(self,driver):
        self.driver = driver
        self.optEle = OptElment(driver)
    
    u'''点击角色添加按钮'''
    def role_add_button(self):
        self.optEle.from_frame_to_otherFrame('mainFrame')
        self.optEle.find_element_wait_and_click_EC('id','add_role')
    
    u'''设置角色名称和简写
        parameter:
        - roleName：角色名称
        - abbrName：角色简称'''
    def set_role_name(self,roleName,abbrName):
        self.optEle.from_frame_to_otherFrame('mainFrame')
        self.optEle.find_element_sendkyes_EC('id','fortRoleName',roleName)
        self.optEle.find_element_sendkyes_EC('id','fortRoleShortName',abbrName)
   
    u'''展开角色树(三角形)；
        parameter:
        - treeDemo_switch:角色演示树'''
    def set_tree_demo_switch(self,treeDemo_switch):
        self.optEle.find_element_wait_and_click_EC('id',treeDemo_switch)
        
    u'''勾选一层角色; 
        parameter:
        - treeDemo_check:角色树勾选框'''
    def set_tree_demo_check(self,treeDemo_check):
        self.optEle.find_element_wait_and_click_EC('id',treeDemo_check)       
    
    u'''勾选底层角色框; 
        parameter:
       - inputName:底层角色勾选框'''
    def set_input_click(self,inputName):
        self.optEle.find_element_wait_and_click_EC('name',inputName)
        
    u'''选择其他角色；
    parameter:
    - roleName：角色名称 '''
    def set_other_role(self,roleName):
        #self.driver.execute_script("window.scrollBy(1000,0)","")
        sele = self.optEle.find_element_EC('id','allOtherPrivileges')
        time.sleep(1)
        self.optEle.select_element_by_visible_text(sele,roleName)
        self.optEle.find_element_wait_and_click_EC('id','add_privileges')
        
    #保存角色
    def save_role_button(self):
        self.optEle.find_element_wait_and_click_EC('id','save_role')
        self.optEle.click_login_msg_button()
        time.sleep(1)
        self.optEle.switch_to_main()
        self.optEle.find_element_wait_and_click_EC('id','history_skip')
        
    #编辑运维角色  
    def edit_operation_role(self):
        #编辑运维操作员
        self.optEle.switch_to_main()
        time.sleep(1)
        #点击编辑按钮
        editXpath = "//table[@id='content_table']/tbody/tr[2]/td[6]/input[1]"
        self.optEle.find_element_wait_and_click_EC('xpath',editXpath)
        self.set_tree_demo_switch('treeDemo_1_switch')
        time.sleep(1)     
        #去掉监控、回放、窗体识别、阻断、下载
        self.set_input_click('1001010000002')
        self.set_input_click('1001010000003')
        self.set_input_click('1001010000010')
        self.set_input_click('1001010000011')
        self.set_input_click('1001010000013')
        self.save_role_button()
    
    #添加系统管理员角色
    def set_sysAdmin_role(self):
        time.sleep(1)
        self.role_add_button()
        self.set_role_name(u'系统管理员',u'系管')
        #勾选组织定义和用户
        self.set_tree_demo_switch('treeDemo_3_switch')
        time.sleep(1)
        self.set_tree_demo_switch('treeDemo_4_check')
        self.set_tree_demo_check('treeDemo_8_check')
        #用户去掉角色
        self.set_input_click('1002020000008')
        #勾选流程控制
        self.set_tree_demo_switch('treeDemo_21_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_21_check')
        #去掉全部历史
        self.set_tree_demo_check('treeDemo_25_check')
        #口令计划
        self.set_tree_demo_switch('treeDemo_26_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_26_check')
        self.driver.execute_script("window.scrollBy(1000,0)","")
        #勾选系统配置
        self.set_tree_demo_switch('treeDemo_38_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_38_check')
        #去掉初始化设置
        self.set_tree_demo_switch('treeDemo_54_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_56_check')
        #去掉维护配置
        self.set_tree_demo_check('treeDemo_59_check')
        self.set_tree_demo_switch('treeDemo_38_switch')
        #选择密码包角色
        self.set_other_role(u'密码包接收人')
        time.sleep(1)
        #保存角色
        self.save_role_button()
    
    #添加安全保密管理员角色
    def set_secAdmin_role(self):
        time.sleep(1)
        self.role_add_button()
        self.set_role_name(u'安全保密管理员',u'安保')
        #运维管理用户
        self.set_tree_demo_switch('treeDemo_3_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_8_check')
        #只勾选删除、编辑、用户状态、证书
        self.set_input_click('1002020000001')
        self.set_input_click('1002020000004')
        self.set_input_click('1002020000005')
        self.set_input_click('1002020000006')
        self.set_input_click('1002020000007')
        self.set_input_click('1002020000008')
        #资源、授权（不勾选可访问外部报表）、规则定义
        self.set_tree_demo_check('treeDemo_9_check')
        self.set_tree_demo_check('treeDemo_10_check')
        self.set_input_click('1002040000009')
        self.set_tree_demo_check('treeDemo_11_check')
        #配置审计、运维审计只勾选审计删除，命令详情，审批记录，查看历史，键盘记录，文件传输，剪切板
        self.set_tree_demo_switch('treeDemo_17_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_18_check')
        self.set_input_click('1003010000003')
        self.set_input_click('1003010000004')
        self.set_input_click('1003010000011')
        self.set_input_click('1003010000012')
        self.set_input_click('1003010000014')
        self.set_input_click('1003010000015')
        self.set_tree_demo_check('treeDemo_19_check')
        self.set_tree_demo_switch('treeDemo_17_switch')
        time.sleep(1)
        #流程控制（全部历史）
        self.set_tree_demo_switch('treeDemo_21_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_25_check')
        #计划任务只勾选备份文件查看、备份文件下载
        self.set_tree_demo_switch('treeDemo_26_switch')
        time.sleep(1)
        self.set_tree_demo_switch('treeDemo_27_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_28_check')
        self.set_input_click('1005010100001')
        self.set_input_click('1005010100002')
        self.set_input_click('1005010100003')
        self.set_input_click('1005010100005')
        time.sleep(1)
        self.set_input_click('1005010100007')
        self.set_input_click('1005010100008')
        self.set_input_click('1005010100009')
        self.set_input_click('1005010100010')
        self.set_tree_demo_check('treeDemo_29_check')
        time.sleep(1)
        self.set_input_click('1005010200001')
        self.set_input_click('1005010200002')
        self.set_input_click('1005010200003')
        self.set_input_click('1005010200005')
        time.sleep(1)
        self.set_input_click('1005010200007')
        self.set_input_click('1005010200008')
        self.set_input_click('1005010200009')
        time.sleep(1)
        #系统配置（审计存留）
        self.set_tree_demo_switch('treeDemo_38_switch')
        time.sleep(1)
        self.set_tree_demo_switch('treeDemo_59_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_60_check')
        #策略配置
        self.set_tree_demo_switch('treeDemo_65_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_65_check')
        self.set_tree_demo_switch('treeDemo_38_switch')
        time.sleep(1)
        #选择密码包角色
        self.set_other_role(u'解密密钥接收人')
        time.sleep(1)
        #保存角色
        self.save_role_button()
       
    #添加安全审计员角色
    def set_sysAudit_role(self):
        time.sleep(1)
        self.role_add_button()
        self.set_role_name(u'安全审计员',u'安审')
        #运维管理（用户只勾选删除、编辑、用户状态、证书）
        self.set_tree_demo_switch('treeDemo_3_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_8_check')
        #只勾选删除、编辑、用户状态、证书
        self.set_input_click('1002020000001')
        self.set_input_click('1002020000004')
        self.set_input_click('1002020000005')
        self.set_input_click('1002020000006')
        self.set_input_click('1002020000007')
        self.set_input_click('1002020000008')
        self.set_tree_demo_switch('treeDemo_3_switch')
        time.sleep(1)
        #审计管理（配置审计，告警归纳）
        self.set_tree_demo_switch('treeDemo_17_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_19_check')
        self.set_tree_demo_check('treeDemo_20_check')
        self.set_tree_demo_switch('treeDemo_17_switch')
        time.sleep(1)
        #报表管理
        self.set_tree_demo_switch('treeDemo_31_switch')
        time.sleep(1)
        self.set_tree_demo_check('treeDemo_31_check')
        #保存角色
        self.save_role_button()
        
u'''登录'''    
class UserLogin():
    def __init__(self,driver):
        #selenium驱动
        self.driver = driver
    def login(self):
        ele = OptElment(self.driver)
        ele.find_element_wait_and_clear_EC('id','username')
        ele.find_element_sendkyes_EC('id','username','isomper')
        ele.find_element_wait_and_clear_EC('id','pwd')
        ele.find_element_sendkyes_EC('id','pwd','1')
        time.sleep(1)
        ele.find_element_wait_and_click_EC('id','do_login')
    
u'''添加用户
    parameter:
        - account:用户账号
        - name：用户名称
        - pwd:用户密码
        - roleText:赋予用户的角色
'''  
class AddUser():
    def __init__(self,driver):
        #selenium驱动
        self.driver = driver
    
    def add_user(self,account,name,pwd,roleText):
        ele = OptElment(self.driver)
        time.sleep(1)
        #点击添加按钮
        ele.switch_to_main()
        ele.find_element_wait_and_click_EC('classname','btn_tj')
        #输入账号，名称，密码
        ele.find_element_wait_and_clear_EC('id','fortUserAccount')
        ele.find_element_sendkyes_EC('id','fortUserAccount',account)
        ele.find_element_wait_and_clear_EC('id','fortUserName')
        ele.find_element_sendkyes_EC('id','fortUserName',name)
        ele.find_element_wait_and_clear_EC('id','fortUserPassword')
        ele.find_element_sendkyes_EC('id','fortUserPassword',pwd)
        ele.find_element_wait_and_clear_EC('id','fortUserPasswordAgain')
        ele.find_element_sendkyes_EC('id','fortUserPasswordAgain',pwd)
        #切换到角色页面
        ele.find_element_wait_and_click_EC('id','userMessage')
        #为用户赋予角色
        selem = ele.find_element_EC('id','Roles')
        ele.select_element_by_visible_text(selem,roleText)
        ele.find_element_wait_and_click_EC('id','add_roles')
        #保存用户
        ele.find_element_wait_and_click_EC('id','save_user')
        #点击弹框确定按钮
        ele.click_login_msg_button()
        time.sleep(1)
        #点击返回
        ele.switch_to_main()
        ele.find_element_wait_and_click_EC('id','history_skip')

if __name__ == "__main__" :
    initDriver = initDriver()
    driver = initDriver.open_driver('172.16.10.183')
    loginEle = UserLogin(driver)
    loginEle.login()
    optElem = OptElment(driver)
    optElem.select_menu(u'角色管理',u'角色定义')
    roleEle = AddRole(driver)
    roleEle.edit_operation_role()
    roleEle.set_sysAdmin_role()
    roleEle.set_secAdmin_role()
    roleEle.set_sysAudit_role()
    optElem.select_menu(u'运维管理',u'用户')
    userEle = AddUser(driver)
    userEle.add_user('sysAdmin',u'系统管理员','sysAdmin@123','系统管理员')
    userEle.add_user('sysAudit',u'安全审计员','sysAudit@123','安全审计员')
    userEle.add_user('secAdmin',u'安全保密管理员','secAdmin@123','安全保密管理员')
    
    
    
    
    
        

