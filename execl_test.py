import openpyxl,sys,time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

iedirver_path = r'C:\Program Files\Internet Explorer\IEDriverServer.exe'
url = r'http://uums-a.hn.sgcc.com.cn/uums/login.view?method=getMenu&currentIndex=50'
Dirver = webdriver.Ie(executable_path=iedirver_path)


#运行时间增加装饰器
def time_out(f):
    def add_time(*args,**kwargs):
        time.sleep(3)
        ret = f(*args,**kwargs)
        return ret
    return add_time()


#统一用户登录
Dirver.get(url)
Submit = Dirver.find_element_by_xpath('//input[@name="commit"]')
Dirver.find_element_by_xpath('//input[@name="username"]').send_keys('admin')
Dirver.find_element_by_xpath('//input[@name="password"]').send_keys('n0v3ll??')
loginto = Dirver.find_element_by_xpath('//input[@name="commit"]').click()

#用户维护
@time_out
def add_cread():
    Dirver.find_element_by_xpath('//ul/li[@id="two2"]').click()
    time.sleep(1)
    Dirver.find_element_by_xpath('//div[@id="leftmenu_0"]/li[@id="twomenu1"]/a[@class="topmenu"]').click()

Dirver.switch_to.frame('iframe1')
Dirver.switch_to.frame('treeframe')

every_element = Dirver.find_elements_by_xpath('//*')
for elements in every_element:
    display_prop = elements.value_of_css_property('display')
    if display_prop == 'none':
        Dirver.execute_script('arguments[0].style.display = "block";',elements)

org = ['国网长沙供电公司', '国网长沙县供电公司']

if len(org) >= 1:
    test = Dirver.find_element_by_link_text('国网长沙供电公司')
    print(test)
if len(org) >= 2:
    test1 = Dirver.find_element_by_link_text('国网长沙县供电公司')
    print(test1)
if len(org) >= 3:
    pass





