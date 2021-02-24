import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

url_path = r'http://newmail.sgcc.com.cn/webmail/login/login.do'
driver_path = r'C:\Users\Admin\AppData\Local\Google\Chrome\Application\chromedriver'
Dirver = webdriver.Chrome(executable_path=driver_path)

#登录邮件系统
def log_in():
    Dirver.get(url_path)
    Dirver.find_element_by_xpath('//input[@id="usernumber"]').send_keys('ml@hn.sgcc.com.cn')
    Dirver.find_element_by_xpath('//input[@id="password"]').send_keys('hn@4306!')
    Dirver.find_element_by_xpath('//span[@id="b_text"]').click()
    time.sleep(2)


#首页图片截取
def au_picture():
    '''将首页今天、昨天、更早的邮件截图'''
    Dirver.find_element_by_xpath('//span[@class="tf" and text()="收件箱"]').click()
    time.sleep(1)
    Tobay = Dirver.find_elements_by_xpath('//div[@class="mailList list"]/div/table[1]/tbody/tr')

    num = 0
    for tday in Tobay:
        text =tday.get_attribute('href')
        print(text)
        time.sleep(1)
        text_in = Dirver.find_element_by_xpath('//div[@class="mailList list"]/div/table[1]/tbody/tr[@id="tr_sys1_{}"]/td[5]'.format(num)).text
        if ':' in text_in:
            text_info = text_in.replace(':','')
            print(text_info)
        time.sleep(1)
        Dirver.find_element_by_xpath('//div[@class="mailList list"]/div/table[1]/tbody/tr[@id="tr_sys1_{}"]'.format(num)).click()
        Dirver.set_window_size(1200, 900)
        Dirver.get_screenshot_as_file('D:\\pychar_program_image\\venv\\screenshot\\{}.png'.format(text_info))
        time.sleep(1)
        Dirver.find_element_by_xpath('//p[@style="width: 75px;" and text()="收件箱"]').click()
        time.sleep(1)
        num += 1

    # Yesterday = Dirver.find_elements_by_xpath('//div[@class="mailList list"]/div/table[2]/tbody')
    # sceendnum = len(Yesterday)
    # for yday in range(sceendnum):
    #     Dirver.find_element_by_xpath('//table[@id="period_table_sys10"/tbody/id[@"tr_sys1_{}"]'.format(yday + firstnum)).click()
    #
    # Morday = Dirver.find_elements_by_xpath('//div[@class="mailList list"]/div/table[3]/tbody')
    # for mday in range(len(Morday)):
    #     Dirver.find_element_by_xpath('//table[@id="period_table_sys10"/tbody/id[@"tr_sys1_{}]'.format(mday + firstnum + Morday)).click()

#其他页面截取
def other_picture():
    pass

if __name__ == "__main__":
    log_in()
    au_picture()
