import time
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from PIL import ImageGrab
import xlrd
import xlwt

'''
打开要处理的文件test.xlsx并读取Shhet1
'''
workbook=xlrd.open_workbook("test.xlsx")  #文件路径
worksheet=workbook.sheet_by_name("Sheet1")
book=xlwt.Workbook(encoding="utf-8",style_compression=0)

'''
输出文件的页叫做compare
'''
sheet_output= book.add_sheet('compare', cell_overwrite_ok=True)

'''
每一列存放在不同变量中 
'''
col_data=worksheet.col_values(0)
col_address=worksheet.col_values(1)
col_title=worksheet.col_values(5)

index=0
option = webdriver.ChromeOptions()
option.add_argument('--headless')
# option.add_argument('--no-sandbox')
# option.add_argument('--start-maximized')
driver = webdriver.Chrome(executable_path='bin/chromedriver')  # 打开浏览器
# driver = webdriver.Chrome(executable_path='chromedriver',chrome_options=option)#静默运行，不打开浏览器，但截图尺寸有误
driver.maximize_window()  # 最大化窗口，方便截图
for line in col_data:
    while 1:
        try:
            if not(line=='单位名称'):
                driver.switch_to.window(driver.window_handles[0])
                driver.get("https://xin.baidu.com/")
                search_Name = driver.find_element_by_css_selector("[class='search-text ZX_LOG_INPUT']")
                search_Name.send_keys(line)
                search_Name.send_keys(Keys.ENTER)
                #time.sleep(1)
                search_Result = driver.find_elements_by_css_selector("[class='zx-list-item']")
                count_Result = (len(search_Result))
                if count_Result == 0:
                    address='没有找到结果'
                    sheet_output.write(index, 0, line)
                    sheet_output.write(index, 1, col_address[index])
                    sheet_output.write(index, 2, address)
                    sheet_output.write(index, 5, col_title[index])
                else:
                    driver.find_element_by_class_name('zx-ent-logo').click()
                    #time.sleep(1)
                    driver.switch_to.window(driver.window_handles[1])
                    address_parent = driver.find_elements_by_css_selector(
                        "[class='zx-detail-company-item right-lab item-address']")
                    address=address_parent[0].text.split('：')[1]
                    print (address)
                    # picture_url = driver.save_screenshot('截图/'+line+'.png')
                    im = ImageGrab.grab()  # 截屏
                    [width,height]=im.size
                    im=im.crop((0,0,width,height/2))
                    im.save('截图/'+str(col_title[index]) + '.png')
                    
                    sheet_output.write(index, 0, line)
                    sheet_output.write(index, 1, col_address[index])
                    sheet_output.write(index, 2, address)
                    sheet_output.write(index, 5, col_title[index])
                    driver.close()
            else:
                for i in range(5):
                    tmp=worksheet.col_values(i)[index]
                    sheet_output.write(index, i, tmp)

            index+=1
            book.save('out.xls')
            break
        except KeyboardInterrupt:
            print ("终止")
            os._exit(0)
        except:
            pass
        finally:
            pass
