import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

def main():
    filePath = input("Path to excel spreadsheet:")
    if filePath == "":
        #default path
        filePath = "C:\\Users\\baily\\Scripts\\Python\\YuGiOh 401k\\YuGiOhCards.xls"
    else:
        #ensure filePath is a usage raw string
        pass
    if ".xls" not in filePath:
        exit("Invalid path, no spreadsheet (.xls) found:"+ str(filePath))
    file = pd.ExcelFile(filePath)
    #Must in the following format
    #       A       |       B       |       C
    #   Card Name        Edition        Set Number
    #  BIO-MAGE     |  1st Edition	| LON-043
    data = pd.read_excel(file)
    data = data.to_numpy()
    low_values = []
    high_values = []
    aver_values = []
    werid = []
    redirect = []
    driverPath = input("Is the chromedriver.exe in the same folder?(y/n):")
    if driverPath == "y":
        driverPath = ""
        driverPath = filePath[:filePath.rfind("\\")] + "\chromedriver.exe"
    else:
        exit("\nWell, move it there.")
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--log-level=3')
    driver=webdriver.Chrome(executable_path=driverPath,options=chrome_options)
    for card in data:
        name = str(card[0])
        print("Card:"+name)
        driver.get("https://yugiohprices.com/")
        driver.find_element_by_name("search_text").send_keys(name)
        driver.find_element_by_xpath('/html/body/div[1]/div/form/input[2]').click()
        print("sent keys")
        if driver.find_element_by_id('item_name').text.upper() == name:
            pass
        else:
            werid.append(name)
        if  "Browse Cards" in driver.title:
            cap = " ".join([word.capitalize()for word in name.split(" ")])
            redirect.append(cap)
            try:
                el = driver.find_element_by_link_text(cap)
                el.click()
            except:
                el = driver.find_element_by_xpath("/html/body/div[4]/table/tbody/tr[1]/td[2]/a")
                el.click()
        print("Waiting on server... (timeout=600secs")
        wait = WebDriverWait(driver, 600)
        element = wait.until(EC.invisibility_of_element_located((By.ID, 'loading')))
        print("Loading warning disappeared")
        for nu in range(1,999,2):
            if "Gateway time-out" in driver.title:
                driver.refresh()
            title = driver.find_element_by_xpath("/html/body/div[5]/div/table["+str(nu)+"]/tbody/tr[1]/td/b").text
            if "Element not found." in title:
                input("tell me to try again")
            title = driver.find_element_by_xpath("/html/body/div[5]/div/table["+str(nu)+"]/tbody/tr[1]/td/b").text
            web_name = title[title.index("-- ")+3:]
            if web_name == card[2]:
                lowest = driver.find_element_by_xpath('/html/body/div[5]/div/table['+str(nu+1)+']/tbody/tr/td[1]/table[1]/tbody/tr[1]/td[2]/p').text
                highest = driver.find_element_by_xpath('/html/body/div[5]/div/table['+str(nu+1)+']/tbody/tr/td[1]/table[1]/tbody/tr[2]/td[2]/p').text
                average = driver.find_element_by_xpath('/html/body/div[5]/div/table['+str(nu+1)+']/tbody/tr/td[1]/table[1]/tbody/tr[3]/td[2]/p').text
                low_values.append(lowest)
                high_values.append(highest)
                aver_values.append(average)
                print("lowest:"+lowest)
                print("highest:"+highest)
                print("average:"+average)
                break
            else:
                next
    driver.close()
     #Append mode isn't working for me
    #Creating new excel file to hold results instead
    df = pd.DataFrame({"Lowest": low_values, "Highest": high_values, "Average": aver_values})
    with pd.ExcelWriter(filePath[:filePath.rfind("\\")] + "\\results.xls", mode="w", engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Sheet1")
    # print("lowest")
    # print(low_values)
    # print("highest")
    # print(high_values)
    # print("averge")
    # print(aver_values)
    # f = open("lowest.txt", "w") 
    # for line in low_values:
    #     f.write(line+"\n")
    # f.close()
    # f = open("highest.txt", "w") 
    # for line in high_values:
    #     f.write(line+"\n")
    # f.close()
    # f = open("average.txt", "w") 
    # for line in aver_values:
    #     f.write(line+"\n")
    # f.close()
    print('Cards not searched:')
    print(werid)
    print('Cards that got redirected:')
    print(redirect)
    y = input("You can find the results in 'results.xls' and copy/paste into the original excel file. You may now exit")
    exit()
    
    
        
    



if __name__ == '__main__':
	main()