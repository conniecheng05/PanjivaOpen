# PanjivaOpen
Search thru each account name in excel file and check if this one exists in database


```python
	from selenium import webdriver
	from selenium.webdriver.common.keys import Keys
	from selenium.common.exceptions import NoSuchElementException
	import time
	import xlrd
```

open chrome, login
```python
	driver = webdriver.Chrome("Test/Driver/chromedriver.exe")
	driver.get("https://panjiva.com/account/login")
	driver.find_element_by_name("email").send_keys("xxx.com")
	driver.find_element_by_name("password").send_keys("xxxxx")
```

found out the path is gonna change frequently, now only catch these two paths
```python
	try:
		driver.find_element_by_id("account_login").click()
	except NoSuchElementException:
		print('')
	try:
		driver.find_element_by_xpath('//*[@id="main_login_signin"]').click()
		except NoSuchElementException:
	print('')
```

try to see which path work
```python
	try:
		driver.find_element_by_xpath('//*[@id="wm-shoutout-103922"]/div[1]').click()
	except NoSuchElementException:
		print('element(22) not found')

	try:
		driver.find_element_by_xpath('//*[@id="wm-shoutout-93799"]/div[1]').click()
	except NoSuchElementException:
		print('element(99) not found')
```

Open excel sheet

```python
	wb = xlrd.open_workbook(r'C:\Users\Gazelle_OriginalData.xlsx')
```

get sheet name consolidated customer,and active worksheet

```python
	sn = wb.sheet_names()
	sheet0 = wb.sheet_by_name(sn[0])
```
number of row

```python
	num_row = sheet0.nrows
	num_row = num_row + 1
```

starting loop

```python
	for i in range(27,num_row):
			compy = sheet0.cell(i,0)
			driver.find_element_by_id("header-search").send_keys(compy.value)
			driver.find_element_by_name("button").click()
```

if page not found

```python
			source = driver.page_source
			if 'No Results Found' in source:
				driver.find_element_by_xpath('//*[@id="header-search"]').send_keys(Keys.CONTROL + "a")
				driver.find_element_by_id("header-search").send_keys(Keys.DELETE)
			else:
```
download
```python				
				driver.find_element_by_class_name("export-search-link").click()
				driver.find_element_by_name("download_excel").click()
				time.sleep(3)
```
there is another popup window to close
```python
				driver.find_element_by_xpath('//*[@id="excel_export_facebox"]/div/a/span').click()
				driver.find_element_by_xpath('//*[@id="header-search"]').send_keys(Keys.CONTROL+"a")
				driver.find_element_by_id("header-search").send_keys(Keys.DELETE)
```
