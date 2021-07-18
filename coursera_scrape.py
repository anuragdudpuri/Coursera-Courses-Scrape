from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import time
import numpy as np
import pandas as pd


startTime = time.time()
opts = Options()
opts.add_argument("--incognito")
PATH = "dataset/chromedriver"
#opts.add_argument("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36")
driver = webdriver.Chrome(PATH , options=opts)

driver.maximize_window()
driver.get("https://www.coursera.org/in")
print(driver.title)

MasterTableDict = {
	"Tags_(Field_of_Study)" : 	[],
	"Tags_(Tools_and_Tech)" :   [],
	"Title"					:	[],
	"Description"			:	[],
	"Week_Titles"			:	[],
	"Language"				:	[],
	"Rating"				:	[],
	"Number_of_Ratings"		:	[],
	"Offered_By"			:	[],
	"Website"				:	[],
	"Enrolled_Students"		:	[],
	"Instructor"			:	[],
	"Level"					:	[],
	"Price"					:	[],
	"Time_(to_display)"		:	[],
	"Pace"					:	[],
	"Pre-requisite"			:	[],
	"Certificate"			:	[],
	"Link"					:	[],
	"Image_Link"			:	[],
	"Page"					:	[]
}

searchList = ["digital marketing"]

for searchWord in searchList:
	print("\n")
	print("Searching for: " + searchWord)
	courseNumber = 1
	
	
	search = WebDriverWait(driver,20).until(
		lambda x:x.find_element(By.CLASS_NAME,'react-autosuggest__input')
	)
	search.send_keys(searchWord)
	search.send_keys(Keys.RETURN)

	filterElement = WebDriverWait(driver, 10).until(
		lambda x:x.find_element(By.XPATH,"/html/body/div[2]/div/main/div/div/div[1]/div/main/div/div[2]/div/div[2]/div[7]/div")
	)
	filterElement.click()
	listElement = WebDriverWait(driver, 3).until(
		lambda x:x.find_element(By.XPATH,"/html/body/div[2]/div/main/div/div/div[1]/div/main/div/div[2]/div/div[2]/div[7]/div/div[1]/div[2]/div/div[2]/div/button")
	)
	listElement.click()
	
	#Enter Link Here to start from a particular page
	#driver.get("https://www.coursera.org/search?query=data%20science&page=71&index=prod_all_products_term_optimization&entityTypeDescription=Courses")
	
	currWindow = driver.window_handles[0]
	pageLimit = 28	

	for page in range(1,pageLimit+1):
		print("Page Number: ", page, end = " ")
		print("Page Link: ", driver.current_url)

		try:
			allElements = WebDriverWait(driver, 10).until(
				lambda x : x.find_elements(By.XPATH,"//li[@class='ais-InfiniteHits-item']")
			)
			
			#print(allElements)
			for ele in allElements:
				imgLink = ele.find_element(By.XPATH, ".//img[@class='product-photo']").get_attribute("src")
				#----- Image Link -----#
				MasterTableDict['Image_Link'].append(imgLink)
				#----- Image Link -----#
				
				MasterTableDict['Page'].append(page)
				ele.click()
				nextWindow = driver.window_handles[1]
				driver.switch_to.window(nextWindow)
				
				print("Course Link: ", driver.current_url)

				#---- Tags Field of Study -----#
				MasterTableDict['Tags_(Field_of_Study)'].append(searchWord)
				#---- Tags Field of Study -----#

				#---- Website -----#
				MasterTableDict['Website'].append('Coursera')
				#---- Website -----#

				#---- Pace -----#
				MasterTableDict['Pace'].append('Self Paced')
				#---- Pace -----#

				#---- Link -----#
				MasterTableDict['Link'].append(driver.current_url)
				#---- Link -----#

				#----- Title ----- #
				try:
					titleElement = WebDriverWait(driver, 5).until(
						lambda x: x.find_element(By.XPATH, "//h1[@class='banner-title banner-title-without--subtitle m-b-0']")
						)
					#print(titleElement.text)
					MasterTableDict['Title'].append(titleElement.text)
				except:
					MasterTableDict['Title'].append("No title provided")
				#----- Title ----- #

				#----- Rating and Number of Ratings ----- #
				try:
					RatingElement = driver.find_element(By.XPATH, "//span[@data-test='number-star-rating']")
					ratings = RatingElement.text.split("\n")[0]
					MasterTableDict['Rating'].append(ratings)
				except:
					MasterTableDict['Rating'].append("No Rating Available")

				try:
					NumberRatingElement = driver.find_element(By.XPATH,"//div[@class='_wmgtrl9 color-white ratings-count-expertise-style']")
					numberOfRatings = NumberRatingElement.text
					#print(ratings,numberOfRatings)
					MasterTableDict['Number_of_Ratings'].append(numberOfRatings)
				except:
					MasterTableDict['Number_of_Ratings'].append("No Number Available")
				#----- Rating and Number of Ratings ----- #

				#----- Enrolled -----#
				try:
					EnrolledElement = driver.find_element(By.XPATH , "//div[@class='_1fpiay2']")
					space = EnrolledElement.text.find(' ')
					txt = EnrolledElement.text[0:space]
					#print(txt)
					MasterTableDict['Enrolled_Students'].append(txt)
				except:
					MasterTableDict['Enrolled_Students'].append("No Enrolled Students")
				#----- Enrolled -----#

				#----- Instructor-----#
				try:
					InstructorElement = driver.find_elements(By.XPATH, "//h3[@class='instructor-name headline-3-text bold']")
					txtList = []
					for instructor in InstructorElement:
						txtList.append(instructor.text)
					#print(txtList)
					MasterTableDict['Instructor'].append(",".join(txtList))
				except:
					MasterTableDict['Instructor'].append("No Instructor Available")
				#----- Instructor-----#
				

				#----- OfferedBy-----#
				try:
					OfferedElement = driver.find_elements(By.XPATH, "//h3[@class='headline-4-text bold rc-Partner__title']")
					txtList = []
					for offer in OfferedElement:
						txtList.append(offer.text)
					#print(txtList)
					MasterTableDict['Offered_By'].append(",".join(txtList))
				except:
					MasterTableDict['Offered_By'].append("No Offered By Available")
				#----- OfferedBy-----#

				#----- Language Certificate Time Level Pre-----#
				try:
					DetailsElement = driver.find_elements(By.XPATH, "//div[@class='_cs3pjta _1pfe6xlx p-l-2 p-r-0']/div[@class='ProductGlance']/div[@class='_y1d9czk m-b-2 p-t-1s ']")
					txtList = []
					for lang in DetailsElement:
						txtList.append(lang.text)
					certificate = False
					level = ""
					pre = ""
					#print(txtList)
					for txt in txtList:
						if txt.startswith("Approx."):
							timeLimit = txt
						elif "Level" in txt:
							end = txt.find("\n")
							if end != -1:
								level = txt[0:end]
								pre = txt[end:]
							else:
								level = txt
						elif "Certificate" in txt:
							certificate = True
							MasterTableDict['Certificate'].append("Yes")
						elif "Subtitles" in txt:
							end = txt.find("\n")
							language = txt[0:end]
					#print(timeLimit,level,certificate,language,pre)
					if certificate == False:
						MasterTableDict['Certificate'].append("No")

					if level == "":
						MasterTableDict['Level'].append("No Level Given")
					else:
						MasterTableDict['Level'].append(level)

					if pre == "":
						MasterTableDict['Pre-requisite'].append("No Pre-requisite Given")
					else:
						MasterTableDict['Pre-requisite'].append(pre)
					MasterTableDict['Time_(to_display)'].append(timeLimit)
					MasterTableDict['Language'].append(language)
				except:
					print("No box on right")
					MasterTableDict['Certificate'].append("No")
					MasterTableDict['Pre-requisite'].append("No Pre-requisite Given")
					MasterTableDict['Level'].append("No Level Given")
					MasterTableDict['Time_(to_display)'].append("No Time Given")
					MasterTableDict['Language'].append("No Language Given")
				#----- Language Certificate Time Level Pre-----#


				#----- Price -----#
				MasterTableDict['Price'].append("No Price Given")
				#----- Price -----#


				#----- Description -----#
				try:
					DescriptionElement = driver.find_element(By.XPATH, "//div[@class='m-t-1 description']")
					txt = DescriptionElement.text.split("\n")[0]
					#print(txt)
					MasterTableDict['Description'].append(txt)
				except:
					MasterTableDict['Description'].append("No Description Given")
				#----- Description -----#


				#----- Skill Tags -----#
				try:
					SkillElement = driver.find_element(By.XPATH, "//div[@class='Skills m-y-2 p-x-2 p-t-1 p-b-2 border-a']/div[@class='_t6niqc3']")
					txtList = SkillElement.text.split("\n")
					#print(txtList)
					txt = ",".join(txtList)
					#print(txt)
					MasterTableDict['Tags_(Tools_and_Tech)'].append(txt)
				except:
					MasterTableDict['Tags_(Tools_and_Tech)'].append("No Tags Given")
				#----- Skill Tags -----#


				#----- Week Titles-----#
				try:
					showMoreElement = driver.find_element(By.XPATH,"//button[@data-track-component='syllabus_button_show_more']")
					showMoreElement.click()
				except:
					print("No show more button")
				try:
					weekElements = driver.find_elements(By.XPATH, "//h2[@class='headline-2-text bold m-b-2 ']")
					txtList = []
					for week in weekElements:
						txtList.append(week.text)

					#print(txtList)
					MasterTableDict['Week_Titles'].append(",".join(txtList))
				except:
					MasterTableDict['Week_Titles'].append("No Week Information")
				#----- Week Titles-----#


				driver.close()
				driver.switch_to.window(currWindow)

		except:
			print("Error here")
			driver.refresh()
			# allElements = WebDriverWait(driver, 10).until(
			# 	lambda x : x.find_elements(By.XPATH,"//li[@class='ais-InfiniteHits-item']")
			# )
			# page -= 1
			break

		nextPage = driver.find_element(By.XPATH, "//button[@data-track-component='pagination_right_arrow']")
		nextPage.click()
		time.sleep(5)

minVal = 1000
for arr in MasterTableDict:
	minVal = min(minVal, len(MasterTableDict[arr]))
for arr in MasterTableDict:
	while len(MasterTableDict[arr]) > minVal:
		MasterTableDict[arr].pop()
masterDf = pd.DataFrame.from_dict(MasterTableDict)
#print(masterDf)
masterDf.to_excel("coursera_MasterTable_dmcourses.xlsx")
driver.close()
