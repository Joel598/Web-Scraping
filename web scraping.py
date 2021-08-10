#imports here
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import wget
import win32com.client as win32
from docx import Document

# path to chromedriver.exe
path = "C:\Program Files (x86)\chromedriver.exe" #path of web-browser driver

# create instance of webdriver
driver = webdriver.Chrome(path)

# Code to open a specific url
driver.get("https://www.instagram.com")

#exception handling
try:
#target username
    username = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='username']")))
    password = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='password']")))

#enter username and password
    username.clear()
    password.clear()
    username.send_keys("your username") #enter your username here 
    password.send_keys("your password") #enter your password here

#target the login button and click it
    log_in = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))).click()

#NOT NOW pop-up exception handled
    not_now = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Not Now')]"))).click()
    not_now2 = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Not Now')]"))).click()

#target the search input field
    searchbox = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Search']")))
    searchbox.clear()

#search for any hashtag here 
    keyword = "#dog"
    searchbox.send_keys(keyword)

    # Wait for 5 seconds
    time.sleep(5)
    searchbox.send_keys(Keys.ENTER)
    time.sleep(5)
    searchbox.send_keys(Keys.ENTER)
    time.sleep(5)

    #scroll down panel to scrape more images
    driver.execute_script("window.scrollTo(0, 5000)")

    #target all images on the page
    images = driver.find_elements_by_tag_name('img')
    images = [image.get_attribute('src') for image in images]
    images = images[:-2]

    print('Number of scraped images: ', len(images))

    #Save images to computer
    path = os.getcwd() #current working directory
    path = os.path.join(path, keyword[1:] + "s")

    #create the directory
    os.mkdir(path)


    #download images
    counter = 0
    for image in images:
        save_as = os.path.join(path, keyword[1:] + str(counter)+ '.jpg')
        wget.download(image, save_as)
        counter = counter + 1
#exception(error) saved in log file
except Exception as ex:
    with open("error.log", 'a') as errorlog:
        # print(time.asctime() + ":" + ex, file=errorlog)
        traceback.print_exc(file=errorlog)

#creating a word application object
wordApp = win32.gencache.EnsureDispatch('Word.Application') #create a word application object
wordApp.Visible = True # hide the word application
doc = wordApp.Documents.Add()

#Formating the document
doc.PageSetup.RightMargin = 20
doc.PageSetup.LeftMargin = 20
doc.PageSetup.Orientation = win32.constants.wdOrientLandscape
# a4 paper size: 595x842
doc.PageSetup.PageWidth = 595
doc.PageSetup.PageHeight = 842

# Inserting Tables
my_dir=r"E:\5th sem\assignment\webscraping(instagram) using selenium and inserting its images into word doc\program\dogs"
filenames = os.listdir(my_dir)
piccount=0
file_count = 0
for i in filenames:
    if i[len(i)-3: len(i)].upper() == 'JPG': # check whether the current object is a JPG file
        piccount = piccount + 1
        file_count= file_count + 1
print ("\n" + str(piccount) + "images will be inserted")
#print filenames
total_column = 2
total_row = int(piccount/total_column)+2
rng = doc.Range(0,0)
rng.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter
table = doc.Tables.Add(rng,total_row, total_column)
table.Borders.Enable = False
if total_column > 1:
    table.Columns.DistributeWidth()

#Collecting images in the same directory and inserting them into the document
frame_max_width= 167 # the maximum width of a picture
frame_max_height= 125 # the maximum height of a picture


piccount = 1

for index, filename in enumerate(filenames): # loop through all the files and folders for adding pictures
    #if os.path.isfile(os.path.join(os.path.abspath("."), filename)):
    if os.path.isfile(os.path.join(os.path.abspath(my_dir), filename)): # check whether the current object is a file or not
        if filename[len(filename)-3: len(filename)].upper() == 'JPG': # check whether the current object is a JPG file
            piccount = piccount + 1

            cell_column = (piccount % total_column + 1) #calculating the position of each image to be put into the correct table cell
            cell_row = (piccount/total_column + 1)

            #we are formatting the style of each cell
            cell_range= table.Cell(cell_row, cell_column).Range
            cell_range.ParagraphFormat.LineSpacingRule = win32.constants.wdLineSpaceSingle
            cell_range.ParagraphFormat.SpaceBefore = 0
            cell_range.ParagraphFormat.SpaceAfter = 3

            #this is where we are going to insert the images
            current_pic = cell_range.InlineShapes.AddPicture(os.path.join(os.path.abspath(my_dir), filename))
            width, height = (frame_max_height*frame_max_width/frame_max_height, frame_max_height)

            #changing the size of each image to fit the table cell
            current_pic.Height= height
            current_pic.Width= width

            #putting a name underneath each image which can be handy
            table.Cell(cell_row, cell_column).Range.InsertAfter("\n"+filename)
        else: continue
