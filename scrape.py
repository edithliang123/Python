import requests
from bs4 import BeautifulSoup
import json
import xlwings
import os
import math

count = 0

def write_workbook_projectdata(sheetName):
    row_number = 3
    with open(os.path.expanduser('~/Desktop/filter_results.txt')) as f:
        for result_list in range(1, count+1):
            project_info = []
            project_url = f.readline()
            project_content = requests.get(project_url)
            soup = BeautifulSoup(project_content.text, 'html.parser')
            project_name = soup.find('h2', attrs={'class':'type-28 type-24-md soft-black mb1 project-name'}).text
            project_founder = soup.find('img', attrs={'class':'border-box radius100p bg-grey-400 w7 h7 shrink0 mr2'})['alt']
            project_category = soup.find('span', attrs={'class':'ml1'}).text
            project_founder_location = soup.find('span').findNext('span', attrs={'class':'ml1'}).text.split(', ')
            project_founder_city = project_founder_location[0]
            try:
                project_founder_state = project_founder_location[1]
            except IndexError:
                row_number += 1
                continue;
    
            data = [
                json.loads(i["data-initial"])
                for i in soup.find_all("div")
                if i.get("data-initial")
                ]
            for i in data:
                project_pledged = i.get("project").get("pledged").get("amount")
                project_goal = i.get("project").get("goal").get("amount")
                project_backers = i.get("project").get("backersCount")
                campaign_created_count = i.get("project").get("creator").get("launchedProjects").get("totalCount")
                creator_url = i.get("project").get("creator").get("url")+"/created"
    
            project_funding_difference = str(float(project_pledged) - float(project_goal))
            project_success = str(0) if (float(project_pledged) - float(project_goal)) < 0 else str(1)
            project_has_FAQ = str(1) if int(soup.find('a', attrs={'data-analytics':'faq'})['emoji-data']) > 0 else str(0)
            project_updates_count = soup.find('a', attrs={'data-analytics':'updates'})['emoji-data']
            project_comments_count = soup.find('a', attrs={'data-analytics':'comments'})['data-comments-count']
    
            #Code below are all for finding previous_success_count
            creator_content = requests.get(creator_url)
            soup2 = BeautifulSoup(creator_content.text, 'html.parser')
            data2 = [
                json.loads(j["data-projects"])
                for j in soup2.find_all("div")
                if j.get("data-projects")
                ]
            previous_success_count = 0
            previous_suspended_count = 0
            previous_canceled_count = 0
            for j in data2:
                for x in range(0, len(j)):
                    if j[x].get("state") == 'canceled':
                        previous_canceled_count += 1
                    elif j[x].get("state") == 'suspended':
                        previous_suspended_count += 1
                    elif j[x].get("state") == 'success':
                        previous_suspended_count += 1 
            project_info.extend((project_name,project_category,project_founder,project_founder_city,project_founder_state,project_goal,project_pledged,project_backers,project_funding_difference,project_success,project_has_FAQ,project_updates_count,project_comments_count,campaign_created_count,str(previous_success_count),str(previous_suspended_count),str(previous_canceled_count)))
            write_workbook_data(project_info, row_number,sheetName)
            row_number += 1


def write_workbook_data(project_info, row_number,sheetName):
    for i in range(2, 25):
        if i == 2:
            sheetName.range((row_number, i)).value = project_info[0]
        elif i == 3:
            sheetName.range((row_number, i)).value = project_info[1]
        elif i == 4:
            sheetName.range((row_number, i)).value = project_info[2]
        elif i == 5:
            sheetName.range((row_number, i)).value = project_info[3]
        elif i == 6:
            sheetName.range((row_number, i)).value = project_info[4]
        elif i == 7:
            sheetName.range((row_number, i)).value = project_info[5]
        elif i == 8:
            sheetName.range((row_number, i)).value = project_info[6]
        elif i == 9:
            sheetName.range((row_number, i)).value = project_info[7]
        elif i == 10:
            sheetName.range((row_number, i)).value = project_info[8]
        elif i == 11:
            sheetName.range((row_number, i)).value = project_info[9]
        elif i == 15:
            sheetName.range((row_number, i)).value = project_info[10]
        elif i == 16:
            sheetName.range((row_number, i)).value = project_info[11]
        elif i == 17:
            sheetName.range((row_number, i)).value = project_info[12]
        elif i == 19:
            sheetName.range((row_number, i)).value = project_info[13]
        elif i == 20:
            sheetName.range((row_number, i)).value = project_info[14]
        elif i == 21:
            sheetName.range((row_number, i)).value = project_info[15]
        elif i == 22:
            sheetName.range((row_number, i)).value = project_info[16]        


#scrape urls of results to workbook
def write_workbook_link(sheetName):
    with open(os.path.expanduser('~/Desktop/filter_results.txt')) as f:
        for i in range(3, count+2):
            sheetName.range((i, 1)).value = f.readline()


def collect_filter_urls(page_number):
    global count
    website_url = 'https://www.kickstarter.com/discover/advanced?state=live&category_id=16&woe_id=23424977&sort=newest&seed=2713308&page='+str(page_number)
    website_content = requests.get(website_url)
    soup = BeautifulSoup(website_content.text, 'html.parser')

    data = [
        (json.loads(i["data-project"]), i["data-ref"])
        for i in soup.find_all("div")
        if i.get("data-project")
        ]

    for i in data:
        count += 1
        f.write(f'{i[0]["urls"]["web"]["project"]}?ref={i[1]}\n')


if __name__ == '__main__':
    #find page_number
    website_url = 'https://www.kickstarter.com/discover/advanced?state=live&category_id=16&woe_id=23424977&sort=newest&seed=2713308'
    website_content = requests.get(website_url)
    soup = BeautifulSoup(website_content.text, 'html.parser')
    total_projects = soup.find('b',attrs={'class':'count ksr-green-500'}).text.split(' projects')
    page_number = math.ceil(int(total_projects[0])/12)
    
    #create filter_result file for url lists
    f = open('filter_results.txt', 'w')
    for i in range(1, page_number+1):
        collect_filter_urls(i)
    f.close()
    
    #write to workbook
    workBook = xlwings.Book('KickStarter Campaign Data.xlsx');
    sheetName = workBook.sheets[-1]
    write_workbook_link(sheetName)
    write_workbook_projectdata(sheetName)
    workBook.save()
    workBook.close()    
    print('Data has been collected')
    