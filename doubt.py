import requests
from bs4 import BeautifulSoup
import xlwt
import time

book=xlwt.Workbook(encoding='utf-8',style_compression=0)

sheet=book.add_sheet('豆瓣电影Top250',cell_overwrite_ok=True)
sheet.write(0,0,'名称')
sheet.write(0,1,'图片')
sheet.write(0,2,'排名')
sheet.write(0,3,'评分')
sheet.write(0,4,'导演')
sheet.write(0,5,'主演')
sheet.write(0,6,'年份')
sheet.write(0,7,'地区')
sheet.write(0,8,'类型')
sheet.write(0,9,'评论')

n=1

def save_to_excel(soup):
    list = soup.find(class_='grid_view').find_all('li')

    for item in list:
        item_name = item.find(class_='title').string
        item_img = item.find('a').find('img').get('src')
        item_index = item.find(class_='').string
        item_score = item.find(class_='rating_num').string
        bodydiv25 = item.find(class_='bd').text
        body = bodydiv25.strip().replace("\n", "")
        twocontent = body.split("               ")
        threecontent = twocontent[0].split(":")
        if len(threecontent) == 3:
            item_daoyan=threecontent[1].strip("主演 ").replace("\xa0", "")
            item_zhuyan=threecontent[2].strip().replace("\xa0", "")
        elif len(threecontent) == 2:
            item_daoyan=threecontent[1].strip().replace("\xa0", "")
            item_zhuyan='~'
        else:
            item_daoyan="~"
            item_zhuyan="~"
        if len(twocontent) <= 1:
                    item_year="~"
                    item_area="~"
                    item_type="~"
        else:
                        right_three_content = twocontent[1].strip().replace("\xa0", "").split("/")
                        item_year=right_three_content[0]
                        item_area=right_three_content[1]
                        item_type=right_three_content[2]
        if(item.find(class_='inq')!=None):
            item_quote = item.find(class_='inq').string
            
        global n

        sheet.write(n,0,item_name)
        sheet.write(n,1,item_img)
        sheet.write(n,2,item_index)
        sheet.write(n,3,item_score)
        sheet.write(n,4,item_daoyan)
        sheet.write(n,5,item_zhuyan)
        sheet.write(n,6,item_year)
        sheet.write(n,7,item_area)
        sheet.write(n,8,item_type)
        sheet.write(n,9,item_quote)
        
        n = n + 1
  
  def main(page):
   headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36'}

   url = 'https://movie.douban.com/top250?start='+ str(page*25)+'&filter='
   html = requests.get(url,headers = headers).text
   soup = BeautifulSoup(html, 'lxml')
   save_to_excel(soup)
   
 if __name__ == '__main__':
    start = time.time()
    for i in range(0, 10):
        main(i)
        i=i+1
    end = time.time()
    print("完成时间: %f s" % (end - start))  
 
 book.save(u'豆瓣最受欢迎的250部电影.csv')
 
