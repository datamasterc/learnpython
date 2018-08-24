

import urllib.request
url = r'https://movie.douban.com/top250?start=25&filter='
res = urllib.request.urlopen(url)
html = res.read().decode('utf-8')
print(html)
#获取网页的html