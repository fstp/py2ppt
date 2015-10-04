import os
from urllib import urlretrieve
from urlparse import urljoin
from hashlib import md5
from win32com.client import Dispatch
from lxml.html import parse
#import MSO
import MSPPT

tree = parse("http://www.imdb.com/chart/top")
movies = tree.xpath("//td[contains(@class,'titleColumn')]//a")

if not os.path.exists("img"):
    os.makedirs("img")

def filename(movie):
    ''' Filename is the MD5 hash of the title. '''
    name = md5(movie.text.encode("utf8")).hexdigest()
    return os.path.join("img", name + ".jpg")

for movie in movies:
    if os.path.exists(filename(movie)):
        continue

    url = urljoin("http://www.imdb.com/", movie.get("href"))
    tree = parse(url)
    img = tree.find(".//td[@id='img_primary']//img")
    urlretrieve(img.get("src"), filename(movie))

Application = Dispatch("PowerPoint.Application")
Application.Visible = 1
Presentation = Application.Presentations.Add()
Base = Presentation.Slides.Add(1, MSPPT.constants.ppLayoutBlank)

width, height = 28, 41
for i, movie in enumerate(movies):
    x = 10 + width * (i % 25)
    y = 100 + height * (i // 25)
    r = Base.Shapes.AddPicture(
            os.path.abspath(filename(movie)),
            LinkToFile=True,
            SaveWithDocument=False,
            Left=x, Top=y,
            Width=width, Height=height)
