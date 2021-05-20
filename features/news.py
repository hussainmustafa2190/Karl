import requests
import json



def get_news():
    url = 'https://newsapi.org/v2/top-headlines?country=in&apiKey=a74adaf566194b479ef9842968cda207'
    news = requests.get(url).text
    news_dict = json.loads(news)
    articles = news_dict['articles']
    try:

        return articles
    except:
        return False


def getNewsUrl():
    return 'http://newsapi.org/v2/top-headlines?sources=the-times-of-india&apiKey=ae5ccbe2006a4debbe6424d7e4b569ec'
