import requests      
from win32com.client import Dispatch
def NewsFromBBC(): 
    main_url = " https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
    open_bbc_page = requests.get(main_url).json() 
    article = open_bbc_page["articles"] 
    results = []   
    for ar in article: 
        results.append(ar["title"]) 
    for i in range(len(results)): 
        print(i + 1, results[i])  
    speak = Dispatch("SAPI.Spvoice") 
    speak.Speak(results)                  
if __name__ == '__main__': 

    NewsFromBBC()  
