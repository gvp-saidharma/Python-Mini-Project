import os
import sys
from typing_extensions import TypeGuard
from pytube import YouTube
import time
import tkinter
from tkinter import Message, ttk
from urllib.request import Request, urlopen, urlparse
import requests 

# https://www.youtube.com/watch?v=giVqFIoOU_A
# https://www.learningcontainer.com/wp-content/uploads/2020/05/sample-mp4-file.mp4
# https://www.facebook.com/favicon.ico
# https://www.google.com

def getFileName(url):
    headers=requests.head(url).headers
    contenttype = headers.get('content-type')
    a = urlparse(url)
    if "video" in contenttype or "image/png" in contenttype: 
        return os.path.basename(a.path)
    return None

def downloadFileFromYoutubeURL(url):
    download_location = "./"
    print(url)
    try: 
        # object creation using YouTube
        # which was imported in the beginning 
        yt = YouTube(url) 
        
        yt.streams.first().download(download_location)
    except: 
        return "Download Failed, Unexpected Error:" + str(sys.exc_info()[0])
    else:
        return "Download sucess"

def downloadFileFromURL(url):
    file_name = getFileName(url)
    if file_name is None:
        print("unable to download")
        return "unable to download"
    else:
        request = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        rsp = urlopen(request)
        with open(file_name,'wb') as f:
            f.write(rsp.read())
    return "Download sucess"

def entry_fields():
    start_time=time.time()
    download_location = "./"
    msg = ""
    try: 
        url = entry.get()
        if "youtube.com" in url: 
            print("downloadFileFromYoutubeURL")
            msg = downloadFileFromYoutubeURL(url)
        else:
            print("downloadFileFromURL")
            msg = downloadFileFromURL(url)        
    except: 
        msg= "Download Failed, Unexpected Error:" + str(sys.exc_info()[0])
    popup = tkinter.Tk()
    popup.wm_title(msg)
    label = ttk.Label(popup, text=msg)
    label.pack(side="top", fill="x", pady=50)
    B1 = ttk.Button(popup, text="Ok", command = popup.destroy)
    B1.pack()
    popup.mainloop()


master=tkinter.Tk()
master.geometry("1500x1000")
master.wm_title("Download Video")
tkinter.Label(master,text="Enter Video URL: ").grid(row=0)
entry = tkinter.Entry(master)
entry.grid(row=0,column=1)

tkinter.Button(master, text='Download', command=entry_fields,anchor=tkinter.CENTER).grid(row=1, sticky=tkinter.W, pady=4)          
tkinter.mainloop()
print('Task Completed!') 

