import os,time

CLIENT_FOLDER = 'C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson Web App/'
file_list = os.listdir(CLIENT_FOLDER)
dic = {key: time.ctime(os.path.getmtime(
    os.path.join(CLIENT_FOLDER, key))) for key in file_list}

print(dic)