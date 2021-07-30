from aip import AipNlp
import pandas as pd

# 调用百度的ai接口进行情感分析
def isPostive(text):
    APP_ID = '24359968'
    API_KEY = '9SbGCCXdhu2GvEGo6WS3qVLV'
    SECRET_KEY = 'x9EUbsSkg9aX2bBlBdnDAU2Fl4RTAnAX'

    client = AipNlp(APP_ID, API_KEY, SECRET_KEY)
    try:
        if client.sentimentClassify(text)['items'][0]['positive_prob'] > 0.5:
            return "积极"
        else:
            return "消极"
    except:
        return "积极"


if __name__ == '__main__':
    # 读取文件，注意修改文件路径
    file_path = 'weibodata/高考.xls'
    data = pd.read_excel(file_path)
    moods = []
    count = 1
    for i in data['微博内容']:
        moods.append(isPostive(i))
        count += 1
        print("目前分析到：" + str(count))
    data['情感倾向'] = pd.Series(moods)

    # 此处为覆盖保存
    data.to_excel(file_path)
    print("分析完成，已保存")


