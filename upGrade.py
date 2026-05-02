from requests import get
from bs4 import BeautifulSoup


def upGrade():
    try:
        # 获取最新版本号
        response = get('https://gitee.com/cxmStudio/fixWord/raw/main/version_info.json', timeout=10)
        if response.status_code == 200:
            version_info = response.json()
            return version_info
        else:
            return ""
    except Exception as e:
        return ""
