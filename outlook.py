import requests
import json
from bs4 import BeautifulSoup 

# Disable SSL warning
import urllib3
urllib3.disable_warnings()

# --------------------------------------------------------------------------------

# Config
domain URL = ""
URL = "https://" + domainURL + "/owa/"
pwdURL = "https://" + domainURL + "/ecp/PersonalSettings/Password.aspx"
changepwdURL = "https:/" + domainURL + "/ecp/DDI/DDIService.svc/SetObject?schema=PasswordService&msExchEcpCanary="

header = {'User-Agent' : 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/537.36'}

user = ""
domain = ""
username = user + "@" + domain
password = ""
newPassword = ""

# --------------------------------------------------------------------------------

# Establish Session
s = requests.Session()

# GET: Login page
s.get(URL, headers= header, verify= False)

login_data= {
	"destination": URL,
	"flags": 4,
	"forcedownlevel": 0,
	"username": username,
	"password": password,
	"passwordText": "",
	"isUtf8": 1
}

# POST: Login
s.post(URL + "auth.owa", data= login_data, verify= False)

# --------------------------------------------------------------------------------

# GET: Change password page
r = s.get(pwdURL, headers= header, verify= False)

# Parse required data
dom = BeautifulSoup(r.text, "lxml")
__VIEWSTATE = dom.find(id="__VIEWSTATE").get("value")
ecpCanary = dom.find(id="ecpCanary").get("value")

temp = dom.find(id="ResultPanePlaceHolder_ctl00_ctl02_ctl01_ChangePassword").get("vm-preloadresults")
i = temp.index("RawIdentity")
t = temp[i+11:].replace("&quot;", "")
RawIdentity = t[3:37]

changepwdURL += ecpCanary

changePwd_data = {
	"identity": {
		"__type": "Identity:ECP",
		"DisplayName": user,
		"RawIdentity": RawIdentity
		},
		"properties": {
			"Parameters": {
				"__type":"JsonDictionaryOfanyType:#Microsoft.Exchange.Management.ControlPanel",
				"PlainPassword": newPassword,
				"PlainOldPassword": password
			}
		}
}

# POST: Change password
s.post(changepwdURL, data= json.dumps(changePwd_data), headers={'content-type':'application/json'}  ,verify= False)

