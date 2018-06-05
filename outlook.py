import requests
import json
from bs4 import BeautifulSoup 

# Disable SSL warning
import urllib3
urllib3.disable_warnings()

# Argumnet
from argparse import ArgumentParser

# --------------------------------------------------------------------------------

# Setting Global Values
def setVals(args):
	global domainURL
	global URL
	global pwdURL
	global changepwdURL

	global email
	global user
	global password
	global newPassword

	global header

	domainURL = args.domain_url
	URL = "https://" + domainURL + "/owa/"
	pwdURL = "https://" + domainURL + "/ecp/PersonalSettings/Password.aspx"
	changepwdURL = "https://" + domainURL + "/ecp/DDI/DDIService.svc/SetObject?schema=PasswordService&msExchEcpCanary="

	email = args.email
	user = email[email.index("@")+1:]
	password = args.old_password
	newPassword = args.new_password

	header = {'User-Agent' : 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/537.36'}


# --------------------------------------------------------------------------------

def test():
	print(pwdURL, user, newPassword)

def login():
	login_data= {
		"destination": URL,
		"flags": 4,
		"forcedownlevel": 0,
		"username": email,
		"password": password,
		"passwordText": "",
		"isUtf8": 1
	}

	# Establish Session
	global s
	s = requests.Session()

	# GET: Login page
	s.get(URL, headers= header, verify= False)

	# POST: Login
	s.post(URL + "auth.owa", data= login_data, verify= False)

def changePwd():
	# GET: Change password page
	r = s.get(pwdURL, headers= header, verify= False)

	# Parse required data
	dom = BeautifulSoup(r.text, "lxml")
	__VIEWSTATE = dom.find(id="__VIEWSTATE").get("value")
	ecpCanary = dom.find(id="ecpCanary").get("value")

	# Cleanup and get RawIdentity
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
	s.post(changepwdURL, data= json.dumps(changePwd_data), headers={'content-type':'application/json'}, verify= False)


if __name__ == "__main__":
	p = ArgumentParser(prog="famigrp.exe", description="從 Outlook 網頁更改 domain 密碼", epilog="本程式必須連網，請注意網路連線！")#, add_help=False)
	required = p.add_argument_group("必填訊息", "")
	optional = p.add_argument_group("選填參數", "")

	required.add_argument('-m','--email', help="Outlook 登入的電子信箱", required= True)
	required.add_argument('-opwd','--old-password', help="舊密碼", required= True)
	required.add_argument('-pwd','--new-password', help="新密碼", required= True)
	
	optional.add_argument('-d','--domain', dest='domain_url', default="famigrp.com.tw", help='設定 Outlook domain URL (預設: famigrp.com.tw)')
	# p.print_help()

	args = p.parse_args()
	setVals(args)
	login()
	changePwd()
# "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\Scripts\pyinstaller.exe" -F myscript.spec --clean