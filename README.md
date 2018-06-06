# Introduction
Change Outlook **2013** Web App's password using python crawler.

# Requirements
- Python 3.5.2 or later

# Install
- For Linux User
```bash
$ pip install pipenv
$ git clone https://github.com/Oschangkai/outlook-pwd.git
$ cd outlook-pwd
$ pipenv install
$ pipenv run python outlook.py
```
- For Windows Visual Studio User
```cmd
> "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\Scripts\pip.exe" install pipenv
> "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\Scripts\pipenv.exe" run python outlook.py
```

# Build
- For Linux User
```bash
$ cd outlook-pwd
$ pipenv run pyinstaller -F outlook.py --clean
```

- For Windows Visual Studio User
```cmd
> cd outlook-pwd
> "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\Scripts\pipenv.exe" run pyinstaller -F outlook.py --clean
```