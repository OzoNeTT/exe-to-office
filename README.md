# 🎁 Exe To xlsm/docm

Executable file injection to Office documents: `.xlsm`, `.docm`

## 📄 Description

Simple python script to create xlsm or docm dropper of given executable file. You can change code or expand it for another Office extensions

## 📺 How to use

```console
python3 exe-to-office.py -i input_file.exe --xlsm output_document.xlsm
python3 exe-to-office.py -i input_file.exe --docm
```

## ❗ Notice 

This script can be used to deliver potential dangerous software, so use it only for educational purposes or for self-practicing 

- Do not distribute
- Do not use for malicious purposes

## 👀 Supported formats:

- xlsm - working perfect (need macro eneble on guest)
- docm - current test shows, that macro is not running on document opening

# 💎 Version  0.2 💎
