[app]
title = Yemen Telecom App
package.name = yementecomapp
package.domain = org.yementecom
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx,pdf
version = 0.1
requirements = python3,kivy,flask,pandas,matplotlib,openpyxl,reportlab,arabic-reshaper,python-bidi,python-docx,werkzeug,pillow
orientation = portrait
fullscreen = 0

android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.ndk = 25b
android.archs = arm64-v8a, armeabi-v7a

[buildozer]
log_level = 2
warn_on_root = 1