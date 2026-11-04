[app]
title = Yemen Telecom App
package.name = yementecomapp
package.domain = org.yementecom
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx,pdf,ttf
source.exclude_dirs = tests, bin
version = 0.1
requirements = python3,kivy,flask,pandas,matplotlib,openpyxl,reportlab,arabic-reshaper,python-bidi,python-docx,werkzeug,pillow,cython
orientation = portrait
fullscreen = 0

# Android specific
android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.ndk = 25b
android.ndk_api = 21
android.private_storage = True
android.archs = arm64-v8a, armeabi-v7a

# إجبار Buildozer على استخدام الإصدار 33
android.sdk = 33
android.accept_sdk_license = True
android.build_tools = 33.0.2

[buildozer]
log_level = 2
warn_on_root = 1
