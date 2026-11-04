[app]

# (str) Title of your application
title = Yemen Telecom App

# (str) Package name
package.name = yementecomapp

# (str) Package domain (needed for android/ios packaging)
package.domain = org.yementecom

# (str) Source code where the main.py live
source.dir = .

# (list) Source files to include (let empty to include all the files)
source.include_exts = py,png,jpg,kv,atlas,xlsx,pdf,ttf

# (list) Source files to exclude
source.exclude_dirs = tests, bin

# (str) Application versioning
version = 0.1

# (list) Application requirements
# تمت إضافة جميع المكتبات التي يحتاجها تطبيقك
requirements = python3,kivy,flask,pandas,matplotlib,openpyxl,reportlab,arabic-reshaper,python-bidi,python-docx,werkzeug,pillow,cython

# (str) Presplash of the Application
# presplash.filename = %(source.dir)s/data/presplash.png

# (str) Icon of the Application
# icon.filename = %(source.dir)s/data/icon.png

# (str) Supported orientation (one of landscape, sensorLandscape, portrait or all)
orientation = portrait

# (bool) Indicate if the application should be fullscreen or not
fullscreen = 0


#
# Android specific
#

# (list) Permissions
android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE

# (int) Target Android API, should be as high as possible.
android.api = 33

# (int) Minimum API your APK will support.
android.minapi = 21

# (int) Android SDK version to use
android.sdk = 33

# (str) Android NDK version to use
android.ndk = 25b

# (int) Android NDK API to use. This is the minimum API your app will support.
android.ndk_api = 21

# (bool) Use --private data storage (True) or --dir public storage (False)
android.private_storage = True

# (str) Android ABI:
android.archs = arm64-v8a, armeabi-v7a

# (str) Kotlin version to use
android.kotlin_version = 1.9.10

#
# Buildozer specific
#

[buildozer]

# (int) Log level (0 = error only, 1 = info, 2 = debug (with command output))
log_level = 2

# (int) Display warning if buildozer is run as root
warn_on_root = 1
