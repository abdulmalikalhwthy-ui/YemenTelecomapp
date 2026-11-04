import threading
import time
import socket
from kivy.app import App
from kivy.uix.webview import WebView
from kivy.clock import Clock
import os
import sys

# تشغيل سيرفر Flask في الخلفية
def run_flask():
    from app import app
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)

class MyApp(App):
    def build(self):
        # بدء السيرفر في Thread منفصل
        threading.Thread(target=run_flask, daemon=True).start()
        
        # الانتظار حتى يبدأ السيرفر
        time.sleep(2) 
        
        # فتح صفحة الويب داخل التطبيق
        webview = WebView(url='http://127.0.0.1:5000')
        return webview

if __name__ == '__main__':
    MyApp().run()