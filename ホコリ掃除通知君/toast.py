#ホコリ掃除を通知してくれるスクリプト
from plyer import notification

notification.notify(
    title='通知だよ',
    message='掃除機のホコリ掃除する日だよ',
    timeout=3600
    # app_icon='C:/Users/taked/Desktop/python/ホコリ掃除通知君/c.ico'
    )