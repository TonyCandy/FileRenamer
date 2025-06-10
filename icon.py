import base64

# 读取图标文件并转换为base64
with open('app.ico', 'rb') as icon_file:
    icon_data = base64.b64encode(icon_file.read()).decode('utf-8')

# 创建Python代码
with open('icon_rc.py', 'w') as py_file:
    py_file.write('icon_data = """\n')
    py_file.write(icon_data)
    py_file.write('"""\n')