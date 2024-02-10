from pathlib import Path
from time import sleep

p = Path('fonts')

from pywinauto.application import Application
import shutil


origin_names = ['スーラ', 'ニューシネマ', 'ハミング']
modifed_names = ['LXGW WenKai Screen', '微软雅黑', '霞鹜新晰黑']

sizes = [12, 21, 30, 39]

apps = []

fp = Path('C:/Windows/Fonts/')
for ft in p.iterdir():
    shutil.move(ft, fp / ft.name)
print(fp.exists(), Path().absolute(), list(Path('C:\\Users').iterdir()))
print(Path.home().absolute())

for origin_name, modifed_name in zip(origin_names, modifed_names):
    for size in sizes:
        app = Application().start('assets/krkrfont.exe') 
        win = app.window()
        sleep(1)
        print(win.TComboBox1.ItemTexts())
        win.TComboBox1.select(modifed_name)
        win.TEdit2.set_text(str(size))
        for i in (1, 3, 4):
            win[f'TCheckBox{i}'].check()
        win.TEdit1.set_text((Path('fonts') / f'{origin_name}{size}.tft').absolute())
        win.TButton2.click()
        apps.append(app)
        break


for app in apps:
    app.TFontMakerWriteForm.wait_not('active', timeout=10*60)
    app.kill()
