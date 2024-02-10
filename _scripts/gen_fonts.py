from pathlib import Path

p = Path('fonts')

from pywinauto.application import Application


origin_names = ['スーラ', 'ニューシネマ', 'ハミング']
modifed_names = ['LXGW WenKai Screen', '微软雅黑', '霞鹜新晰黑']

sizes = [12, 21, 30, 39]

apps = []

fp = Path('C:/Windows/Fonts/')
print(fp.exists(), list(fp.iterdir()),Path().absolute(), ((Path() / 'AppData/Local/Microsoft/Windows/Fonts').iterdir()))

for origin_name, modifed_name in zip(origin_names, modifed_names):
    for size in sizes:
        app = Application().start('assets/krkrfont.exe') 
        win = app.window()
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
