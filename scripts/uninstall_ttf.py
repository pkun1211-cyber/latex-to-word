# -*- coding: utf-8 -*-
r"""完整卸载用户级 TrueType 版 LM Roman。

只删 install_ttf.py 本次装进去的文件（existed=False 的项），
安装前就已存在的同名文件一律保留。注册表项按清单精确删除。

用法: python uninstall_ttf.py
"""
import os, sys, io, json, winreg
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import l2wconf as C

MAP = C.tmp('install_ttf_map.json')


def main():
    if not os.path.exists(MAP):
        print('未找到安装清单: %s' % MAP)
        print('说明字体尚未通过 install_ttf.py 安装，无需卸载。')
        return
    plan = json.load(open(MAP, encoding='utf-8'))
    k = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER,
                           r'Software\Microsoft\Windows NT\CurrentVersion\Fonts',
                           0, winreg.KEY_SET_VALUE)
    nr = nf = 0
    for it in plan:
        try:
            winreg.DeleteValue(k, it['reg']); nr += 1
        except FileNotFoundError:
            pass
    winreg.CloseKey(k)
    for it in plan:
        if it['existed']:
            continue
        if os.path.exists(it['dst']):
            os.remove(it['dst']); nf += 1
    print(f'卸载完成: 注册表 {nr} 项, 文件 {nf} 个')
    print('Word 需重启才会释放字体。')


if __name__ == '__main__':
    main()
