def main():
    if not os.path.exists(MAP):
        print('x')
        return
    for it in plan:
        try:
            winreg.DeleteValue(k, it['reg'])
        except FileNotFoundError:
            pass
