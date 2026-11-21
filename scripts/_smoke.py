# -*- coding: utf-8 -*-
r"""冒烟自检：逐个 import 脚本，验证「可加载」且「不改用户文件」。

为什么必须校验文件哈希
----------------------
早期版本的 refine_layout / fix_fontsize / fix_justify / fix_settings 等
把主逻辑写在模块顶层，import 即执行并写盘。一次自检就把 326489 字节的
成品改成 318668，行偏差从 35 退到 56。现在这些脚本都有 main() 守卫，
本工具用 SHA256 前后比对来**证明** import 无副作用。

用法: L2W_CONF=... python _smoke.py
"""
import os, sys, io, subprocess, json, hashlib
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import l2wconf as C

HERE = os.path.dirname(os.path.abspath(__file__))
SKIP = {'l2wconf.py', 'docxio.py', '_smoke.py', 'run_chain.py'}


def sha(p):
    if not p or not os.path.exists(p):
        return None
    return hashlib.sha256(open(p, 'rb').read()).hexdigest()


def main():
    docx = C.DOCX
    before = sha(docx)
    if before:
        print('目标 docx SHA256(前): %s  %d 字节'
              % (before[:16], os.path.getsize(docx)))
    else:
        print('未配置 docx，只测可加载性')

    mods = sorted(f[:-3] for f in os.listdir(HERE)
                  if f.endswith('.py') and f not in SKIP and not f.startswith('_'))
    ok, bad = [], []
    for m in mods:
        code = 'import sys; sys.argv=["x","--plan"]; import %s' % m
        r = subprocess.run([sys.executable, '-c', code], cwd=HERE,
                           capture_output=True, text=True, encoding='utf-8',
                           errors='replace', timeout=240)
        if r.returncode == 0:
            ok.append(m)
        else:
            err = (r.stderr or '').strip().split('\n')
            bad.append((m, err[-1][:150] if err else '?'))

    print('\n导入成功 %d/%d' % (len(ok), len(mods)))
    for m, e in bad:
        print('  FAIL %-20s %s' % (m, e))

    after = sha(docx)
    unchanged = (before == after)
    if before and after:
        if unchanged:
            print('\n[OK] 文件完整性校验通过：import 全程未修改目标 docx')
        else:
            print('\n[!!] 目标 docx 被修改！SHA256(后)=%s  %d 字节'
                  % (after[:16], os.path.getsize(docx)))
            print('     说明仍有脚本在 import 时写盘，必须加 main() 守卫。')
    json.dump({'ok': ok, 'bad': bad, 'sha_before': before, 'sha_after': after},
              open(os.path.join(HERE, '_smoke_result.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)
    return 0 if not bad and unchanged else 1


if __name__ == '__main__':
    sys.exit(main())
