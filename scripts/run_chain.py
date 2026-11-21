# -*- coding: utf-8 -*-
r"""统一入口：按阶段跑 LaTeX→Word 转换链。

阶段划分
--------
  font    装 LM Roman 字体（一次性）
  base    提取 LaTeX PDF 行基准
  pre     预处理 tex + 建模板 + Pandoc 转换
  post    OOXML 后处理 + 精修链
  hard    逐行硬对齐（插 br + 字距）
  iter    闭环校正（导出 PDF → 实测 → 反馈），可指定轮数
  ref     文献区字距精校
  check   只做验收对比

用法
----
  export L2W_CONF=/path/l2w.json
  python run_chain.py base
  python run_chain.py pre post
  python run_chain.py hard
  python run_chain.py iter --rounds 3
  python run_chain.py check
  python run_chain.py all --level C      # 全链（C=逐行硬对齐）

安全
----
每个阶段前自动备份；--dry-run 只打印将执行的命令。
写盘阶段（post/hard/iter/ref）会先关闭占用中的 Word。
"""
import os, sys, io, subprocess, shutil, time
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import l2wconf as C
import docxio as IO

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable


def sh(script, *args, check=True):
    cmd = [PY, os.path.join(HERE, script)] + [str(a) for a in args]
    print('\n$ %s %s' % (script, ' '.join(str(a) for a in args)))
    r = subprocess.run(cmd, cwd=HERE, capture_output=True, text=True,
                       encoding='utf-8', errors='replace')
    if r.stdout:
        print(r.stdout.rstrip())
    if r.returncode != 0:
        print('!! 失败:', (r.stderr or '').strip()[-800:])
        if check:
            raise SystemExit('%s 失败，链路中止' % script)
    return r


def stage_font():
    sh('install_ttf.py')
    sh('audit_lmfont.py', check=False)
    print('\n>>> 请重启 Word 后再继续后续阶段。')


def stage_base():
    if not C.TEXPDF or not os.path.exists(C.TEXPDF):
        raise SystemExit('未配置 texpdf（LaTeX 编译产物）。先编译 tex 得到 PDF。')
    sh('pdflines2.py', C.TEXPDF, C.TEX_TAG)


def stage_pre():
    if not C.TEX or not os.path.exists(C.TEX):
        raise SystemExit('未配置 tex 源文件。')
    sh('preprocess1120.py')
    base = C.tmp('reference_base.docx')
    if not os.path.exists(base):
        print('\n$ pandoc --print-default-data-file reference.docx')
        r = subprocess.run(['pandoc', '--print-default-data-file', 'reference.docx'],
                           capture_output=True)
        if r.returncode != 0:
            raise SystemExit('pandoc 不可用，请先安装 pandoc。')
        open(base, 'wb').write(r.stdout)
    sh('build_reference.py')
    raw = C.tmp('pandoc_raw.docx')
    cmd = ['pandoc', C.tmp('pre.tex'), '-o', raw,
           '--reference-doc=' + C.tmp('reference_custom.docx'),
           '--from=latex', '--to=docx']
    print('\n$ ' + ' '.join(cmd))
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding='utf-8', errors='replace')
    if r.returncode != 0:
        raise SystemExit('pandoc 转换失败: ' + (r.stderr or '')[-800:])
    if r.stderr:
        print('pandoc 警告:', r.stderr.strip()[:600])
    print('已生成', raw)


def stage_post():
    IO.close_word()
    sh('postprocess1120.py')


def stage_hard():
    IO.close_word()
    sh('hard_align.py', check=False)          # 配对计划
    sh('hard_apply.py', '--apply')
    sh('hard_spacing.py', '--apply')


def stage_iter(rounds=3):
    for k in range(1, rounds + 1):
        tag = 'it%d' % k
        IO.close_word()
        sh('export_cmp.py', C.require_docx(), tag)
        sh('hard_iter.py', tag, '--apply', check=False)
    print('\n>>> 闭环完成。若行偏差回升，用 %s 下的 bak_before_* 回滚。' % C.TMP)


def stage_ref():
    IO.close_word()
    sh('export_cmp.py', C.require_docx(), 'refbase', check=False)
    sh('fix_ref3.py', '--apply', check=False)
    sh('export_cmp.py', C.require_docx(), 'refdone', check=False)


def stage_check():
    tag = 'chk%d' % int(time.time() % 10000)
    IO.close_word()
    sh('export_cmp.py', C.require_docx(), tag)
    sh('cmp_pages.py', C.TEX_TAG, tag, check=False)


STAGES = {'font': stage_font, 'base': stage_base, 'pre': stage_pre,
          'post': stage_post, 'hard': stage_hard, 'iter': stage_iter,
          'ref': stage_ref, 'check': stage_check}
ORDER_C = ['base', 'pre', 'post', 'hard', 'iter', 'ref', 'check']
ORDER_B = ['base', 'pre', 'post', 'check']


def main():
    a = sys.argv[1:]
    if not a:
        print(__doc__)
        print(C.describe())
        return
    dry = '--dry-run' in a
    rounds = 3
    if '--rounds' in a:
        rounds = int(a[a.index('--rounds') + 1])
    names = [x for x in a if not x.startswith('--')
             and x not in (str(rounds), 'B', 'C')]
    if 'all' in names:
        lv = 'C'
        if '--level' in a:
            lv = a[a.index('--level') + 1].upper()
        names = ORDER_C if lv == 'C' else ORDER_B
        print('全链档位 %s: %s' % (lv, ' → '.join(names)))
        if lv == 'C':
            print('注意：C 档会把段落变成硬排版，之后改文字需从 tex 重跑全链。')
    for n in names:
        if n not in STAGES:
            raise SystemExit('未知阶段: %s（可选 %s）' % (n, '/'.join(STAGES)))
    if dry:
        print('dry-run，将执行:', ' → '.join(names))
        return
    # 全链前总备份
    if C.DOCX and os.path.exists(C.DOCX):
        b = C.tmp('bak_chain_start.docx')
        if not os.path.exists(b):
            shutil.copy2(C.DOCX, b)
            print('链路起点备份:', b)
    for n in names:
        print('\n' + '=' * 60 + '\n阶段: %s\n' % n + '=' * 60)
        if n == 'iter':
            stage_iter(rounds)
        else:
            STAGES[n]()
    print('\n全部阶段完成。')


if __name__ == '__main__':
    main()
