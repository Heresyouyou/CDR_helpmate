# -*- coding: utf-8 -*-
"""发布 CDR 助手新版本到两台站点（kust.work 内网 / sjck.kust.work 公网）。

做的事只有三件：
  1. 从 cdr_defaults.pyw 读 APP_VERSION（保证"发布号 = 包里的号"）；
  2. 把 dist\\CDR助手.exe 复制成发布目录里的 cdr-helper.exe（ASCII 文件名，
     避免 URL 编码问题）；
  3. 写 cdr.json（url 用相对路径，服务端按请求 Host 补绝对地址）。

用法：
    python publish.py "1. 修了什么\\n2. 加了什么"
    python publish.py "..." --dry-run           # 只看不写
    python publish.py "..." --dir D:\\其它目录   # 换发布目录

发布后两个地址立刻生效（服务端每次请求都读盘，不用重启 Flask）：
    https://kust.work/cdr
    https://sjck.kust.work/cdr
"""
import argparse
import json
import os
import re
import shutil
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
SRC_PY = os.path.join(BASE, "cdr_defaults.pyw")
EXE = os.path.join(BASE, "dist", "CDR助手.exe")
DEFAULT_DIR = r"D:\PostNews\cdr_release"


def source_version():
    with open(SRC_PY, encoding="utf-8") as f:
        m = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', f.read(), re.M)
    return m.group(1) if m else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("notes", nargs="?", default="")
    ap.add_argument("--dir", default=DEFAULT_DIR)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    ver = source_version()
    if not ver:
        sys.exit("读不到 APP_VERSION（%s）" % SRC_PY)
    if not os.path.isfile(EXE):
        sys.exit("找不到打包产物：%s" % EXE)

    src_ver_mtime = os.path.getmtime(SRC_PY)
    exe_mtime = os.path.getmtime(EXE)
    print("版本号      : %s" % ver)
    print("安装包      : %s (%.1f MB)" % (EXE, os.path.getsize(EXE) / 1048576.0))
    print("发布目录    : %s" % a.dir)
    if exe_mtime < src_ver_mtime - 2:
        print("!! 警告：exe 比源码旧，先 pyinstaller 再发布")

    payload = {"version": ver, "url": "/cdr/cdr-helper.exe"}
    if a.notes:
        payload["notes"] = a.notes.replace("\\n", "\n")
    print("cdr.json    : %s" % json.dumps(payload, ensure_ascii=False))

    if a.dry_run:
        print("（dry-run，不写盘）")
        return

    os.makedirs(a.dir, exist_ok=True)
    dst_exe = os.path.join(a.dir, "cdr-helper.exe")
    shutil.copy2(EXE, dst_exe)
    with open(os.path.join(a.dir, "cdr.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print("已写入      : %s + cdr.json" % dst_exe)
    print("验证：")
    print('  curl -s https://kust.work/cdr')
    print('  curl -s https://sjck.kust.work/cdr')


if __name__ == "__main__":
    main()
