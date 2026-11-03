# -*- coding: utf-8 -*-
"""导出版式配置 / 查看配置。

用法:
    python config_tool.py dump --paper a4 -o 我的版式.json
    python config_tool.py dump --paper letter            # 直接打印到屏幕
    python config_tool.py show 我的版式.json              # 打印合并后的完整数值

导出的 JSON 只写要改的项即可，其余继承 `base` 预设。改完把它传给:
    python new_paper.py 稿件.txt -o 论文.docx --config 我的版式.json
    python apply_format.py 原稿.docx -o 排好版.docx --config 我的版式.json
    python make_template.py --config 我的版式.json --out 模板.docx
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import paper_style as ps                                     # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="导出版式配置 / 查看配置")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("dump", help="把内置预设导出成可编辑的 JSON")
    p.add_argument("--paper", default=ps.DEFAULT_PROFILE,
                   choices=list(ps.PROFILES), help="版式预设（默认 letter）")
    p.add_argument("--config", default=None,
                   help="基于已有配置导出（与 --paper 的 base 叠加）")
    p.add_argument("-o", "--out", default=None,
                   help="输出 JSON 路径；不给就打印到屏幕")

    p = sub.add_parser("show", help="读取配置并打印合并后的完整数值")
    p.add_argument("config", help="JSON 配置文件")
    p.add_argument("-o", "--out", default=None,
                   help="输出 JSON 路径；不给就打印到屏幕")

    args = ap.parse_args()

    if args.cmd == "dump":
        key = args.paper
        if args.config:
            try:
                key = ps.apply_config(args.config, base=args.paper)
            except Exception as e:
                sys.exit(f"读取版式配置失败：{e}")
        cfg = ps.dump_config(key, args.out)
        if args.out:
            print("已生成配置:", args.out)
        else:
            print(json.dumps(cfg, ensure_ascii=False, indent=2))
        return

    try:
        key = ps.apply_config(args.config)
    except Exception as e:
        sys.exit(f"读取版式配置失败：{e}")
    cfg = ps.dump_config(key, args.out)
    if args.out:
        print("已生成配置:", args.out)
    else:
        print(json.dumps(cfg, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
