# -*- coding: utf-8 -*-
"""
睡sleepp的店 · 每周数据更新（合并 + 看板）
- 把新一周的千帆导出（如 每周订单数据/7.27-8.2.xlsx）合并进累计历史，去重(按订单号,保留最新状态)
- 写回 累计订单数据.xlsx（数据记录与分析 文件夹），再调用 weekly_dashboard.py 重新生成看板
- 用法: python update_weekly.py <新周数据.xlsx>
"""
import os, sys, subprocess
from weekly_dashboard import read_orders, WS, amt_col, pay_col

CUM = os.path.join(WS, "累计订单数据.xlsx")
ORIG = os.path.join(WS, "截至7.30数据.xlsx")


def main():
    if len(sys.argv) < 2 or not os.path.exists(sys.argv[1]):
        print("用法: python update_weekly.py <新周数据.xlsx>")
        sys.exit(1)
    new_path = sys.argv[1]

    new = read_orders(new_path)
    new = new[new["订单号"].notna()].copy()
    print("新文件订单行: %d" % len(new))

    # 累计源：优先已有累计文件，否则用原始全量导出
    if os.path.exists(CUM):
        base = read_orders(CUM)
        print("累计源: 累计订单数据.xlsx (%d 行)" % len(base))
    elif os.path.exists(ORIG):
        base = read_orders(ORIG)
        print("累计源: 截至7.30数据.xlsx (%d 行)" % len(base))
    else:
        base = new.iloc[0:0]
        print("累计源: 无，仅用新文件")

    base = base[base["订单号"].notna()].copy() if len(base) else base

    # 合并策略：一个订单可能含多个 SKU 行，绝不能用 drop_duplicates(订单号)（会压掉多行丢金额）。
    # 正确做法：找出新旧重叠的订单号，把旧数据里这些订单号的【所有行】删掉，整段换成新导出的行（多行结构保留）。
    import pandas as pd
    overlap = set(base["订单号"]) & set(new["订单号"])
    if len(base):
        base_keep = base[~base["订单号"].isin(overlap)].copy()
    else:
        base_keep = base
    before = len(base_keep) + len(new)
    merged = pd.concat([base_keep, new], ignore_index=True)
    print("累计保留 %d 行 + 新文件 %d 行（重叠订单号 %d，用新导出覆盖旧状态）-> 合并 %d 行"
          % (len(base_keep), len(new), len(overlap), len(merged)))

    # 写回累计文件（pandas 标准格式，数值为真正数字）
    import pandas as pd
    with pd.ExcelWriter(CUM, engine="openpyxl") as w:
        merged.to_excel(w, sheet_name="包裹详情", index=False)
    print("已写累计文件:", CUM)

    # 重新生成看板
    rc = subprocess.run([sys.executable, os.path.join(WS, "weekly_dashboard.py"), CUM], check=False)
    sys.exit(rc.returncode)


if __name__ == "__main__":
    main()
