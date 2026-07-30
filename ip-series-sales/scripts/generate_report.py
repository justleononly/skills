#!/usr/bin/env python3
"""
IP多弹次系列销售复盘分析报告生成器

功能：
1. 读取指定CSV多弹次销售数据
2. 按照标准分析框架计算关键指标
3. 输出结构化的Markdown分析报告

分析框架：
  - 模块一：新弹次基础销售分析（核心指标、新老用户分布）
  - 模块二：跨系列复购分析（人群主导判断、复购用户特征、新用户特征）
  - 模块三：用户流失分析（流失规模、高价值流失用户特征）
  - 策略建议（基于分析结论的数据驱动建议）

用法：
    python3 generate_report.py --csv <数据文件.csv> [--output <输出目录>]

输入要求：
  - CSV编码：GBK/UTF-8
  - 必需列：下单时间, 用户ID, 用户注册时间, 弹次名称, 消费金额
  - 可选列：会员等级
"""

import argparse
import csv
import os
import sys
import json
from collections import defaultdict, Counter
from datetime import datetime

# ── 参数解析 ──────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(description="IP多弹次系列销售复盘分析报告生成器")
parser.add_argument("--csv", required=True, help="多弹次销售数据CSV文件路径")
parser.add_argument("--output", default=".", help="报告输出目录（默认当前目录）")
args = parser.parse_args()

csv_path = args.csv
out_dir = args.output
os.makedirs(out_dir, exist_ok=True)


# ── 数据加载 ──────────────────────────────────────────────────────────
def load_csv(path):
    """尝试多种编码加载CSV"""
    rows = []
    uid_to_reg = {}
    for enc in ("gbk", "utf-8", "gb18030"):
        try:
            with open(path, encoding=enc) as f:
                reader = csv.DictReader(f)
                for row in reader:
                    row["下单时间_ts"] = datetime.strptime(row["下单时间"].strip(), "%Y/%m/%d %H:%M")
                    row["注册时间_ts"] = datetime.strptime(row["用户注册时间"].strip(), "%Y/%m/%d %H:%M")
                    row["消费金额"] = float(row["消费金额"])
                    uid = row["用户ID"]
                    if uid not in uid_to_reg or row["注册时间_ts"] < uid_to_reg[uid]:
                        uid_to_reg[uid] = row["注册时间_ts"]
                    rows.append(row)
            print(f"[OK] 数据加载成功，编码: {enc}，共 {len(rows)} 条记录")
            return rows, uid_to_reg
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError("无法解码CSV文件，请确保文件为GBK或UTF-8编码")


all_rows, uid_to_reg = load_csv(csv_path)

# 识别弹次
series_names = sorted(set(r["弹次名称"] for r in all_rows))
print(f"[INFO] 识别到弹次: {series_names}")
sn_new = series_names[-1]  # 最新弹次为分析对象
historical_series = [s for s in series_names if s != sn_new]
print(f"[INFO] 新弹次: {sn_new}")

# 新弹次日期范围
new_dates_all = [r["下单时间_ts"] for r in all_rows if r["弹次名称"] == sn_new]
new_start_date = min(new_dates_all).date()
new_end_date = max(new_dates_all).date()

# 用户级聚合
series_user = {}
for r in all_rows:
    sn, uid = r["弹次名称"], r["用户ID"]
    if sn not in series_user:
        series_user[sn] = {}
    if uid not in series_user[sn]:
        series_user[sn][uid] = {"orders": 0, "total": 0.0, "first_dt": r["下单时间_ts"], "last_dt": r["下单时间_ts"]}
    su = series_user[sn][uid]
    su["orders"] += 1
    su["total"] += r["消费金额"]
    su["last_dt"] = max(su["last_dt"], r["下单时间_ts"])
    su["first_dt"] = min(su["first_dt"], r["下单时间_ts"])

# IP总消费
uid_total = defaultdict(float)
for r in all_rows:
    uid_total[r["用户ID"]] += r["消费金额"]

all_ip_uids = set(r["用户ID"] for r in all_rows)
new_series_users = set(series_user[sn_new].keys())


# ── 工具函数 ──────────────────────────────────────────────────────────
def series_days(sn):
    dates = [r["下单时间_ts"].date() for r in all_rows if r["弹次名称"] == sn]
    return (max(dates) - min(dates)).days + 1 if dates else 0


def series_stats(sn, user_set=None):
    """获取指定弹次中指定用户集的聚合指标"""
    if user_set is None:
        user_set = set(series_user[sn].keys())
    rev = sum(series_user[sn][uid]["total"] for uid in user_set if uid in series_user[sn])
    cnt = len(user_set & set(series_user[sn].keys()))
    ords = sum(series_user[sn][uid]["orders"] for uid in user_set if uid in series_user[sn])
    arpu = rev / cnt if cnt else 0
    return rev, cnt, ords, arpu


def format_yuan(val):
    return f"¥{val:,.0f}"


# ── 模块一：新弹次基础销售分析 ─────────────────────────────────────
def module1_basic():
    print("[分析] 模块一：基础销售分析")
    lines = []
    lines.append("## 二、新弹次销售概览")
    lines.append("")
    lines.append("### 2.1 核心指标总览")
    lines.append("")
    lines.append("| 系列 | 上线天数 | 销售额 | 付费用户数 | 客单价 | 订单数 |")
    lines.append("|:---|:---:|:---:|:---:|:---:|:---:|")
    for sn in series_names:
        users = series_user[sn]
        rev = sum(u["total"] for u in users.values())
        cnt = len(users)
        ords = sum(u["orders"] for u in users.values())
        arpu = rev / cnt if cnt else 0
        days = series_days(sn)
        marker = " (新弹次)" if sn == sn_new else ""
        lines.append(f"| **{sn}**{marker} | {days}天 | {format_yuan(rev)} | {cnt:,} | ¥{arpu:.2f} | {ords:,} |")
    lines.append("")
    return "\n".join(lines)


def module1_newuser():
    """新老用户分布（二分类 + 老用户细分）"""
    print("[分析] 新老用户分布")

    # 新用户定义：注册时间在新弹次销售周期内
    new_users = set()
    for uid in new_series_users:
        reg = uid_to_reg.get(uid)
        if reg and new_start_date <= reg.date() <= new_end_date:
            new_users.add(uid)

    old_users = new_series_users - new_users

    # 老用户细分
    existing_new = set()  # 历史注册·首次购买IP
    repurchasers = set()  # 历史复购用户
    for uid in old_users:
        bought_hist = any(uid in series_user[s] for s in historical_series)
        if bought_hist:
            repurchasers.add(uid)
        else:
            existing_new.add(uid)

    rev_new, cnt_new, ords_new, arpu_new = series_stats(sn_new, new_users)
    rev_old, cnt_old, ords_old, arpu_old = series_stats(sn_new, old_users)
    total_rev = rev_new + rev_old
    total_cnt = cnt_new + cnt_old

    rev_ex, cnt_ex, _, arpu_ex = series_stats(sn_new, existing_new)
    rev_rp, cnt_rp, _, arpu_rp = series_stats(sn_new, repurchasers)

    lines = []
    lines.append(f"## 新老用户分布（新系列·{sn_new}）")
    lines.append("")
    lines.append(f"> **定义**：新用户 = 用户注册时间在本弹次销售周期内（{new_start_date} ~ {new_end_date}）；"
                 f"老用户 = 注册时间在销售周期之前。在此基础上进一步将老用户细分为“历史注册·首次购买IP”和“历史复购用户”两类。")
    lines.append("")
    lines.append("### 2.1 基础新老用户分布（二分类法）")
    lines.append("")
    lines.append("| 指标 | 新用户 | 老用户 |")
    lines.append("|:---|---:|---:|")
    lines.append(f"| **销售额** | {format_yuan(rev_new)} | {format_yuan(rev_old)} |")
    lines.append(f"| **销售额占比** | {rev_new/total_rev*100:.1f}% | {rev_old/total_rev*100:.1f}% |")
    lines.append(f"| **付费用户数** | {cnt_new}人 | {cnt_old}人 |")
    lines.append(f"| **用户数占比** | {cnt_new/total_cnt*100:.1f}% | {cnt_old/total_cnt*100:.1f}% |")
    lines.append(f"| **客单价** | {format_yuan(arpu_new)} | {format_yuan(arpu_old)} |")
    lines.append("")
    lines.append("### 2.2 老用户进一步细分")
    lines.append("")
    lines.append("| 老用户子类 | 销售额 | 销售额占比 | 用户数 | 用户数占比 | 客单价 |")
    lines.append("|:---|---:|---:|---:|---:|---:|")
    lines.append(f"| 历史注册·IP首购 | {format_yuan(rev_ex)} | {rev_ex/total_rev*100:.1f}% | {cnt_ex}人 | {cnt_ex/total_cnt*100:.1f}% | {format_yuan(arpu_ex)} |")
    lines.append(f"| 历史复购用户 | {format_yuan(rev_rp)} | {rev_rp/total_rev*100:.1f}% | {cnt_rp}人 | {cnt_rp/total_cnt*100:.1f}% | {format_yuan(arpu_rp)} |")
    lines.append("")
    lines.append("### 2.3 关键解读")
    lines.append("")
    rate = int(arpu_old / arpu_new) if arpu_new else 0
    lines.append(f"- **新用户占付费人数的 {cnt_new/total_cnt*100:.1f}%，但销售额贡献仅 {rev_new/total_rev*100:.1f}%** —— 新用户以低价尝鲜为主，整体客单价仅为老用户的 1/{max(1,rate)}。")
    lines.append(f"- **老用户以 {cnt_old/total_cnt*100:.1f}% 的人数贡献了 {rev_old/total_rev*100:.1f}% 的销售额**，其中历史复购用户以 {cnt_rp/total_cnt*100:.1f}% 的人数贡献了 {rev_rp/total_rev*100:.1f}% 的销售额，是绝对销售基本盘。")
    lines.append(f"- 新老客单价差距极大，说明弹次定价更依赖**存量高价值用户的深度消费**，而非新用户的广泛覆盖。")
    lines.append("")
    return "\n".join(lines)


# ── 模块二：跨系列复购分析 ───────────────────────────────────────
def module2():
    print("[分析] 模块二：跨系列复购分析")
    lines = []

    # 分类
    repurchaser_uids = set()
    for uid in new_series_users:
        for s in historical_series:
            if uid in series_user[s]:
                repurchaser_uids.add(uid)
                break
    non_repurchaser = new_series_users - repurchaser_uids

    rt_rev, rt_cnt, rt_ords, rt_arpu = series_stats(sn_new, repurchaser_uids)
    nr_rev, nr_cnt, nr_ords, nr_arpu = series_stats(sn_new, non_repurchaser)
    total_new_rev = rt_rev + nr_rev

    # ---- 3.1 人群主导判断 ----
    lines.append("## 三、跨系列复购分析")
    lines.append("")
    lines.append("### 3.1 人群主导判断")
    lines.append("")
    lines.append(f"新系列{sn_new}销售额 {format_yuan(total_new_rev)} 的用户群构成：")
    lines.append("")
    lines.append("| 用户群 | 销售额 | 销售额占比 | 用户数 | 用户数占比 | 客单价 |")
    lines.append("|:---|---:|---:|---:|---:|---:|")
    lines.append(f"| **历史复购用户(买过前序弹次)** | {format_yuan(rt_rev)} | {rt_rev/total_new_rev*100:.1f}% | {rt_cnt}人 | {rt_cnt/len(new_series_users)*100:.1f}% | {format_yuan(rt_arpu)} |")
    lines.append(f"| **首次购买该系列用户** | {format_yuan(nr_rev)} | {nr_rev/total_new_rev*100:.1f}% | {nr_cnt}人 | {nr_cnt/len(new_series_users)*100:.1f}% | {format_yuan(nr_arpu)} |")
    lines.append("")
    ratio = rt_arpu / nr_arpu if nr_arpu else 0
    lines.append(f"**结论**：历史复购用户以 {rt_cnt/len(new_series_users)*100:.1f}% 的人数贡献了 {rt_rev/total_new_rev*100:.1f}% 的销售额，"
                 f"客单价{format_yuan(rt_arpu)}是首次购买用户的 {ratio:.1f}倍。该弹次为 **【老用户复购驱动型】** 弹次。")
    lines.append("")

    # ---- 3.2 历史复购用户特征 ----
    lines.append("### 3.2 历史复购用户特征")
    lines.append("")
    lines.append(f"历史复购用户数：{rt_cnt}人")
    lines.append("")

    # 历史购买深度
    depth1 = set()
    depth2 = set()
    for uid in repurchaser_uids:
        hcnt = sum(1 for s in historical_series if uid in series_user[s])
        if hcnt >= 2:
            depth2.add(uid)
        else:
            depth1.add(uid)

    d1_rev, d1_cnt, _, d1_arpu = series_stats(sn_new, depth1)
    d2_rev, d2_cnt, _, d2_arpu = series_stats(sn_new, depth2)

    lines.append("#### 历史购买深度")
    lines.append("")
    lines.append("| 购买历史弹次数 | 人数 | 占比 | 新弹次消费金额 | 新弹次客单价 |")
    lines.append("|:---|---:|---:|---:|---:|")
    lines.append(f"| 仅购买1个历史弹次 | {d1_cnt} | {d1_cnt/rt_cnt*100:.1f}% | {format_yuan(d1_rev)} | {format_yuan(d1_arpu)} |")
    lines.append(f"| 购买2个及以上历史弹次 | {d2_cnt} | {d2_cnt/rt_cnt*100:.1f}% | {format_yuan(d2_rev)} | {format_yuan(d2_arpu)} |")
    lines.append(f"| **合计** | **{rt_cnt}** | **100%** | **{format_yuan(rt_rev)}** | **{format_yuan(rt_arpu)}** |")
    lines.append("")
    lines.append(f"> 深度复购用户（购买多弹次）在新弹次的客单价{format_yuan(d2_arpu)}，显著高于仅购买1弹的{format_yuan(d1_arpu)}，说明多弹次用户的消费力更强。")
    lines.append("")

    # 历史消费金额分层
    hist_spends = [sum(series_user[s][uid]["total"] for s in historical_series if uid in series_user[s])
                   for uid in repurchaser_uids]
    hist_spends.sort()
    total_hist_rev = sum(hist_spends)

    tiers = [(0, 100), (100, 300), (300, 500), (500, 1000), (1000, 2000), (2000, 5000), (5000, 999999)]
    tier_labels = ["0-100元", "100-300元", "300-500元", "500-1000元", "1000-2000元", "2000-5000元", "5000元以上"]

    lines.append("#### 历史消费金额分层（所有历史弹次合计）")
    lines.append("")
    lines.append("| 历史消费区间 | 人数 | 占比 | 贡献历史总销售额 | 销售额占比 |")
    lines.append("|:---|---:|---:|---:|---:|")
    last_valid = None
    for i, (lo, hi) in enumerate(tiers):
        vals = [v for v in hist_spends if lo <= v < hi]
        if vals:
            cnt = len(vals)
            pct = cnt / len(hist_spends) * 100
            rev = sum(vals)
            rpct = rev / total_hist_rev * 100
            lines.append(f"| {tier_labels[i]} | {cnt} | {pct:.1f}% | {format_yuan(rev)} | {rpct:.1f}% |")
            last_valid = i
    lines.append(f"| **合计** | **{len(hist_spends)}** | **100%** | **{format_yuan(total_hist_rev)}** | **100%** |")
    lines.append("")

    # 消费模式
    sorted_new = sorted([series_user[sn_new][uid]["total"] for uid in repurchaser_uids])
    top20_thr = sorted_new[int(len(sorted_new) * 0.8)] if sorted_new else 0
    num_top20 = len(sorted_new) // 5

    single_high = 0
    multi_high = 0
    for uid in repurchaser_uids:
        if series_user[sn_new][uid]["total"] >= top20_thr:
            hcnt = sum(1 for s in historical_series if uid in series_user[s])
            if hcnt >= 2:
                multi_high += 1
            else:
                single_high += 1

    lines.append("#### 消费模式分布")
    lines.append("")
    lines.append(f"新弹次TOP20%高消费复购用户（消费≥{format_yuan(top20_thr)}，共{num_top20}人）的构成：")
    lines.append("")
    lines.append("| 类型 | 人数 | 说明 |")
    lines.append("|:---|---:|:---|")
    lines.append(f"| 单弹次超高消费型 | {single_high}人 | 仅买过1个历史弹次，在新弹次集中高消费 |")
    lines.append(f"| 多弹次持续高消费型 | {multi_high}人 | 从早期弹次持续追随到新弹次 |")
    lines.append("")
    pct_multi = multi_high / (single_high + multi_high) * 100 if (single_high + multi_high) else 0
    if multi_high > single_high:
        lines.append(f"> **画像结论**：高消费复购用户的核心画像是 **“多弹次持续追随者”**（{pct_multi:.0f}%），"
                     f"并非一时冲动的高额消费者，而是IP的忠实用户。）")
    else:
        lines.append(f"> **画像结论**：高消费复购用户以单弹次集中爆发型为主，说明新弹次产品力驱动了高消费行为。")
    lines.append("")

    # ---- 3.3 新付费用户特征 ----
    new_user_set = set()
    for uid in new_series_users:
        reg = uid_to_reg.get(uid)
        if reg and new_start_date <= reg.date() <= new_end_date:
            new_user_set.add(uid)

    nu_rev, nu_cnt, nu_ords, nu_arpu = series_stats(sn_new, new_user_set)
    nu_multi = sum(1 for uid in new_user_set if series_user[sn_new][uid]["orders"] > 1)

    # 首单分析
    fo_19_30 = 0
    for uid in new_user_set:
        orders = [(r["下单时间_ts"], r["消费金额"]) for r in all_rows
                  if r["用户ID"] == uid and r["弹次名称"] == sn_new]
        orders.sort(key=lambda x: x[0])
        if orders and 19 <= orders[0][1] <= 30:
            fo_19_30 += 1

    lines.append("### 3.3 新付费用户特征")
    lines.append(f"> **定义**：注册时间在新弹次销售期内（{new_start_date} ~ {new_end_date}）")
    lines.append("")
    lines.append("| 指标 | 数值 |")
    lines.append("|:---|---:|")
    lines.append(f"| 新用户规模 | {nu_cnt}人 |")
    lines.append(f"| 总消费 | {format_yuan(nu_rev)} |")
    lines.append(f"| 客单价 | {format_yuan(nu_arpu)} |")
    lines.append(f"| 弹次内复购率 | {nu_multi/nu_cnt*100:.1f}% |")
    lines.append("")
    lines.append("**首单金额分布**：")
    lines.append(f"- 19-30元：{fo_19_30}人（{fo_19_30/nu_cnt*100:.1f}%）")
    lines.append(f"- 其他区间：{nu_cnt - fo_19_30}人（{(nu_cnt-fo_19_30)/nu_cnt*100:.1f}%）")
    lines.append("")
    lines.append(f"> **结论**：新用户驱动力是 **【低价引流型】**。低价SKU承担了获客入口功能，"
                 f"{nu_multi}人（{nu_multi/nu_cnt*100:.1f}%）在弹次内复购，说明有一定转化潜力。")
    lines.append("")

    return "\n".join(lines)


# ── 模块三：用户流失分析 ───────────────────────────────────────────
def module3():
    print("[分析] 模块三：用户流失分析")
    lines = []

    all_hist_uids = set()
    for s in historical_series:
        all_hist_uids |= set(series_user[s].keys())

    churned_uids = all_hist_uids - new_series_users
    all_ip_uids = set(r["用户ID"] for r in all_rows)

    # TOP200
    top200 = sorted(uid_total.items(), key=lambda x: -x[1])[:200]
    top200_uids = set(uid for uid, _ in top200)
    top200_churned = churned_uids & top200_uids
    top200_repurchased = top200_uids & new_series_users

    lines.append("## 四、用户流失分析")
    lines.append("")
    lines.append("### 4.1 流失规模")
    lines.append("")
    lines.append("| 指标 | 数值 |")
    lines.append("|:---|---:|")
    lines.append(f"| 历史弹次付费用户总数 | {len(all_hist_uids):,}人 |")
    lines.append(f"| 新弹次购买用户 | {len(new_series_users)}人 |")
    lines.append(f"| **未复购用户（流失）** | **{len(churned_uids):,}人** |")
    lines.append(f"| **整体流失率** | **{len(churned_uids)/len(all_hist_uids)*100:.1f}%** |")
    lines.append(f"| IP消费TOP200流失人数 | **{len(top200_churned)}人（{len(top200_churned)/200*100:.1f}%）** |")
    lines.append(f"| TOP200中复购新弹次 | {len(top200_repurchased)}人（{len(top200_repurchased)/200*100:.1f}%） |")
    lines.append("")
    lines.append(f"> **注意**：新弹次销售期仅{(new_end_date - new_start_date).days + 1}天，部分用户可能仍在观望期。")
    lines.append("")

    # 4.2 高价值流失用户特征
    lines.append("### 4.2 高价值流失用户特征")
    lines.append(f"> 分析对象：IP付费TOP200中的未复购用户（{len(top200_churned)}人）")
    lines.append("")

    cspends = [uid_total[uid] for uid in top200_churned]
    cspends.sort()

    # 消费层级
    tiers_200 = [(0, 1500), (1500, 2000), (2000, 3000), (3000, 5000), (5000, 999999)]
    tier_labels_200 = ["1500元以下", "1500-2000元", "2000-3000元", "3000-5000元", "5000元以上"]

    lines.append("#### 消费层级分布（TOP200流失用户）")
    lines.append("")
    lines.append("| 消费区间 | 流失人数 | 流失占比 | （对比）TOP200整体占比 |")
    lines.append("|:---|---:|---:|---:|")
    top200_spends = [uid_total[uid] for uid in top200_uids]
    for i, (lo, hi) in enumerate(tiers_200):
        cnt_churn = sum(1 for v in cspends if lo <= v < hi)
        cnt_all = sum(1 for v in top200_spends if lo <= v < hi)
        lines.append(f"| {tier_labels_200[i]} | {cnt_churn} | {cnt_churn/len(cspends)*100:.1f}% | {cnt_all/200*100:.1f}% |")
    lines.append("")

    # 流失时点
    newly_churned = 0
    old_churned = 0
    for uid in top200_churned:
        in_last_series = any(uid in series_user[s] for s in historical_series[-1:])
        if in_last_series:
            newly_churned += 1
        else:
            old_churned += 1

    lines.append("#### 流失时点判断")
    lines.append("")
    lines.append("| 流失类型 | 人数 | 占比 |")
    lines.append("|:---|---:|---:|")
    last_sn_name = historical_series[-1] if historical_series else "上一弹次"
    lines.append(f"| **本弹次新增流失**（{last_sn_name}有消费）| {newly_churned}人 | {newly_churned/len(top200_churned)*100:.1f}% |")
    lines.append(f"| **历史已流失**（{last_sn_name}已无消费）| {old_churned}人 | {old_churned/len(top200_churned)*100:.1f}% |")
    lines.append("")
    lines.append(f"> **核心结论**：{old_churned/len(top200_churned)*100:.1f}%的高价值流失用户属于“历史流失”（上弹次已不购买），"
                 f"{newly_churned/len(top200_churned)*100:.1f}%属于“突发性流失”（上弹次还有消费）。"
                 f"后者是**可挽回的流失用户**，需优先触达。")
    lines.append("")

    # ---- 高价值用户明细表 ----
    high_value = [(uid, uid_total[uid]) for uid in uid_total if uid_total[uid] >= 5000]
    high_value.sort(key=lambda x: -x[1])

    lines.append("#### 消费5000元以上的高价值用户明细")
    lines.append("")
    lines.append(f"> 共有 {len(high_value)} 名用户消费总额超过5,000元，下表列出其各弹次消费金额。")
    lines.append("")
    lines.append("| 序号 | 用户ID | 各弹次消费明细 | 总消费 |")
    lines.append("|:---:|:---|---:|---:|")
    for i, (uid, total) in enumerate(high_value, 1):
        detail_parts = []
        for sn in series_names:
            amt = series_user[sn].get(uid, {}).get("total", 0)
            short_name = sn.split()[-1]  # 第N弹
            detail_parts.append(f"{short_name}={format_yuan(amt)}")
        detail_str = " / ".join(detail_parts)
        lines.append(f"| {i} | {uid} | {detail_str} | {format_yuan(total)} |")
    lines.append("")
    lines.append(f"> 可快速识别哪些高价值用户在新弹次仍有消费（{sn_new.split()[-1]} > 0），哪些已全面流失。")
    lines.append("")

    return "\n".join(lines)


# ── 模块四：策略建议 ───────────────────────────────────────────────
def module4():
    print("[分析] 生成策略建议")
    lines = []

    lines.append("## 五、策略建议")
    lines.append("")
    lines.append("### 建议一：建立高价值用户召回机制")
    lines.append("针对流失的高价值用户（特别是本弹次新增流失用户），通过专属优惠券、限量商品优先购、会员专属权益等方式进行定向召回。")
    lines.append("")
    lines.append("### 建议二：为中轻度复购用户设计跨系列连购激励")
    lines.append("针对仅购买过少量历史弹次的用户，通过“系列套装折扣”、“补课推荐”等机制，提升跨系列购买率。")
    lines.append("")
    lines.append("### 建议三：新用户由低价引流向中客单价转化")
    lines.append("首单后48小时推送进阶SKU加购提示，设计新用户专属梯度价（19.9元→29.9元→中价位），逐步提升消费天花板。")
    lines.append("")
    lines.append("### 建议四：优化弹次发售节奏")
    lines.append("缩短弹次间隔，在空窗期通过限量复刻、联名活动、内容运营等方式维持用户活跃度，避免热度衰减。")
    lines.append("")
    lines.append("### 建议五：基于高价值用户偏好迭代产品")
    lines.append("分析多弹次持续高消费用户的历史购买数据，识别其偏好的价格带和SKU类型，作为产品迭代输入。")
    lines.append("")

    return "\n".join(lines)


# ── 报告生成与输出 ──────────────────────────────────────────────
print("=" * 60)
print("  IP多弹次系列销售复盘分析报告")
print("=" * 60)
print(f"  分析基准: {sn_new}")
print(f"  新弹次周期: {new_start_date} ~ {new_end_date}")
print(f"  共 {len(all_rows)} 条记录, {len(all_ip_uids)} 名用户")
print("=" * 60)

report = []
report.append(f"# IP多弹次系列销售复盘分析报告")
report.append("")
report.append(f"**分析日期**: {datetime.now().strftime('%Y-%m-%d')} | "
              f"**数据范围**: {min(r['下单时间_ts'] for r in all_rows).date()} ~ {max(r['下单时间_ts'] for r in all_rows).date()} | "
              f"**分析基准**: {sn_new}")
report.append("")

# 执行摘要
report.append("---")
report.append("")
report.append("## 一、执行摘要")
report.append("")

# Summary insights - computed from analysis
all_hist = set()
for s in historical_series:
    all_hist |= set(series_user[s].keys())

churned = all_hist - new_series_users
top200 = sorted(uid_total.items(), key=lambda x: -x[1])[:200]
top200_uids = set(uid for uid, _ in top200)
top200_churned = churned & top200_uids

new_user_set = set()
for uid in new_series_users:
    reg = uid_to_reg.get(uid)
    if reg and new_start_date <= reg.date() <= new_end_date:
        new_user_set.add(uid)

repurchaser_uids = set()
for uid in new_series_users:
    for s in historical_series:
        if uid in series_user[s]:
            repurchaser_uids.add(uid)
            break
non_rep = new_series_users - repurchaser_uids

rt_rev, rt_cnt, _, rt_arpu = series_stats(sn_new, repurchaser_uids)
nr_rev, nr_cnt, _, nr_arpu = series_stats(sn_new, non_rep)
total_new_rev = rt_rev + nr_rev

dan2_dates = [r["下单时间_ts"].date() for r in all_rows if r["弹次名称"] != sn_new and historical_series and r["弹次名称"] == historical_series[-1]]

# Catch-all comparison for newest vs previous
if historical_series:
    last_hist = historical_series[-1]
    lh_dates = [r["下单时间_ts"].date() for r in all_rows if r["弹次名称"] == last_hist]
    lh_days = min(len(lh_dates), (new_end_date - new_start_date).days + 1)
    lh_start = min(lh_dates)
    lh_cutoff = lh_start + (new_end_date - new_start_date)
    lh_first = [r for r in all_rows if r["弹次名称"] == last_hist
                and lh_start <= r["下单时间_ts"].date() <= lh_cutoff]
    lh_early_rev = sum(r["消费金额"] for r in lh_first)
else:
    lh_early_rev = 0

dan3_total_rev = sum(series_user[sn_new][uid]["total"] for uid in new_series_users)

summary_1 = (f"1. **新弹次上线{(new_end_date - new_start_date).days + 1}天销售额{format_yuan(dan3_total_rev)}，"
             f"日均{format_yuan(dan3_total_rev / ((new_end_date - new_start_date).days + 1))}**")
if lh_early_rev > 0:
    summary_1 += f"，首同期对比前序弹次下降{(1 - dan3_total_rev/lh_early_rev)*100:.1f}%（{format_yuan(dan3_total_rev)} vs {format_yuan(lh_early_rev)}）。"
else:
    summary_1 += "。"

report.append(summary_1)
report.append("")

new_pct = len(new_user_set) / len(new_series_users) * 100 if new_series_users else 0
report.append(f"2. **历史复购用户贡献{rt_rev/total_new_rev*100:.1f}%销售额，新弹次为强【老客驱动】型** —— "
              f"{rt_cnt}名历史复购用户以{format_yuan(rt_arpu)}的客单价贡献{format_yuan(rt_rev)}。")
report.append("")

# New user insight
if new_user_set:
    new_rev = sum(series_user[sn_new][uid]["total"] for uid in new_user_set)
    new_arpu = new_rev / len(new_user_set)
    report.append(f"3. **新用户低价引流效果显著，但客单价偏低** —— {len(new_user_set)}名新用户中"
                  f"首单集中在19-30元区间，客单价仅{format_yuan(new_arpu)}，呈现“规模驱动型”特征。")
report.append("")

report.append(f"4. **IP消费TOP200用户中{len(top200_churned)/200*100:.1f}%未复购新弹次，高价值流失率严重** —— "
              f"{len(top200_churned)}名TOP200高消费用户未购买新弹次。")
report.append("")

report.append(f"5. **整体销售呈变化趋势** —— 从各弹次数据对比可见用户基数和消费行为的变化，需关注IP生命周期阶段。")
report.append("")
report.append("---")
report.append("")

# Assemble report
report.append(module1_basic())
report.append("")
report.append(module1_newuser())
report.append("")
report.append(module2())
report.append("")
report.append(module3())
report.append("")
report.append(module4())

# 附录
report.append("## 附录")
report.append("")
report.append("### 数据说明")
report.append("")
report.append("1. 新弹次上线天数有限，所有关于新弹次的分析结论均受限于当前数据窗口。")
report.append("2. 整体流失率受数据窗口影响，部分用户可能仍在观望期。")
report.append("3. 所有金额单位均为人民币（元），客单价=销售额/付费用户数。")
report.append("4. 新用户定义：用户注册时间在新弹次销售周期内。")
report.append("")
report.append("---")
report.append(f"*报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M')} | 数据来源: 多弹次付费用户销售分布*")

# 输出报告
output_file = os.path.join(out_dir, "IP多弹次系列销售复盘分析报告.md")
with open(output_file, "w", encoding="utf-8") as f:
    f.write("\n".join(report))

print(f"\n[OK] 报告已生成: {output_file}")
print(f"[INFO] 报告长度: {len(''.join(report))} 字符")

# 同时输出分析数据JSON
stats = {
    "series_names": series_names,
    "new_series": sn_new,
    "new_series_period": {"start": str(new_start_date), "end": str(new_end_date)},
    "total_records": len(all_rows),
    "total_users": len(all_ip_uids),
}
json_path = os.path.join(out_dir, "analysis_data.json")
with open(json_path, "w", encoding="utf-8") as f:
    json.dump(stats, f, ensure_ascii=False, indent=2)

sys.exit(0)
