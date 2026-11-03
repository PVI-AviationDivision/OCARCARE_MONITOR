"""Soạn tin nhắn Zalo tổng kết hằng ngày từ dữ liệu OCARCARE Monitor.
- tin_nhan_noibo.txt : nhóm nội bộ các đơn vị BH PVI (đủ số liệu, phần BH PVI 40%, đơn vị, đối soát)
- tin_nhan_chung.txt : nhóm có cán bộ VPBank / OPES (chỉ số liệu chung, không có dữ liệu nhạy cảm)
"""
import datetime as dt
from collections import defaultdict

LINK = "https://pvi-aviationdivision.github.io/OCARCARE_MONITOR/"
THANG = lambda d: f"tháng {d.month}"


def tr(v):
    v = v or 0
    s = f"{v/1e9:,.2f} tỷ" if abs(v) >= 1e9 else f"{v/1e6:,.1f} tr"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def vi(n):
    return f"{int(n):,}".replace(",", ".")


def hub(n):
    u = (n or "").upper()
    return "Hub xe Miền Bắc" if "BAC" in u else ("Hub xe Miền Nam" if "NAM" in u else n)


def make(D):
    M, days, P = D["meta"], D["days"], D["points"]
    share = M.get("share", 0.4)
    rd = dt.date.fromisoformat(M["report_date"])
    y = rd - dt.timedelta(days=1)                          # ngày đủ số liệu gần nhất
    dmap = {dt.date.fromisoformat(d): (n, p) for d, n, p in days}
    yn, yp = dmap.get(y, (0, 0))
    mk = y.strftime("%Y-%m")
    mdays = [(d, n, p) for d, (n, p) in dmap.items() if d.strftime("%Y-%m") == mk and d <= y]
    mn, mp = sum(x[1] for x in mdays), sum(x[2] for x in mdays)
    nd = max(1, len(mdays))
    # cùng kỳ tháng trước
    pm = (y.replace(day=1) - dt.timedelta(days=1))
    pmk = pm.strftime("%Y-%m")
    pmp = sum(p for d, (n, p) in dmap.items() if d.strftime("%Y-%m") == pmk and d.day <= y.day)
    pm_full = sum(p for d, (n, p) in dmap.items() if d.strftime("%Y-%m") == pmk)
    pm_fn = sum(n for d, (n, p) in dmap.items() if d.strftime("%Y-%m") == pmk)
    # 7 ngày
    w7 = [dmap.get(y - dt.timedelta(days=i), (0, 0)) for i in range(7)]
    w7n, w7p = sum(x[0] for x in w7), sum(x[1] for x in w7)

    # theo đơn vị (tháng hiện tại) & điểm bán
    U = defaultdict(lambda: [0, 0, 0, 0])        # n tháng, p tháng, số điểm, số điểm có đơn tháng
    for z in P:
        if z["k"] == "HO":
            continue
        u = U[z["v"] or "Chưa xác định"]
        mm = (z.get("m") or {}).get(mk, [0, 0])
        u[0] += mm[0]; u[1] += mm[1]; u[2] += 1; u[3] += 1 if mm[0] else 0
    topU = sorted([(k, v) for k, v in U.items() if v[1] > 0], key=lambda kv: -kv[1][1])
    zeroU = sorted(k for k, v in U.items() if v[0] == 0 and k != "Chưa xác định")
    name = lambda z: hub(z["n"]) if z["k"] == "HUB" else z["n"]
    yp_pts = sorted([z for z in P if z["y"][0] > 0], key=lambda z: (-z["y"][0], -z["y"][1]))
    topP = sorted([z for z in P if (z.get("m") or {}).get(mk, [0, 0])[1] > 0],
                  key=lambda z: -z["m"][mk][1])[:5]
    first_new = [z for z in P if z.get("first") and z["first"] >= (y - dt.timedelta(days=2)).isoformat()]
    stale = [z for z in P if z["pn"] > 0 and z["last"] and (rd - dt.date.fromisoformat(z["last"])).days > M.get("stale_days", 7)]
    staleU = defaultdict(int)
    for z in stale:
        staleU[z["v"]] += 1
    cmp_txt = ""
    if pmp and y.day >= 5:
        ch = (mp - pmp) / pmp * 100
        cmp_txt = f" | {'▲' if ch >= 0 else '▼'}{abs(ch):.0f}% so cùng kỳ {THANG(pm)}".replace(".", ",")

    # ---------- nội bộ
    L = [f"📊 OCARCARE – BÁO CÁO NGÀY {rd:%d/%m/%Y}",
         f"(Số liệu báo cáo BH PVI {M.get('report_time') or ''} {rd:%d/%m}, phí chưa VAT)".replace("  ", " "), "",
         f"▶ Ngày {y:%d/%m}: {yn} HĐBH | {tr(yp)} | phần BH PVI {int(share*100)}%: {tr(yp*share)}",
         f"▶ Luỹ kế {THANG(y)} (01–{y:%d/%m}): {vi(mn)} HĐBH | {tr(mp)} | BQ {tr(mp/nd)}/ngày{cmp_txt}",
         f"▶ 7 ngày gần nhất: {w7n} HĐBH | {tr(w7p)}",
         f"▶ Luỹ kế từ 21/09: {vi(M['total_n'])} HĐBH | {tr(M['total_p'])} | phần BH PVI: {tr(M['total_p']*share)}"]
    if y.day <= 3 and pm_full:
        L.append(f"   ({THANG(pm).capitalize()} chốt: {vi(pm_fn)} HĐBH | {tr(pm_full)})")
    if topU:
        L += ["", f"🏢 Đơn vị BH PVI dẫn đầu {THANG(y)}:"]
        L += [f"{i}. {k.replace('BH PVI ', '')} – {tr(v[1])} ({v[0]} HĐ)" for i, (k, v) in enumerate(topU[:5], 1)]
    if yp_pts:
        L += ["", f"📍 Điểm bán ra đơn ngày {y:%d/%m}: " + "; ".join(f"{name(z)} {z['y'][0]} HĐ" for z in yp_pts[:6])
              + (f" (+{len(yp_pts)-6} điểm khác)" if len(yp_pts) > 6 else "")]
    if first_new:
        L.append("🆕 Điểm bán có đơn đầu tiên: " + ", ".join(f"{name(z)} ({z['v'].replace('BH PVI ', '')})" for z in first_new[:5]))
    warn = []
    if zeroU and y.day >= 5:
        warn.append(f"• Chưa có đơn trong {THANG(y)}: " + ", ".join(u.replace("BH PVI ", "") for u in zeroU))
    if stale:
        top_s = sorted(staleU.items(), key=lambda kv: -kv[1])[:5]
        warn.append(f"• {len(stale)} điểm bán quá {M.get('stale_days', 7)} ngày chưa có đơn mới, nhiều nhất: "
                    + ", ".join(f"{k.replace('BH PVI ', '')} {v}" for k, v in top_s))
    R = D.get("recon") or {}
    if R.get("only_opes"):
        n_o = len(R["only_opes"]) if isinstance(R["only_opes"], list) else R["only_opes"]
        warn.append(f"• Đối soát: {n_o} HĐ có bên OPES chưa có trên báo cáo BH PVI (xem tab Đối soát)")
    if warn:
        L += ["", "⚠ Cần lưu ý:"] + warn
    L += ["", f"👉 Chi tiết: {LINK}", "(Dashboard nội bộ BH PVI – không chia sẻ ra ngoài)"]
    noibo = "\n".join(L)

    # ---------- nhóm chung (VPBank / OPES)
    G = [f"🚗 OCARCARE – CẬP NHẬT NGÀY {rd:%d/%m/%Y}", "",
         f"• Ngày {y:%d/%m}: {yn} hợp đồng | {tr(yp)}",
         f"• Luỹ kế {THANG(y)}: {vi(mn)} hợp đồng | {tr(mp)}"]
    if y.day <= 3 and pm_full:
        G.append(f"• {THANG(pm).capitalize()}: {vi(pm_fn)} hợp đồng | {tr(pm_full)}")
    if topP:
        G += ["", f"🏆 CN/PGD dẫn đầu {THANG(y)}:"]
        G += [f"{i}. {name(z)} – {tr(z['m'][mk][1])} ({z['m'][mk][0]} HĐ)" for i, z in enumerate(topP, 1)]
    if yp_pts:
        G += ["", f"👏 Cảm ơn các CN/PGD đã ra đơn ngày {y:%d/%m}: " + ", ".join(name(z) for z in yp_pts[:8])
              + (f" và {len(yp_pts)-8} điểm khác" if len(yp_pts) > 8 else "")]
    G += ["", "Bảo hiểm PVI đồng hành cùng VPBank – chúc các anh chị một ngày bán hàng hiệu quả! 💪"]
    chung = "\n".join(G)
    return noibo, chung
