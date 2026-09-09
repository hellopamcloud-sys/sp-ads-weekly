"""规则化优化建议：从成熟期数据生成，写入飞书「优化记录」（按记录键去重，已存在的只刷新最近数据）"""
import datetime, re
from collections import defaultdict
import feishu as fs
from data import APP, T

MONEY = lambda v: f'${v:,.0f}'
PCT = lambda v: f'{v * 100:.0f}%'


def agg(rows, keyf, mature):
    m = {}
    for r in rows:
        if not mature(r): continue
        k = keyf(r)
        g = m.setdefault(k, {'imp': 0.0, 'clk': 0.0, 'cost': 0.0, 'ord': 0.0, 'sales': 0.0, 'row': r})
        g['imp'] += r['展示量']; g['clk'] += r['点击量']; g['cost'] += r['花费']; g['ord'] += r['订单数']; g['sales'] += r['销售额']
    for g in m.values():
        g['acos'] = g['cost'] / g['sales'] if g['sales'] else None; g['cpc'] = g['cost'] / g['clk'] if g['clk'] else 0
    return g_sorted(m)


def g_sorted(m): return dict(sorted(m.items(), key=lambda kv: -kv[1]['cost']))


def ev(g):
    a = f"ACOS {PCT(g['acos'])}" if g['acos'] is not None else '零单'
    return f"{MONEY(g['cost'])} / {g['clk']:.0f} 点击 / {g['ord']:.0f} 单 / {a}"


def generate(d, P):
    """返回建议列表（dict，键=飞书字段名）"""
    START, MEND = P['start'], P['mature_end']; MIN = P['min_clicks']; ZC = P['zero_cost']; DEF = P['default_target']
    weeks = P['mature_weeks']
    tgt_of = lambda series: d.series_target(series, None) or DEF
    mat_day = lambda r: START <= r['报表日期'] <= MEND
    mat_wk = lambda r: START <= r['周(周一)'] and r['周(周一)'] + datetime.timedelta(days=6) <= MEND
    out = []
    def add(key, pri, store, series, camp, level, obj, action, value, evidence):
        out.append({'记录键': key, '优先级': pri, '店铺': store, '系列': series, '活动名称': camp, '层级': level, '对象': obj, '动作': action, '建议值': value, '依据': evidence, '最近数据': evidence})

    # ---- 活动级：预算/溢价（活动日报 + 活动配置） ----
    camp_g = agg(d.camp, lambda r: (r['店铺'], r['广告活动ID']), mat_day)
    daily = defaultdict(list)
    for r in d.camp:
        if mat_day(r): daily[(r['店铺'], r['广告活动ID'])].append(r['花费'])
    for (st, cid), g in camp_g.items():
        cfg = d.cfg_by_id.get(cid)
        if not cfg or cfg.get('状态') != 'enabled' or g['cost'] < 3 * 35: continue
        series = d.camp_series(cid); tgt = tgt_of(series); name = d.camp_name(cid) or cid
        budget = cfg['日预算']; avg = g['cost'] / (7 * weeks); full = sum(1 for x in daily[(st, cid)] if budget and x >= budget * 0.9)
        use = avg / budget if budget else 0
        base = f"成熟期 {ev(g)}；日预算 ${budget:g}，日均花费 ${avg:.0f}（{PCT(use)}），跑满 {full} 天/{7 * weeks}"
        if budget and use > 3:
            add(f'{st}|{cid}|预算|异常', 'P2预算', st, series, name, '活动', f'日预算 ${budget:g}', '预算最近被改过（历史日均花费远高于当前预算）：确认是有意关停还是要恢复', f'恢复到 ${avg:.0f} 左右' if budget < 5 else '', base)
        elif g['acos'] is not None and g['acos'] <= tgt and full >= 0.4 * 7 * weeks:
            add(f'{st}|{cid}|预算|加', 'P2预算', st, series, name, '活动', f'日预算 ${budget:g}', '加预算（好活动被预算卡住）', f'${budget * 1.3:.0f}', base)
        elif g['acos'] is not None and g['acos'] > tgt * 1.3 and full >= 0.4 * 7 * weeks:
            add(f'{st}|{cid}|预算|超标跑满', 'P1止血', st, series, name, '活动', f'日预算 ${budget:g}', '顶着预算亏钱：先降投放组出价/溢价，不要加预算', '出价 −20%~−30%', base)
        if cfg['顶部溢价%'] >= 50 and g['acos'] is not None and g['acos'] > tgt:
            add(f'{st}|{cid}|溢价|顶部', 'P1止血', st, series, name, '活动', f"首页顶部溢价 +{cfg['顶部溢价%']:.0f}%", '降顶部溢价', '0%~+20%', base)

    # ---- 投放组级（投放周报） ----
    tg = agg(d.tgt, lambda r: (r['店铺'], r['广告活动ID'], r.get('投放类型', ''), str(r.get('投放ID', ''))), mat_wk)
    AUTO_NAME = {'queryHighRelMatches': '自动-紧密匹配', 'queryBroadRelMatches': '自动-宽泛匹配', 'asinSubstituteRelated': '自动-同类商品', 'asinAccessoryRelated': '自动-关联商品'}
    for (st, cid, ttype, tid), g in tg.items():
        if g['clk'] < MIN: continue
        r = g['row']; series = d.camp_series(cid); tgt = tgt_of(series); name = d.camp_name(cid) or cid
        c_ = d.cfg_by_id.get(cid)
        if c_ and c_.get('状态') != 'enabled': continue
        txt = AUTO_NAME.get(r.get('投放文本', ''), r.get('投放文本', '')); match = r.get('匹配类型', '')
        obj = f"{ttype} {txt}" + (f" [{match}]" if match and ttype == '关键词' else '')
        bid = d.bid_by_id.get(tid); bidv = bid['当前出价'] if bid else None
        if bid and bid.get('状态') == 'paused': continue  # 已被手动暂停
        bidtxt = lambda f: (f'${bidv * f:.2f}（现 ${bidv:.2f}）' if bidv else f'出价 ×{f}')
        if g['ord'] == 0 and g['cost'] >= ZC * 2:
            add(f'{st}|{cid}|组|{tid}|零单', 'P1止血', st, series, name, '投放组', obj, '暂停（大样本零单）', bidtxt(0.5) + ' 或暂停', ev(g))
        elif g['acos'] is not None and g['acos'] > tgt * 2:
            add(f'{st}|{cid}|组|{tid}|超标', 'P1止血', st, series, name, '投放组', obj, '出价 −30%', bidtxt(0.7), ev(g) + f'（目标 {PCT(tgt)}）')
        elif g['acos'] is not None and g['acos'] > tgt * 1.3:
            add(f'{st}|{cid}|组|{tid}|偏高', 'P1止血', st, series, name, '投放组', obj, '出价 −15%', bidtxt(0.85), ev(g) + f'（目标 {PCT(tgt)}）')
        elif g['acos'] is not None and g['acos'] <= tgt * 0.8 and g['ord'] >= 5:
            add(f'{st}|{cid}|组|{tid}|加码', 'P2收割加码', st, series, name, '投放组', obj, '出价 +10%（盈利组多拿量）', bidtxt(1.1), ev(g) + f'（目标 {PCT(tgt)}）')

    # ---- 搜索词级（搜索词周报） ----
    qg = agg(d.qry, lambda r: (r['店铺'], r['广告活动ID'], r.get('用户搜索词', '')), mat_wk)
    by_term = defaultdict(list)
    for (st, cid, q), g in qg.items():
        by_term[(st, q)].append((cid, g))
        series = d.camp_series(cid); tgt = tgt_of(series); name = d.camp_name(cid) or cid
        neg = d.neg_status(st, cid, q)
        c_ = d.cfg_by_id.get(cid)
        if neg or (c_ and c_.get('状态') != 'enabled'): continue
        is_asin = bool(re.match(r'^b0[a-z0-9]{8}$', q.strip().lower()))
        is_kw = g['row'].get('投放类型', '') == '关键词'
        same_kw = is_kw and q.strip().lower() == str(g['row'].get('投放词/投放组', '')).strip().lower()
        if same_kw: continue  # 精准词本身的表现已在投放组规则里处理
        if g['ord'] == 0 and g['cost'] >= ZC and g['clk'] >= 20:
            add(f'{st}|{cid}|词|{q}|否', 'P1否词', st, series, name, '搜索词', q, '否定商品' if is_asin else '否定精准（活动级）', '', ev(g) + (f"（自家/竞品 ASIN）" if is_asin else ''))
        elif g['acos'] is not None and g['acos'] > tgt * 2 and g['clk'] >= MIN:
            add(f'{st}|{cid}|词|{q}|超标', 'P1否词', st, series, name, '搜索词', q, '否定精准（活动级）' if not is_asin else '否定商品', '', ev(g) + f'（目标 {PCT(tgt)}）')
        elif g['acos'] is not None and g['acos'] <= tgt and g['ord'] >= 3 and g['cost'] >= 30 and not is_kw and not is_asin:
            add(f'{st}|{cid}|词|{q}|收割', 'P2收割加码', st, series, name, '搜索词', q, '收割：加入精准活动（EXACT），原自动活动里否定精准', f"起始出价 ≈ ${g['cpc'] * 1.1:.2f}", ev(g) + f'（自动/商品定位跑出的好词）')
    # ---- 结构：同一词跨多活动 ----
    for (st, q), lst in by_term.items():
        lst = [x for x in lst if x[1]['cost'] >= 20]
        if len(lst) >= 3 and sum(g['cost'] for _, g in lst) >= 100:
            parts = '；'.join(f"{d.camp_name(c) or c}：{ev(g)}" for c, g in sorted(lst, key=lambda x: -x[1]['cost']))
            keep = [c for c, g in lst if g['acos'] is not None and g['acos'] <= tgt_of(d.camp_series(c))]
            act = (f"只留转化好的 {len(keep)} 个活动跑这个词（{'、'.join(d.camp_name(c) or c for c in keep)}），其余活动否定精准" if keep else '没有一个活动达标：全部降出价，只保留 CPC 最低的一个观察')
            add(f'{st}|结构|{q}', 'P3结构', st, '多系列' if len({d.camp_series(c) for c, _ in lst}) > 1 else d.camp_series(lst[0][0]), f'{len(lst)} 个活动', '搜索词', q, act, '', f"同一搜索词在 {len(lst)} 个活动同时消耗（内部竞价）：{parts}")
    return out


def sync(d, P, suggestions, dry=False):
    """把建议 upsert 到飞书优化记录；返回 (新增数, 刷新数, 全部记录列表)"""
    week = fs.date_ms(P['week_of']); today = fs.date_ms(P['snap'])
    exist = {o.get('记录键'): o for o in d.opt if o.get('记录键')}
    creates, updates = [], []
    for s in suggestions:
        o = exist.get(s['记录键'])
        if o:
            if o.get('最近数据') != s['最近数据']: updates.append({'record_id': o['_rid'], 'fields': {'最近数据': s['最近数据'], '更新日期': today}})
            o['最近数据'] = s['最近数据']
        else:
            f = dict(s); f.update({'提出周': week, '状态': '待处理', '来源': '自动', '更新日期': today}); creates.append({'fields': f})
    if not dry:
        if creates: fs.batch_create(APP, T['opt'], creates)
        if updates: fs.batch_update(APP, T['opt'], updates)
    rows = list(d.opt)
    for c in creates:
        f = dict(c['fields']); f['提出周'] = P['week_of']; f['更新日期'] = P['snap']; rows.append(f)
    return len(creates), len(updates), rows
