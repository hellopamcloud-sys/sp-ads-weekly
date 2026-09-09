"""从飞书多维表格「领星SP广告数据」读取周报所需数据，统一成内部结构"""
import datetime, json, os, re
from collections import defaultdict
import feishu as fs

APP = os.environ.get('FS_ADS_APP', 'I3UqbrPWmaBBU3s5ABtcz9BFnNc')
T = {'sku': 'tblxQ8UwgCqpPWab', 'campMap': 'tbl7pd4TorXLHlIX', 'camp': 'tblSsvJeZM9ZhuS6', 'pad': 'tblNDX1Xi7ZzvSPw', 'tgt': 'tbl8W411wHBMl8dY',
     'qry': 'tblTA2ilRmMHAjBI', 'cfg': 'tblbxT6Q8dNFehBU', 'tcfg': 'tbl4bpjV62f7NNER', 'neg': 'tbl24DuQ1OtHTBUr', 'pl': 'tblchFLHyh4yYoeD', 'opt': 'tbloF1hcgwkpFhtX', 'arch': 'tblIlvDcvk7E4Y70'}
STORES = ['HLZ-US', 'YYK-US']
NUM = ['展示量', '点击量', '花费', '订单数', '销售额', '销量']


def series_of(sku):
    s = str(sku or '').upper()
    m = re.search(r'EZ70[56]|FL70[1-6]|VL100|TL801|(?<![A-Z])70[1-6](?![0-9])', s)
    if not m: return '旧品类'
    k = m.group(0); k = 'FL' + k if k.startswith('70') else k
    return 'FL702' if k == 'FL703' else k


def n(v):
    try: return float(v or 0)
    except Exception: return 0.0


def date_range_filter(field, a, b):
    return {'conjunction': 'and', 'conditions': [{'field_name': field, 'operator': 'isGreater', 'value': ['ExactDate', str(fs.date_ms(a) - 86400000)]},
                                                 {'field_name': field, 'operator': 'isLess', 'value': ['ExactDate', str(fs.date_ms(b) + 86400000)]}]}


class Data:
    """所有表的行都是 dict，键为飞书字段名；日期字段已转 datetime.date"""
    def __init__(self, start, end):
        self.start, self.end = start, end
        self.sku = []; self.campMap = []; self.camp = []; self.pad = []; self.tgt = []; self.qry = []; self.cfg = []; self.tcfg = []; self.neg = []; self.pl = []; self.opt = []

    # ---------- 读取 ----------
    def load(self, cache_dir=None):
        if cache_dir and os.path.exists(os.path.join(cache_dir, 'camp.json')):
            for k in T:
                p = os.path.join(cache_dir, k + '.json')
                if os.path.exists(p): setattr(self, k, json.load(open(p)))
        else:
            self.sku = fs.read_all(APP, T['sku']); self.campMap = fs.read_all(APP, T['campMap'])
            self.camp = fs.read_all(APP, T['camp'], filter_=date_range_filter('报表日期', self.start, self.end))
            self.pad = fs.read_all(APP, T['pad'], filter_=date_range_filter('报表日期', self.start, self.end))
            self.tgt = fs.read_all(APP, T['tgt'], filter_=date_range_filter('周(周一)', self.start, self.end))
            self.qry = fs.read_all(APP, T['qry'], filter_=date_range_filter('周(周一)', self.start, self.end))
            self.pl = fs.read_all(APP, T['pl'], filter_=date_range_filter('周(周一)', self.start, self.end))
            self.cfg = fs.read_all(APP, T['cfg']); self.tcfg = fs.read_all(APP, T['tcfg']); self.neg = fs.read_all(APP, T['neg']); self.opt = fs.read_all(APP, T['opt'])
            if cache_dir:
                os.makedirs(cache_dir, exist_ok=True)
                for k in T: json.dump(getattr(self, k), open(os.path.join(cache_dir, k + '.json'), 'w'), ensure_ascii=False, default=str)
        self._normalize()
        return self

    def _normalize(self):
        def fixdate(rows, fields):
            for r in rows:
                for f in fields:
                    v = r.get(f)
                    if isinstance(v, (int, float)): r[f] = fs.ms_date(v)
                    elif isinstance(v, str) and re.match(r'\d{4}-\d{2}-\d{2}$', v): r[f] = datetime.date.fromisoformat(v)
        fixdate(self.camp, ['报表日期']); fixdate(self.pad, ['报表日期']); fixdate(self.tgt, ['周(周一)']); fixdate(self.qry, ['周(周一)']); fixdate(self.pl, ['周(周一)'])
        fixdate(self.cfg, ['更新日期', '开始日期']); fixdate(self.opt, ['提出周', '执行日期', '更新日期'])
        for rows in (self.camp, self.pad, self.tgt, self.qry, self.pl):
            for r in rows:
                for f in NUM + ['订单1d', '销售额1d', '有数据天数']:
                    if f in r or f in NUM: r[f] = n(r.get(f))
                r['广告活动ID'] = str(r.get('广告活动ID') or '')
        # 只保留区间内（过滤器已做，双保险）
        self.camp = [r for r in self.camp if r['报表日期'] and self.start <= r['报表日期'] <= self.end]
        self.pad = [r for r in self.pad if r['报表日期'] and self.start <= r['报表日期'] <= self.end]
        for k in ('tgt', 'qry', 'pl'):
            setattr(self, k, [r for r in getattr(self, k) if r['周(周一)'] and self.start <= r['周(周一)'] <= self.end])
        # 映射
        self.sku_info = {}
        for r in self.sku:
            k = str(r.get('SKU') or '')
            if k: self.sku_info[k] = {'store': r.get('店铺', ''), 'asin': r.get('ASIN', ''), 'series': r.get('系列') or series_of(k), 'name': r.get('品名', ''), 'price': n(r.get('售价')), 'target': n(r.get('目标ACOS%'))}
        self.camp_info = {}
        for r in self.campMap:
            k = str(r.get('广告活动ID') or '')
            if k: self.camp_info[k] = {'store': r.get('店铺', ''), 'alias': r.get('活动别名', ''), 'type': r.get('投放类型', ''), 'mainSku': r.get('主SKU', ''), 'series': r.get('系列', ''), 'skus': r.get('投放SKU列表', '')}
        self.cfg_by_id = {str(r.get('广告活动ID')): r for r in self.cfg}
        for r in self.cfg:
            for f in ('顶部溢价%', '详情页溢价%', '其他位置溢价%', '日预算'): r[f] = n(r.get(f))
        for r in self.tcfg: r['当前出价'] = n(r.get('当前出价'))
        self.bid_by_id = {str(r.get('投放ID')): r for r in self.tcfg}
        # 否定词索引 (store, cid, text.lower) -> set(types)
        self.neg_set = defaultdict(set)
        for r in self.neg:
            self.neg_set[(r.get('店铺'), str(r.get('广告活动ID')), str(r.get('否定内容', '')).lower().strip())].add(r.get('类型', ''))
        self.neg_phrase = defaultdict(set)
        for (st, cid, txt), types in self.neg_set.items():
            if '否定词组' in types: self.neg_phrase[(st, cid)].add(txt)
        # 活动主SKU（按花费）
        cs = defaultdict(lambda: defaultdict(float))
        for r in self.pad: cs[r['广告活动ID']][r.get('SKU', '')] += r['花费']
        self.camp_main_sku = {cid: max(m, key=m.get) for cid, m in cs.items() if m}
        # 系列列表：按成熟期花费排序
        cost = defaultdict(float)
        for r in self.pad: cost[self.sku_series(r.get('SKU'))] += r['花费']
        self.series_list = sorted([s for s in cost if s != '旧品类'], key=lambda s: -cost[s]) + ['旧品类']
        for r in self.sku_info.values():
            if r['series'] not in self.series_list: self.series_list.insert(-1, r['series'])

    # ---------- 辅助 ----------
    def sku_series(self, sku):
        i = self.sku_info.get(str(sku or ''))
        return (i and i['series']) or series_of(sku)

    def sku_name(self, sku):
        i = self.sku_info.get(str(sku or '')); return (i and i['name']) or ''

    def camp_name(self, cid):
        i = self.camp_info.get(str(cid)); c = self.cfg_by_id.get(str(cid))
        return (i and i['alias']) or (c and c.get('活动名称')) or ''

    def camp_series(self, cid):
        i = self.camp_info.get(str(cid))
        if i and i['series']: return i['series']
        if i and i['mainSku']: return self.sku_series(i['mainSku'])
        c = self.cfg_by_id.get(str(cid))
        if c and c.get('系列'): return c['系列']
        m = self.camp_main_sku.get(str(cid)); return self.sku_series(m) if m else ''

    def camp_type(self, cid):
        ts = sorted({r.get('投放类型', '') for r in self.tgt if r['广告活动ID'] == str(cid) and r.get('投放类型')})
        if ts: return '/'.join(ts)
        c = self.cfg_by_id.get(str(cid)); return (c and c.get('投放类型')) or ''

    def camp_store(self, cid):
        i = self.camp_info.get(str(cid)); c = self.cfg_by_id.get(str(cid))
        return (i and i['store']) or (c and c.get('店铺')) or ''

    def neg_status(self, store, cid, query):
        q = str(query).lower().strip(); m = self.neg_set.get((store, str(cid), q))
        if m: return '已否定-' + '/'.join(sorted(x.replace('否定', '') for x in m))
        for p in self.neg_phrase.get((store, str(cid)), ()):
            if p and (' ' + p + ' ') in (' ' + q + ' '): return f'已被词组否定({p})'
        return ''

    def series_target(self, series, default):
        """SKU映射里该系列填得最多的目标ACOS%（>0），否则默认"""
        vals = [i['target'] for i in self.sku_info.values() if i['series'] == series and i['target'] > 0]
        if not vals: return None
        v = max(set(vals), key=vals.count); return v / 100 if v > 1 else v
