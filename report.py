"""SP 广告周报 Excel 生成（数据来自 data.Data）。看板层全部为 SUMIFS/INDEX-MATCH 公式。"""
import datetime, json, re
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import FormulaRule

HFONT = Font(name='Arial', bold=True, color='FFFFFF'); HFILL = PatternFill('solid', fgColor='1F4E78'); SUBFILL = PatternFill('solid', fgColor='D9E1F2')
BF = Font(name='Arial'); GF = Font(name='Arial', color='008000'); BLUE = Font(name='Arial', color='0000FF'); BOLD = Font(name='Arial', bold=True)
GRAY = Font(name='Arial', color='808080'); NOTE = Font(name='Arial', italic=True, color='808080')
thin = Side(style='thin', color='BFBFBF'); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
YELLOW = PatternFill('solid', fgColor='FFFF00'); RED = PatternFill('solid', fgColor='F8CBAD'); GREENF = PatternFill('solid', fgColor='C6EFCE'); AMBER = PatternFill('solid', fgColor='FFE699')
FMT = {'money': '$#,##0.00;($#,##0.00);-', 'int': '#,##0;(#,##0);-', 'pct': '0.0%;(0.0%);-', 'x': '0.00;(0.00);-', 'date': 'yyyy-mm-dd', 'txt': '@'}
NUML = ['展示量', '点击量', '花费', '订单数', '销售额', '销量']; NF = ['int', 'int', 'money', 'int', 'money', 'int']
MLAB = ['CTR', 'CPC', 'CVR', 'ACOS', 'ROAS']; MF = ['pct', 'money', 'pct', 'pct', 'x']
MATURE = '"成熟"'; FRESH = '"新鲜"'
PL_ORDER = ['首页顶部', '商品详情页', '搜索结果其他位置', '站外']
STATUS_ORDER = {'待处理': 0, '观察中': 1, '暂缓': 2, '已执行': 3, '放弃': 4}


def header(ws, cols, row=1, fill=HFILL):
    for i, c in enumerate(cols, 1):
        x = ws.cell(row=row, column=i, value=c); x.font = HFONT; x.fill = fill; x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); x.border = BORDER
    ws.freeze_panes = ws.cell(row=row + 1, column=1); ws.row_dimensions[row].height = 30


def widths(ws, w):
    for i, x in enumerate(w, 1): ws.column_dimensions[L(i)].width = x


def put(ws, r, vals, fmts, fonts=None):
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=i, value=v); c.border = BORDER
        c.font = (fonts or {}).get(i, GF if (isinstance(v, str) and v.startswith('=')) else BF)
        if i in fmts: c.number_format = FMT[fmts[i]]


def DATE(d): return f'DATE({d.year},{d.month},{d.day})'


def build(d, P, out_path):
    """d: data.Data；P: dict(start,end,mature_end,snap,weeks[(mon,a,b)],default_target,zero_cost,min_clicks,opt_rows)"""
    START, END, MEND, SNAP, WEEKS = P['start'], P['end'], P['mature_end'], P['snap'], P['weeks']
    FRESH_A = MEND + datetime.timedelta(days=1)
    series_list = d.series_list
    wb = Workbook()

    # ================= 参数 =================
    wsP = wb.active; wsP.title = '参数'
    wsP['A1'] = '参数'; wsP['A1'].font = BOLD
    rows = [('数据起始日', START, 'date'), ('数据截止日', END, 'date'), ('成熟截止日（≤此日为成熟数据；最后一个完整成熟周的周日）', MEND, 'date'), ('生成日期', SNAP, 'date'), ('默认目标ACOS', P['default_target'], 'pct'), ('零单高花费阈值($)', P['zero_cost'], 'money'), ('最低判断点击数（低于此不下结论）', P['min_clicks'], 'int')]
    for i, (k, v, f) in enumerate(rows, 2):
        wsP.cell(row=i, column=1, value=k).font = BF; c = wsP.cell(row=i, column=2, value=v); c.font = BLUE; c.fill = YELLOW; c.number_format = FMT[f]; c.border = BORDER
    wsP.cell(row=9, column=1, value='成熟周数（自动）').font = BF; c = wsP.cell(row=9, column=2, value='=(B4-B2+1)/7'); c.font = BF; c.number_format = '0.0'; c.border = BORDER
    wsP['A11'] = '系列目标ACOS（留空则用默认值；来源=飞书 SKU映射.目标ACOS%）'; wsP['A11'].font = BOLD
    header(wsP, ['系列', '目标ACOS', '说明'], row=12)
    for i, s in enumerate(series_list, 13):
        wsP.cell(row=i, column=1, value=s).font = BF; c = wsP.cell(row=i, column=2, value=d.series_target(s, None)); c.font = BLUE; c.fill = YELLOW; c.number_format = FMT['pct']; c.border = BORDER
        wsP.cell(row=i, column=1).border = BORDER
    SER_RNG = f"参数!$A$13:$A${12 + len(series_list)}"; TGT_RNG = f"参数!$B$13:$B${12 + len(series_list)}"
    def tgt_formula(series_cell): return f'=IFERROR(IF(INDEX({TGT_RNG},MATCH({series_cell},{SER_RNG},0))="",参数!$B$6,INDEX({TGT_RNG},MATCH({series_cell},{SER_RNG},0))),参数!$B$6)'
    def tgt_inline(series_cell): return tgt_formula(series_cell)[1:]
    wsP.freeze_panes = None; widths(wsP, [52, 14, 40])
    wsP['D2'] = '黄色=可改。改成熟截止日后，所有"成熟期/新鲜期"列自动重算；改系列目标ACOS后，标红逻辑自动重算。长期改目标ACOS请填在飞书 SKU映射 表的 目标ACOS% 列，下周自动带入。'; wsP['D2'].font = NOTE

    # ================= 映射表 =================
    skus = {}
    for r in d.pad:
        k = (r['店铺'], r['SKU']); skus.setdefault(k, {'asin': r.get('ASIN', ''), 'cost': 0.0}); skus[k]['cost'] += r['花费']
    store_order = {s: i for i, s in enumerate(['HLZ-US', 'YYK-US'])}
    wsS = wb.create_sheet('SKU映射'); header(wsS, ['店铺', 'SKU', 'ASIN', '系列', '品名', '全期花费', '目标ACOS(按系列)'])
    sku_rows = sorted(skus.items(), key=lambda kv: (store_order.get(kv[0][0], 9), -kv[1]['cost']))
    for i, ((st, sku), info) in enumerate(sku_rows, 2):
        put(wsS, i, [st, sku, info['asin'], d.sku_series(sku), d.sku_name(sku), info['cost'], tgt_formula(f'D{i}')], {6: 'money', 7: 'pct'}, {4: BLUE})
        wsS.cell(row=i, column=4).fill = YELLOW
    n_s = len(sku_rows); widths(wsS, [9, 30, 13, 10, 40, 11, 12]); wsS.auto_filter.ref = wsS.dimensions
    wsS.cell(row=n_s + 3, column=1, value='系列列（黄色）可改，改后各汇总表自动重算。长期修改请改飞书 SKU映射 表。').font = NOTE
    SK_SKU = f"SKU映射!$B$2:$B${n_s + 1}"; SK_SER = f"SKU映射!$D$2:$D${n_s + 1}"

    camp_cost = defaultdict(float); camp_store = {}
    for r in d.camp: camp_cost[r['广告活动ID']] += r['花费']; camp_store[r['广告活动ID']] = r['店铺']
    camp_sku = defaultdict(lambda: defaultdict(float))
    for r in d.pad: camp_sku[r['广告活动ID']][r['SKU']] += r['花费']
    def main_sku(cid):
        i = d.camp_info.get(cid)
        if i and i['mainSku']: return i['mainSku']
        m = camp_sku.get(cid, {}); return max(m, key=m.get) if m else ''
    def ser_idx(s): return series_list.index(s) if s in series_list else 99
    cids = sorted(camp_store, key=lambda c: (store_order.get(camp_store[c], 9), ser_idx(d.camp_series(c)), -camp_cost[c]))
    wsC = wb.create_sheet('活动映射'); header(wsC, ['店铺', '广告活动ID', '活动名称(可改)', '投放类型', '主SKU', '系列', '投放SKU列表', '全期花费'])
    for i, cid in enumerate(cids, 2):
        ci = d.camp_info.get(cid, {})
        put(wsC, i, [camp_store[cid], cid, d.camp_name(cid), ci.get('type', '') or d.camp_type(cid), main_sku(cid), d.camp_series(cid) or f'=IFERROR(INDEX({SK_SER},MATCH(E{i},{SK_SKU},0)),"")', '\n'.join(sorted(camp_sku.get(cid, {}))), camp_cost[cid]], {8: 'money'}, {3: BLUE})
        wsC.cell(row=i, column=3).fill = YELLOW; wsC.cell(row=i, column=7).alignment = Alignment(wrap_text=True, vertical='top')
    n_c = len(cids); widths(wsC, [9, 17, 34, 12, 28, 10, 30, 11]); wsC.auto_filter.ref = wsC.dimensions
    CM_ID = f"活动映射!$B$2:$B${n_c + 1}"; CM_ALIAS = f"活动映射!$C$2:$C${n_c + 1}"; CM_TYPE = f"活动映射!$D$2:$D${n_c + 1}"; CM_MSKU = f"活动映射!$E$2:$E${n_c + 1}"; CM_SER = f"活动映射!$F$2:$F${n_c + 1}"
    def cm(rng, idcell): return f'=IFERROR(INDEX({rng},MATCH({idcell},{CM_ID},0))&"","")'

    # ================= 原始表 =================
    wsD = wb.create_sheet('活动日报'); header(wsD, ['店铺', '日期', '广告活动ID', '系列', '主SKU', '投放类型', '成熟', '展示量', '点击量', '花费', '订单数', '销售额', '销量', '订单1d', '销售额1d'])
    camp = sorted(d.camp, key=lambda r: (store_order.get(r['店铺'], 9), r['报表日期'], -r['花费']))
    for i, r in enumerate(camp, 2):
        put(wsD, i, [r['店铺'], r['报表日期'], r['广告活动ID'], cm(CM_SER, f'C{i}'), cm(CM_MSKU, f'C{i}'), cm(CM_TYPE, f'C{i}'), f'=IF(B{i}<=参数!$B$4,"成熟","新鲜")', r['展示量'], r['点击量'], r['花费'], r['订单数'], r['销售额'], r['销量'], r.get('订单1d', 0), r.get('销售额1d', 0)], {2: 'date', 8: 'int', 9: 'int', 10: 'money', 11: 'int', 12: 'money', 13: 'int', 14: 'int', 15: 'money'})
    n_d = len(camp); widths(wsD, [9, 11, 17, 9, 26, 12, 7] + [10] * 8); wsD.auto_filter.ref = wsD.dimensions
    def DR(col): return f"活动日报!${col}$2:${col}${n_d + 1}"

    wsA = wb.create_sheet('广告商品日报'); header(wsA, ['店铺', '日期', '系列', 'SKU', 'ASIN', '品名', '广告活动ID', '广告组ID', '成熟', '展示量', '点击量', '花费', '订单数', '销售额', '销量', '订单1d', '销售额1d'])
    pad = sorted(d.pad, key=lambda r: (store_order.get(r['店铺'], 9), r['报表日期'], -r['花费']))
    for i, r in enumerate(pad, 2):
        put(wsA, i, [r['店铺'], r['报表日期'], f'=IFERROR(INDEX({SK_SER},MATCH(D{i},{SK_SKU},0)),"")', r['SKU'], r.get('ASIN', ''), d.sku_name(r['SKU']), r['广告活动ID'], str(r.get('广告组ID', '')), f'=IF(B{i}<=参数!$B$4,"成熟","新鲜")', r['展示量'], r['点击量'], r['花费'], r['订单数'], r['销售额'], r['销量'], r.get('订单1d', 0), r.get('销售额1d', 0)], {2: 'date', 10: 'int', 11: 'int', 12: 'money', 13: 'int', 14: 'money', 15: 'int', 16: 'int', 17: 'money'})
    n_a = len(pad); widths(wsA, [9, 11, 9, 30, 13, 36, 17, 17, 7] + [10] * 8); wsA.auto_filter.ref = wsA.dimensions
    def AR(col): return f"广告商品日报!${col}$2:${col}${n_a + 1}"
    has_1d = any(r.get('订单1d') for r in pad)

    wsT = wb.create_sheet('投放周报'); header(wsT, ['店铺', '周(周一)', '广告活动ID', '系列', '广告组ID', '投放类型', '投放ID', '投放文本', '匹配类型', '有数据天数', '成熟', '展示量', '点击量', '花费', '订单数', '销售额', '销量'])
    AUTO_NAME = {'queryHighRelMatches': '自动-紧密匹配', 'queryBroadRelMatches': '自动-宽泛匹配', 'asinSubstituteRelated': '自动-同类商品', 'asinAccessoryRelated': '自动-关联商品'}
    trows = sorted(d.tgt, key=lambda r: (store_order.get(r['店铺'], 9), r['周(周一)'], -r['花费']))
    for i, r in enumerate(trows, 2):
        put(wsT, i, [r['店铺'], r['周(周一)'], r['广告活动ID'], cm(CM_SER, f'C{i}'), str(r.get('广告组ID', '')), r.get('投放类型', ''), str(r.get('投放ID', '')), AUTO_NAME.get(r.get('投放文本', ''), r.get('投放文本', '')), r.get('匹配类型', ''), r.get('有数据天数', 0), f'=IF(B{i}+6<=参数!$B$4,"成熟","新鲜")', r['展示量'], r['点击量'], r['花费'], r['订单数'], r['销售额'], r['销量']], {2: 'date', 10: 'int', 12: 'int', 13: 'int', 14: 'money', 15: 'int', 16: 'money', 17: 'int'})
        wsT.cell(row=i, column=26, value=f"{r.get('投放类型', '')}|{r.get('投放ID', '')}").font = GRAY
    wsT.cell(row=1, column=26, value='投放键').font = HFONT; wsT.cell(row=1, column=26).fill = HFILL
    n_t = len(trows); widths(wsT, [9, 11, 17, 9, 17, 10, 17, 40, 14, 8, 7] + [10] * 6); wsT.auto_filter.ref = f'A1:Q{n_t + 1}'
    def TR(col): return f"投放周报!${col}$2:${col}${n_t + 1}"

    # ================= 汇总表通用 =================
    def metrics(r, ci, cc, cco, co, cs):
        return [f'=IF({ci}{r}=0,0,{cc}{r}/{ci}{r})', f'=IF({cc}{r}=0,0,{cco}{r}/{cc}{r})', f'=IF({cc}{r}=0,0,{co}{r}/{cc}{r})', f'=IF({cs}{r}=0,0,{cco}{r}/{cs}{r})', f'=IF({cco}{r}=0,0,{cs}{r}/{cco}{r})']

    def summary_sheet(title, idx, keylabels):
        ws = wb.create_sheet(title, idx)
        hdr = keylabels + ['目标ACOS'] + ['成熟期 ' + x for x in NUML] + ['成熟期 ' + x for x in MLAB] + ['成熟期花费占店铺%', '与目标差(pp)', '判断'] + ['新鲜期 展示量', '新鲜期 点击量', '新鲜期 CTR', '新鲜期 花费', '新鲜期 CPC', '新鲜期 日均花费', '展示量比周均 (量变，无好坏)', '点击量比周均 (量变，无好坏)', 'CTR比周均 (pp) ↑好', 'CPC比周均 ↑坏 ↓好', '花费比周均 (量变，无好坏)', '新鲜期 1d订单', '新鲜期 1d销售额', '新鲜期 1dACOS', '成熟期 1dACOS', '1dACOS变化 (pp) ↓好', '成熟期 1d订单占比', '预估成熟订单', '预估成熟销售额', '预估ACOS ↓好', '新鲜期 订单(未熟·偏低)', '新鲜期 销售额(未熟·偏低)', '新鲜期 ACOS(未熟·偏高)'] + [f'花费 W{w[0].strftime("%m-%d")}' for w in WEEKS] + [f'ACOS W{w[0].strftime("%m-%d")}' for w in WEEKS]
        header(ws, hdr)
        nk = len(keylabels); c0 = nk + 2
        return ws, hdr, nk, c0

    def write_summary(ws, r, keyvals, nk, c0, crit, src, series_cell, store_cell):
        cols = src['cols']; rng = src['rng']; mat = src['mature']; datec = src['date']; storec = src['store']
        vals = list(keyvals); vals.append(tgt_formula(series_cell))
        for k in cols: vals.append(f'=SUMIFS({rng(k)},{mat},{MATURE}{crit})')
        ci, cc, cco, co, cs = [L(c0 + i) for i in range(5)]
        vals += metrics(r, ci, cc, cco, co, cs)
        tot = f'SUMIFS({rng(cols[2])},{mat},{MATURE},{storec},{store_cell})'
        vals.append(f'=IF({tot}=0,0,{cco}{r}/{tot})')
        acos = L(c0 + 9); tgtc = L(nk + 1); clicks = cc
        vals.append(f'=IF({cs}{r}=0,"",{acos}{r}-{tgtc}{r})')
        vals.append(f'=IF({clicks}{r}<参数!$B$8,"样本不足",IF({cs}{r}=0,"零单",IF({acos}{r}<={tgtc}{r},"达标",IF({acos}{r}<={tgtc}{r}*1.3,"偏高","超标"))))')
        F = lambda k: f'SUMIFS({rng(k)},{mat},{FRESH}{crit})'
        M = lambda k: f'SUMIFS({rng(k)},{mat},{MATURE}{crit})'
        f0 = c0 + 14; X = lambda i: L(f0 + i); wk = '参数!$B$9'
        vals.append('=' + F(cols[0])); vals.append('=' + F(cols[1]))
        vals.append(f'=IF({X(0)}{r}=0,0,{X(1)}{r}/{X(0)}{r})'); vals.append('=' + F(cols[2]))
        vals.append(f'=IF({X(1)}{r}=0,0,{X(3)}{r}/{X(1)}{r})'); vals.append(f'={X(3)}{r}/7')
        vals.append(f'=IF({ci}{r}=0,"",{X(0)}{r}/({ci}{r}/{wk})-1)'); vals.append(f'=IF({cc}{r}=0,"",{X(1)}{r}/({cc}{r}/{wk})-1)')
        vals.append(f'=IF({ci}{r}=0,"",{X(2)}{r}-{L(c0 + 6)}{r})'); vals.append(f'=IF({L(c0 + 7)}{r}=0,"",{X(4)}{r}/{L(c0 + 7)}{r}-1)')
        vals.append(f'=IF({cco}{r}=0,"",{X(3)}{r}/({cco}{r}/{wk})-1)')
        if src.get('o1d'):
            vals.append('=' + F(src['o1d'])); vals.append('=' + F(src['s1d']))
            vals.append(f'=IF({X(12)}{r}=0,0,{X(3)}{r}/{X(12)}{r})')
            vals.append(f'=IF({M(src["s1d"])}=0,0,{cco}{r}/{M(src["s1d"])})')
            vals.append(f'=IF(OR({X(13)}{r}=0,{X(14)}{r}=0),"",{X(13)}{r}-{X(14)}{r})')
            vals.append(f'=IF({co}{r}=0,"",{M(src["o1d"])}/{co}{r})')
            vals.append(f'=IF(OR({X(16)}{r}="",{X(16)}{r}=0),"",{X(11)}{r}/{X(16)}{r})')
            vals.append(f'=IF(OR({cs}{r}=0,{M(src["s1d"])}=0),"",{X(12)}{r}/({M(src["s1d"])}/{cs}{r}))')
            vals.append(f'=IF(OR({X(18)}{r}="",{X(18)}{r}=0),"",{X(3)}{r}/{X(18)}{r})')
        else:
            vals += [''] * 9
        vals.append('=' + F(cols[3])); vals.append('=' + F(cols[4])); vals.append(f'=IF({X(21)}{r}=0,0,{X(3)}{r}/{X(21)}{r})')
        for pk, a, b in WEEKS: vals.append(f'=SUMIFS({rng(cols[2])},{datec},">="&{DATE(a)},{datec},"<="&{DATE(b)}{crit})')
        for j, (pk, a, b) in enumerate(WEEKS):
            s = f'SUMIFS({rng(cols[4])},{datec},">="&{DATE(a)},{datec},"<="&{DATE(b)}{crit})'
            vals.append(f'=IF({s}=0,0,{L(c0 + 37 + j)}{r}/{s})')
        fm = {nk + 1: 'pct'}
        for i, f in enumerate(NF): fm[c0 + i] = f
        for i, f in enumerate(MF): fm[c0 + 6 + i] = f
        fm[c0 + 11] = 'pct'; fm[c0 + 12] = 'pct'
        for i, f in enumerate(['int', 'int', 'pct', 'money', 'money', 'money', 'pct', 'pct', 'pct', 'pct', 'pct', 'int', 'money', 'pct', 'pct', 'pct', 'pct', 'x', 'money', 'pct', 'int', 'money', 'pct']): fm[c0 + 14 + i] = f
        for j in range(len(WEEKS)): fm[c0 + 37 + j] = 'money'; fm[c0 + 37 + len(WEEKS) + j] = 'pct'
        put(ws, r, vals, fm)
        return len(vals)

    def finish(ws, nk, c0, last_row, ncols, widths_):
        widths(ws, widths_); ws.auto_filter.ref = f'A1:{L(ncols)}{last_row}'
        acos = L(c0 + 9); tgtc = L(nk + 1); judge = L(c0 + 13); rng = f'{acos}2:{acos}{last_row}'
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND({judge}2="超标")'], fill=RED))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND({judge}2="偏高")'], fill=AMBER))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND({judge}2="达标")'], fill=GREENF))
        ws.conditional_formatting.add(f'{judge}2:{judge}{last_row}', FormulaRule(formula=[f'{judge}2="零单"'], fill=RED))
        for i, cond in [(9, '>0.15'), (15, '>0.05'), (10, '<-0.3'), (10, '>0.3'), (7, '<-0.3')]:
            cc_ = L(c0 + 14 + i); ws.conditional_formatting.add(f'{cc_}2:{cc_}{last_row}', FormulaRule(formula=[f'AND({cc_}2<>"",{cc_}2{cond})'], fill=RED if cond.startswith('>') or i == 7 else AMBER))
        pa = L(c0 + 14 + 19); ws.conditional_formatting.add(f'{pa}2:{pa}{last_row}', FormulaRule(formula=[f'AND({pa}2<>"",{pa}2>{tgtc}2*1.3)'], fill=RED))
        ws.conditional_formatting.add(f'{pa}2:{pa}{last_row}', FormulaRule(formula=[f'AND({pa}2<>"",{pa}2<={tgtc}2)'], fill=GREENF))
        for j in range(len(WEEKS)):
            wc = L(c0 + 37 + len(WEEKS) + j); ws.conditional_formatting.add(f'{wc}2:{wc}{last_row}', FormulaRule(formula=[f'AND({wc}2>{tgtc}2*1.3,{wc}2>0)'], fill=RED))

    SRC_A = {'cols': 'JKLMNO', 'o1d': 'P' if has_1d else None, 's1d': 'Q' if has_1d else None, 'rng': AR, 'mature': AR('I'), 'date': AR('B'), 'store': AR('A')}
    SRC_D = {'cols': 'HIJKLM', 'o1d': 'N', 's1d': 'O', 'rng': DR, 'mature': DR('G'), 'date': DR('B'), 'store': DR('A')}
    W_BASE = [9] * 6 + [8] * 5 + [10, 9, 9] + [9, 9, 8, 9, 8, 9] + [9] * 5 + [9, 10, 9, 9, 9, 9, 9, 10, 9] + [9, 10, 9] + [10] * len(WEEKS) + [9] * len(WEEKS)
    stores = [s for s in ['HLZ-US', 'YYK-US'] if s in set(camp_store.values())]

    # ================= 整体 =================
    ws = wb.create_sheet('整体', 1)
    header(ws, ['店铺', '周(周一)', '起', '止', '天数', '成熟度'] + NUML + MLAB + ['花费占店铺全期%', '目标ACOS(店铺加权)', '订单1d', '销售额1d', '1d订单占比', 'ACOS(1d口径)', '日均花费'])
    r = 2
    for st in stores:
        first = r
        for pk, a, b in WEEKS + [('全期合计', START, END), ('成熟期合计(5周)', START, MEND), ('新鲜期合计', FRESH_A, END)]:
            vals = [st, pk if isinstance(pk, str) else pk, a, b, f'=D{r}-C{r}+1', f'=IF(D{r}<=参数!$B$4,"成熟",IF(C{r}>参数!$B$4,"新鲜","混合"))']
            for k in 'HIJKLM': vals.append(f'=SUMIFS({DR(k)},{DR("A")},$A{r},{DR("B")},">="&$C{r},{DR("B")},"<="&$D{r})')
            vals += metrics(r, 'G', 'H', 'I', 'J', 'K')
            vals.append(f'=IF(I${first + len(WEEKS)}=0,0,I{r}/I${first + len(WEEKS)})')
            vals.append(f'=IFERROR(SUMPRODUCT(({AR("A")}=$A{r})*({AR("B")}>=$C{r})*({AR("B")}<=$D{r})*{AR("L")}*IFERROR(IF(INDEX({TGT_RNG},MATCH({AR("C")},{SER_RNG},0))="",参数!$B$6,INDEX({TGT_RNG},MATCH({AR("C")},{SER_RNG},0))),参数!$B$6))/I{r},0)')
            vals.append(f'=SUMIFS({DR("N")},{DR("A")},$A{r},{DR("B")},">="&$C{r},{DR("B")},"<="&$D{r})')
            vals.append(f'=SUMIFS({DR("O")},{DR("A")},$A{r},{DR("B")},">="&$C{r},{DR("B")},"<="&$D{r})')
            vals.append(f'=IF(J{r}=0,0,T{r}/J{r})'); vals.append(f'=IF(U{r}=0,0,I{r}/U{r})'); vals.append(f'=I{r}/E{r}')
            put(ws, r, vals, {2: 'date', 3: 'date', 4: 'date', 7: 'int', 8: 'int', 9: 'money', 10: 'int', 11: 'money', 12: 'int', 13: 'pct', 14: 'money', 15: 'pct', 16: 'pct', 17: 'x', 18: 'pct', 19: 'pct', 20: 'int', 21: 'money', 22: 'pct', 23: 'pct', 24: 'money'})
            if isinstance(pk, str):
                for ci_ in range(1, 25): ws.cell(row=r, column=ci_).font = Font(name='Arial', bold=True, color=ws.cell(row=r, column=ci_).font.color); ws.cell(row=r, column=ci_).fill = SUBFILL
            r += 1
        r += 1
    ws.conditional_formatting.add(f'P2:P{r}', FormulaRule(formula=['AND(P2>S2*1.3,P2>0,S2>0)'], fill=RED))
    ws.conditional_formatting.add(f'P2:P{r}', FormulaRule(formula=['AND(P2<=S2,P2>0)'], fill=GREENF))
    ws.cell(row=r, column=1, value='读法：成熟度=成熟 的行可下结论；新鲜行要比较请用 ACOS(1d口径) 和 1d订单占比——1天归因在拉取时已完整，与成熟周同口径可比；新鲜行订单/销售额/ACOS 系统性偏低（7天归因未跑完+领星回传延迟），只看花费/点击/CPC。目标ACOS(店铺加权)=各系列目标按该期花费加权。').font = NOTE
    widths(ws, [9, 12, 11, 11, 6, 7] + [10] * 6 + [8, 8, 8, 8, 8, 12, 12, 9, 10, 9, 9, 10])

    # ================= 系列 =================
    ws, hdr, nk, c0 = summary_sheet('系列', 2, ['店铺', '系列'])
    r = 2
    present = defaultdict(set)
    for (st, sku), info in skus.items(): present[st].add(d.sku_series(sku))
    for st in stores:
        for s in [x for x in series_list if x in present[st]]:
            n = write_summary(ws, r, [st, s], nk, c0, f',{AR("A")},$A{r},{AR("C")},$B{r}', SRC_A, f'$B{r}', f'$A{r}'); r += 1
        n = write_summary(ws, r, [st, '店铺合计'], nk, c0, f',{AR("A")},$A{r}', SRC_A, '"__none__"', f'$A{r}')
        ws.cell(row=r, column=nk + 1, value=f'=IFERROR(SUMPRODUCT(({AR("A")}=$A{r})*({AR("I")}="成熟")*{AR("L")}*IFERROR(IF(INDEX({TGT_RNG},MATCH({AR("C")},{SER_RNG},0))="",参数!$B$6,INDEX({TGT_RNG},MATCH({AR("C")},{SER_RNG},0))),参数!$B$6))/{L(c0 + 2)}{r},0)')
        for ci_ in range(1, n + 1): ws.cell(row=r, column=ci_).fill = SUBFILL; ws.cell(row=r, column=ci_).font = Font(name='Arial', bold=True, color=ws.cell(row=r, column=ci_).font.color)
        r += 2
    finish(ws, nk, c0, r - 2, n, [9, 9] + W_BASE)
    ws.cell(row=r, column=1, value='数据源=广告商品日报（SKU级，按SKU映射的系列汇总）。判断：点击<参数最低点击=样本不足；ACOS≤目标=达标；≤目标×1.3=偏高；否则超标。周ACOS超目标×1.3标红。').font = NOTE

    # ================= SKU =================
    ws, hdr, nk, c0 = summary_sheet('SKU', 3, ['店铺', '系列', 'SKU', '品名'])
    r = 2
    for (st, sku), info in sorted(skus.items(), key=lambda kv: (store_order.get(kv[0][0], 9), ser_idx(d.sku_series(kv[0][1])), -kv[1]['cost'])):
        n = write_summary(ws, r, [st, f'=IFERROR(INDEX({SK_SER},MATCH($C{r},{SK_SKU},0)),"")', sku, d.sku_name(sku)], nk, c0, f',{AR("A")},$A{r},{AR("D")},$C{r}', SRC_A, f'$B{r}', f'$A{r}'); r += 1
    finish(ws, nk, c0, r - 1, n, [9, 9, 30, 36] + W_BASE)

    # ================= 活动配置 =================
    wsG = wb.create_sheet('活动配置', 4)
    header(wsG, ['店铺', '系列', '活动名称', '投放类型', '竞价策略', '顶部溢价', '详情页溢价', '日预算', '当前状态', '成熟期日均花费', '预算使用率', '跑满天数(≥90%)', '成熟期日均销售额', '成熟期ACOS', '目标ACOS', '目标日花费(日均销售额×目标ACOS)', '日预算/目标日花费', '新鲜周日均花费', '新鲜周使用率', '建议', '广告活动ID'])
    r = 2
    for cid in cids:
        cfg = d.cfg_by_id.get(cid)
        if not cfg or cfg.get('状态') != 'enabled': continue
        st = camp_store[cid]
        vals = [st, cm(CM_SER, f'$U{r}'), cm(CM_ALIAS, f'$U{r}'), cfg.get('投放类型', ''), cfg.get('竞价策略', ''), cfg['顶部溢价%'] / 100, cfg['详情页溢价%'] / 100, cfg['日预算'], cfg.get('当前状态', ''),
                f'=SUMIFS({DR("J")},{DR("A")},$A{r},{DR("C")},$U{r},{DR("G")},"成熟")/(参数!$B$9*7)',
                f'=IF(H{r}=0,"",J{r}/H{r})',
                f'=COUNTIFS({DR("A")},$A{r},{DR("C")},$U{r},{DR("G")},"成熟",{DR("J")},">="&H{r}*0.9)',
                f'=SUMIFS({DR("L")},{DR("A")},$A{r},{DR("C")},$U{r},{DR("G")},"成熟")/(参数!$B$9*7)',
                f'=IF(M{r}=0,0,J{r}/M{r})',
                tgt_formula(f'$B{r}'),
                f'=M{r}*O{r}',
                f'=IF(P{r}=0,"",H{r}/P{r})',
                f'=SUMIFS({DR("J")},{DR("A")},$A{r},{DR("C")},$U{r},{DR("G")},"新鲜")/7',
                f'=IF(H{r}=0,"",R{r}/H{r})',
                f'=IF(J{r}<3,"花费太小",IF(AND(K{r}>=0.9,N{r}<=O{r}),"预算限制了好活动→加预算",IF(AND(K{r}>=0.9,N{r}>O{r}*1.3),"顶着预算亏钱→降出价/溢价",IF(AND(K{r}<0.4,N{r}<=O{r}),"预算闲置的好活动→加出价",IF(AND(F{r}>=0.5,N{r}>O{r}),"顶部溢价过高→降溢价",IF(AND(Q{r}<>"",Q{r}>2,N{r}>O{r}),"预算远超目标日花费→先降出价再收预算",""))))))',
                cid]
        put(wsG, r, vals, {6: 'pct', 7: 'pct', 8: 'money', 10: 'money', 11: 'pct', 12: 'int', 13: 'money', 14: 'pct', 15: 'pct', 16: 'money', 17: 'x', 18: 'money', 19: 'pct'})
        wsG.cell(row=r, column=21).font = GRAY; r += 1
    n_g = r - 1
    widths(wsG, [9, 9, 36, 8, 14, 9, 9, 9, 11, 11, 10, 10, 11, 10, 9, 13, 10, 11, 10, 34, 17]); wsG.auto_filter.ref = f'A1:U{n_g}'
    wsG.conditional_formatting.add(f'K2:K{n_g}', FormulaRule(formula=['AND(K2<>"",K2>=0.9)'], fill=RED))
    wsG.conditional_formatting.add(f'K2:K{n_g}', FormulaRule(formula=['AND(K2<>"",K2<0.4)'], fill=AMBER))
    wsG.conditional_formatting.add(f'F2:F{n_g}', FormulaRule(formula=['F2>=0.5'], fill=RED))
    wsG.conditional_formatting.add(f'N2:N{n_g}', FormulaRule(formula=['AND(N2>O2*1.3,N2>0)'], fill=RED))
    wsG.conditional_formatting.add(f'N2:N{n_g}', FormulaRule(formula=['AND(N2<=O2,N2>0)'], fill=GREENF))
    wsG.conditional_formatting.add(f'Q2:Q{n_g}', FormulaRule(formula=['AND(Q2<>"",Q2>2)'], fill=AMBER))
    wsG.cell(row=n_g + 2, column=1, value='数据源=领星 spCampaigns（当前设置，每日同步）+ 活动日报。日预算是上限不是预期：亚马逊允许单日超支，按月平均不超过日预算。日均花费=成熟期花费÷35天；预算使用率=日均花费÷日预算（预算最近被改过的活动，这个比值会失真，看"跑满天数"）；目标日花费=成熟期日均销售额×目标ACOS，即按目标 ACOS 这个活动"配得上"花多少；日预算/目标日花费 >2 说明预算上限远高于销售能支撑的花费。判断顺序：先看 ACOS 是否达标，再看是否跑满——跑满且达标才加预算；跑满且超标先降出价；用不完且达标加出价；用不完且超标降预算或合并。').font = NOTE

    # ================= 广告位周报 + 广告位 =================
    plc = sorted(d.pl, key=lambda x: (store_order.get(x['店铺'], 9), x['周(周一)'], -x['花费']))
    wsPL = wb.create_sheet('广告位周报'); header(wsPL, ['店铺', '周(周一)', '广告活动ID', '系列', '活动名称', '广告位', '成熟', '展示量', '点击量', '花费', '订单数', '销售额', '销量'])
    for i, x in enumerate(plc, 2):
        put(wsPL, i, [x['店铺'], x['周(周一)'], x['广告活动ID'], cm(CM_SER, f'C{i}'), cm(CM_ALIAS, f'C{i}'), x.get('广告位', ''), f'=IF(B{i}+6<=参数!$B$4,"成熟","新鲜")', x['展示量'], x['点击量'], x['花费'], x['订单数'], x['销售额'], x['销量']], {2: 'date', 8: 'int', 9: 'int', 10: 'money', 11: 'int', 12: 'money', 13: 'int'})
    n_pl = len(plc); widths(wsPL, [9, 11, 17, 9, 34, 16, 7] + [10] * 6); wsPL.auto_filter.ref = f'A1:M{n_pl + 1}'
    def PR(col): return f"广告位周报!${col}$2:${col}${n_pl + 1}"
    wsP2 = wb.create_sheet('广告位', 5)
    hdrP = ['店铺', '系列', '活动名称', '广告位', '溢价设置', '成熟期 展示量', '成熟期 点击量', '成熟期 花费', '成熟期 订单数', '成熟期 销售额', '成熟期 CTR', '成熟期 CPC', '成熟期 CVR', '成熟期 ACOS', '占活动花费%', '判断', '新鲜期 花费', '新鲜期 CPC', '新鲜期 占活动%'] + [f'花费 W{w[0].strftime("%m-%d")}' for w in WEEKS] + [f'CPC W{w[0].strftime("%m-%d")}' for w in WEEKS] + ['广告活动ID']
    header(wsP2, hdrP); C_ID2 = L(len(hdrP))
    def plc_row(r, cid_cell, crit, extra_keys, tgt_expr):
        vals = list(extra_keys)
        for c in 'HIJKL': vals.append(f'=SUMIFS({PR(c)},{PR("G")},"成熟"{crit})')
        vals += [f'=IF(F{r}=0,0,G{r}/F{r})', f'=IF(G{r}=0,0,H{r}/G{r})', f'=IF(G{r}=0,0,I{r}/G{r})', f'=IF(J{r}=0,0,H{r}/J{r})']
        camp_tot = f'SUMIFS({PR("J")},{PR("G")},"成熟",{PR("A")},$A{r}' + (f',{PR("C")},${C_ID2}{r}' if cid_cell else '') + ')'
        vals.append(f'=IF({camp_tot}=0,0,H{r}/{camp_tot})')
        vals.append(f'=IF(G{r}<参数!$B$8,"样本不足",IF(J{r}=0,"零单",IF(N{r}<={tgt_expr},"达标",IF(N{r}<={tgt_expr}*1.3,"偏高","超标"))))')
        vals.append(f'=SUMIFS({PR("J")},{PR("G")},"新鲜"{crit})')
        vals.append(f'=IF(SUMIFS({PR("I")},{PR("G")},"新鲜"{crit})=0,0,Q{r}/SUMIFS({PR("I")},{PR("G")},"新鲜"{crit}))')
        camp_fresh = f'SUMIFS({PR("J")},{PR("G")},"新鲜",{PR("A")},$A{r}' + (f',{PR("C")},${C_ID2}{r}' if cid_cell else '') + ')'
        vals.append(f'=IF({camp_fresh}=0,0,Q{r}/{camp_fresh})')
        for pk, a, b in WEEKS: vals.append(f'=SUMIFS({PR("J")},{PR("B")},{DATE(a)}{crit})')
        for j, (pk, a, b) in enumerate(WEEKS): vals.append(f'=IF(SUMIFS({PR("I")},{PR("B")},{DATE(a)}{crit})=0,0,{L(20 + j)}{r}/SUMIFS({PR("I")},{PR("B")},{DATE(a)}{crit}))')
        vals.append(cid_cell or '')
        fm = {5: 'pct', 6: 'int', 7: 'int', 8: 'money', 9: 'int', 10: 'money', 11: 'pct', 12: 'money', 13: 'pct', 14: 'pct', 15: 'pct', 17: 'money', 18: 'money', 19: 'pct'}
        for j in range(len(WEEKS)): fm[20 + j] = 'money'; fm[20 + len(WEEKS) + j] = 'money'
        put(wsP2, r, vals, fm)
    r = 2
    pl_by_cid = defaultdict(set)
    for x in plc: pl_by_cid[x['广告活动ID']].add(x.get('广告位', ''))
    for st in stores:
        for pl in PL_ORDER:
            plc_row(r, None, f',{PR("A")},$A{r},{PR("F")},$D{r}', [st, '全店', '店铺合计', pl, ''], '参数!$B$6')
            for ci_ in range(1, len(hdrP) + 1): wsP2.cell(row=r, column=ci_).fill = SUBFILL; wsP2.cell(row=r, column=ci_).font = Font(name='Arial', bold=True, color=wsP2.cell(row=r, column=ci_).font.color)
            r += 1
        r += 1
        for cid in cids:
            if camp_store[cid] != st or not pl_by_cid.get(cid): continue
            cfg = d.cfg_by_id.get(cid, {})
            for pl in sorted(pl_by_cid[cid], key=lambda p: PL_ORDER.index(p) if p in PL_ORDER else 9):
                boost = (cfg.get('顶部溢价%', 0) / 100 if pl == '首页顶部' else (cfg.get('详情页溢价%', 0) / 100 if pl == '商品详情页' else 0))
                plc_row(r, cid, f',{PR("A")},$A{r},{PR("C")},${C_ID2}{r},{PR("F")},$D{r}', [st, cm(CM_SER, f'${C_ID2}{r}'), cm(CM_ALIAS, f'${C_ID2}{r}'), pl, boost], tgt_inline(f'$B{r}'))
                wsP2.cell(row=r, column=5).number_format = FMT['pct']; wsP2.cell(row=r, column=len(hdrP)).font = GRAY
                r += 1
            r += 1
    widths(wsP2, [9, 9, 34, 16, 9] + [9] * 5 + [8] * 4 + [8, 10, 10, 9, 10] + [9] * (2 * len(WEEKS)) + [17])
    wsP2.auto_filter.ref = f'A1:{C_ID2}{r}'
    wsP2.conditional_formatting.add(f'P2:P{r}', FormulaRule(formula=['OR(P2="超标",P2="零单")'], fill=RED))
    wsP2.conditional_formatting.add(f'P2:P{r}', FormulaRule(formula=['P2="偏高"'], fill=AMBER))
    wsP2.conditional_formatting.add(f'P2:P{r}', FormulaRule(formula=['P2="达标"'], fill=GREENF))
    wsP2.conditional_formatting.add(f'E2:E{r}', FormulaRule(formula=['E2>=0.5'], fill=RED))
    wsP2.cell(row=r + 1, column=1, value='数据源=广告位周报（领星 campaignPlacementReports）+ 活动配置里的溢价设置。蓝底=店铺合计（各位置占全店花费%）；活动行的"占活动花费%"是该位置占该活动的比例。读法：首页顶部 CVR 高、ACOS 低但占比低 → 给达标的活动加顶部溢价；某活动顶部 CPC 周环比跳升 + 溢价≥50% → 降溢价。').font = NOTE

    # ================= 活动·投放 =================
    ws, hdr, nk, c0 = summary_sheet('活动·投放', 6, ['店铺', '系列', '活动名称', '层级', '投放类型', '投放文本', '匹配类型', '当前出价/预算', '建议出价'])
    NCOL = nk + 15 + 23 + 2 * len(WEEKS); C_ID, C_KEY = L(NCOL + 1), L(NCOL + 2)
    ws.cell(row=1, column=NCOL + 1, value='广告活动ID').font = HFONT; ws.cell(row=1, column=NCOL + 1).fill = HFILL
    ws.cell(row=1, column=NCOL + 2, value='投放键').font = HFONT; ws.cell(row=1, column=NCOL + 2).fill = HFILL
    SRC_T = {'cols': 'LMNOPQ', 'o1d': None, 's1d': None, 'rng': TR, 'mature': TR('K'), 'date': TR('B'), 'store': TR('A')}
    grp = defaultdict(dict)
    for t in trows:
        k = f"{t.get('投放类型', '')}|{t.get('投放ID', '')}"; txt = t.get('投放文本', '')
        g = grp[t['广告活动ID']].setdefault(k, {'ttype': t.get('投放类型', ''), 'text': AUTO_NAME.get(txt, txt), 'match': ('自动' if txt in AUTO_NAME else ('商品定位' if t.get('投放类型') == '商品定位' else t.get('匹配类型', ''))), 'cost': 0.0, 'id': str(t.get('投放ID', ''))}); g['cost'] += t['花费']
    r = 2
    for cid in cids:
        st = camp_store[cid]; cfg = d.cfg_by_id.get(cid, {})
        strat = (cfg.get('竞价策略', '') + (f" / 顶部+{cfg['顶部溢价%']:.0f}%" if cfg.get('顶部溢价%') else '')) if cfg else ''
        keyvals = [st, cm(CM_SER, f'${C_ID}{r}'), cm(CM_ALIAS, f'${C_ID}{r}'), '活动', cm(CM_TYPE, f'${C_ID}{r}'), f'="主SKU: "&IFERROR(INDEX({CM_MSKU},MATCH(${C_ID}{r},{CM_ID},0))&"","")', strat, (f"预算 ${cfg['日预算']:g}" if cfg else ''), '']
        n = write_summary(ws, r, keyvals, nk, c0, f',{DR("A")},$A{r},{DR("C")},${C_ID}{r}', SRC_D, f'$B{r}', f'$A{r}')
        ws.cell(row=r, column=NCOL + 1, value=cid).font = GRAY
        for ci_ in range(1, n + 1): ws.cell(row=r, column=ci_).fill = SUBFILL; ws.cell(row=r, column=ci_).font = Font(name='Arial', bold=True, color=ws.cell(row=r, column=ci_).font.color)
        crow = r; r += 1
        for key, g in sorted(grp.get(cid, {}).items(), key=lambda kv: -kv[1]['cost']):
            bid = d.bid_by_id.get(g['id']); bidv = bid['当前出价'] if bid else None; paused = bool(bid) and bid.get('状态') == 'paused'
            keyvals = [st, cm(CM_SER, f'${C_ID}{r}'), cm(CM_ALIAS, f'${C_ID}{r}'), '  └ 投放组', g['ttype'], g['text'], g['match'], bidv, ('已暂停' if paused else (f'=IF(H{r}="","",IF({L(c0 + 13)}{r}="超标",ROUND(H{r}*0.7,2),IF({L(c0 + 13)}{r}="偏高",ROUND(H{r}*0.85,2),IF({L(c0 + 13)}{r}="零单",ROUND(H{r}*0.5,2),IF({L(c0 + 13)}{r}="达标",ROUND(H{r}*1.1,2),H{r})))))' if bidv is not None else ''))]
            n = write_summary(ws, r, keyvals, nk, c0, f',{TR("A")},$A{r},{TR("C")},${C_ID}{r},{TR("Z")},${C_KEY}{r}', SRC_T, f'$B{r}', f'$A{r}')
            ws.cell(row=r, column=NCOL + 1, value=cid).font = GRAY; ws.cell(row=r, column=NCOL + 2, value=key).font = GRAY
            r += 1
        if r - 1 > crow: ws.row_dimensions.group(crow + 1, r - 1, outline_level=1, hidden=False)
    finish(ws, nk, c0, r - 1, NCOL + 2, [9, 9, 34, 9, 12, 40, 14, 14, 9] + W_BASE + [17, 30])
    for row in ws.iter_rows(min_row=2, max_row=r - 1, min_col=8, max_col=9):
        for c in row:
            if isinstance(c.value, (int, float)) or (isinstance(c.value, str) and c.value.startswith('=IF(H')): c.number_format = FMT['money']
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.cell(row=r + 1, column=1, value='活动行（蓝底）=活动日报汇总，含 1d 口径与预估，第8列显示竞价策略/顶部溢价、第9列显示日预算；投放组行=投放周报汇总，第8列=领星里的当前出价，第9列=建议出价（规则：超标×0.7 / 偏高×0.85 / 零单×0.5 / 达标×1.1 / 样本不足不动）。左侧 +/− 可折叠。').font = NOTE

    # ================= 搜索词周报 + 搜索词 =================
    qrows = sorted(d.qry, key=lambda x: (store_order.get(x['店铺'], 9), x['周(周一)'], -x['花费']))
    wsQ = wb.create_sheet('搜索词周报'); header(wsQ, ['店铺', '周(周一)', '广告活动ID', '系列', '广告组ID', '投放类型', '投放ID', '用户搜索词', '投放词/投放组', '匹配类型', '成熟', '展示量', '点击量', '花费', '订单数', '销售额', '销量'])
    for i, x in enumerate(qrows, 2):
        put(wsQ, i, [x['店铺'], x['周(周一)'], x['广告活动ID'], cm(CM_SER, f'C{i}'), str(x.get('广告组ID', '')), x.get('投放类型', ''), str(x.get('投放ID', '')), x.get('用户搜索词', ''), x.get('投放词/投放组', ''), x.get('匹配类型', ''), f'=IF(B{i}+6<=参数!$B$4,"成熟","新鲜")', x['展示量'], x['点击量'], x['花费'], x['订单数'], x['销售额'], x['销量']], {2: 'date', 12: 'int', 13: 'int', 14: 'money', 15: 'int', 16: 'money', 17: 'int'})
        wsQ.cell(row=i, column=26, value=f"{x.get('用户搜索词', '')}|{x.get('投放词/投放组', '')}|{x.get('匹配类型', '')}").font = GRAY
    wsQ.cell(row=1, column=26, value='投放键').font = HFONT; wsQ.cell(row=1, column=26).fill = HFILL
    n_q = len(qrows); widths(wsQ, [9, 11, 17, 9, 17, 12, 17, 34, 22, 8, 7] + [10] * 6); wsQ.auto_filter.ref = f'A1:Q{n_q + 1}'
    def QR(col): return f"搜索词周报!${col}$2:${col}${n_q + 1}"
    clicks_tot = defaultdict(float); cost_tot = defaultdict(float)
    for x in qrows:
        k = (x['店铺'], x['广告活动ID'], x.get('用户搜索词', ''), x.get('投放词/投放组', ''), x.get('匹配类型', '')); clicks_tot[k] += x['点击量']; cost_tot[k] += x['花费']
    uniq = [k for k in clicks_tot if clicks_tot[k] >= 3]
    wsQ2 = wb.create_sheet('搜索词', 7)
    hdrQ = ['店铺', '系列', '活动名称', '用户搜索词', '投放词/投放组', '匹配类型', '已否定?', '目标ACOS', '成熟期 展示量', '成熟期 点击量', '成熟期 花费', '成熟期 订单数', '成熟期 销售额', '成熟期 CTR', '成熟期 CPC', '成熟期 CVR', '成熟期 ACOS', '判断', '新鲜期 展示量', '新鲜期 点击量', '新鲜期 CTR', '新鲜期 花费', '新鲜期 CPC', '新鲜期 订单(未熟)', '全期花费', '广告活动ID', '投放键']
    header(wsQ2, hdrQ)
    r = 2
    for k in sorted(uniq, key=lambda k: (store_order.get(k[0], 9), ser_idx(d.camp_series(k[1])), d.camp_name(k[1]), -cost_tot[k])):
        st, cid, q, tt, match = k
        crit = f',{QR("A")},$A{r},{QR("C")},$Z{r},{QR("Z")},$AA{r}'
        vals = [st, cm(CM_SER, f'$Z{r}'), cm(CM_ALIAS, f'$Z{r}'), q, tt, match, d.neg_status(st, cid, q), tgt_formula(f'$B{r}')]
        for c in 'LMNOP': vals.append(f'=SUMIFS({QR(c)},{QR("K")},{MATURE}{crit})')
        vals += [f'=IF(I{r}=0,0,J{r}/I{r})', f'=IF(J{r}=0,0,K{r}/J{r})', f'=IF(J{r}=0,0,L{r}/J{r})', f'=IF(M{r}=0,0,K{r}/M{r})']
        vals.append(f'=IF(J{r}<参数!$B$8,IF(AND(K{r}>=参数!$B$7,L{r}=0),"零单高花费","样本不足"),IF(M{r}=0,"零单",IF(Q{r}<=H{r},"达标",IF(Q{r}<=H{r}*1.3,"偏高","超标"))))')
        vals += [f'=SUMIFS({QR("L")},{QR("K")},{FRESH}{crit})', f'=SUMIFS({QR("M")},{QR("K")},{FRESH}{crit})', f'=IF(S{r}=0,0,T{r}/S{r})', f'=SUMIFS({QR("N")},{QR("K")},{FRESH}{crit})', f'=IF(T{r}=0,0,V{r}/T{r})', f'=SUMIFS({QR("O")},{QR("K")},{FRESH}{crit})']
        vals += [f'=K{r}+V{r}', cid, f'{q}|{tt}|{match}']
        put(wsQ2, r, vals, {8: 'pct', 9: 'int', 10: 'int', 11: 'money', 12: 'int', 13: 'money', 14: 'pct', 15: 'money', 16: 'pct', 17: 'pct', 19: 'int', 20: 'int', 21: 'pct', 22: 'money', 23: 'money', 24: 'int', 25: 'money'})
        wsQ2.cell(row=r, column=26).font = GRAY; wsQ2.cell(row=r, column=27).font = GRAY
        r += 1
    widths(wsQ2, [9, 9, 34, 30, 22, 8, 16, 9] + [9] * 5 + [8] * 4 + [11] + [9] * 6 + [10, 17, 30]); wsQ2.auto_filter.ref = f'A1:AA{r - 1}'
    wsQ2.conditional_formatting.add(f'R2:R{r - 1}', FormulaRule(formula=['OR(R2="超标",R2="零单",R2="零单高花费")'], fill=RED))
    wsQ2.conditional_formatting.add(f'R2:R{r - 1}', FormulaRule(formula=['R2="偏高"'], fill=AMBER))
    wsQ2.conditional_formatting.add(f'R2:R{r - 1}', FormulaRule(formula=['R2="达标"'], fill=GREENF))
    wsQ2.cell(row=r + 1, column=1, value='数据源=搜索词周报（关键词活动 + 自动/商品定位活动的用户搜索词）。本表只列全期点击≥3次的搜索词（其余在搜索词周报里筛）。判断同投放表：成熟期点击≥最低样本且零单 → 否定候选；ACOS超标 → 降价候选。"已否定?"列来自领星否定词表（每日同步）。否定动作在领星：广告 → 搜索词 → 勾选 → 添加否定关键词/否定商品。').font = NOTE

    # ================= 优化方案 =================
    wsO = wb.create_sheet('优化方案', 0)
    wsO['A1'] = f'优化方案（成熟期 {START.strftime("%m/%d")}–{MEND.strftime("%m/%d")} 五周 + 新鲜周 {FRESH_A.strftime("%m/%d")}–{END.strftime("%m/%d")}；生成日 {SNAP}）'; wsO['A1'].font = Font(name='Arial', bold=True, size=12)
    wsO['A2'] = '本页来自飞书「优化记录」表：自动规则每周生成新建议（来源=自动），你补充的记人工。状态/执行日期/备注请在飞书表里改，下周的周报会自动带过来；本文件里改只对这一份有效。执行两周后再回看成熟数据评估效果。'; wsO['A2'].font = NOTE
    header(wsO, ['本周新增', '提出周', '优先级', '店铺', '系列', '活动名称', '层级', '对象', '动作', '建议值', '依据（提出时）', '最近数据（本周）', '状态', '执行日期', '备注', '来源'], row=4)
    opt = [o for o in P.get('opt_rows', []) if o.get('状态') != '放弃']
    pri_order = {'P1止血': 0, 'P1否词': 1, 'P2收割加码': 2, 'P2预算': 3, 'P3结构': 4, '其他': 5}
    opt.sort(key=lambda o: (STATUS_ORDER.get(o.get('状态'), 9), pri_order.get(o.get('优先级'), 9), store_order.get(o.get('店铺'), 9), ser_idx(o.get('系列', ''))))
    r = 5
    for o in opt:
        new = (o.get('提出周') == P['week_of']) and o.get('状态') == '待处理'
        vals = ['★ 新' if new else '', o.get('提出周'), o.get('优先级', ''), o.get('店铺', ''), o.get('系列', ''), o.get('活动名称', ''), o.get('层级', ''), o.get('对象', ''), o.get('动作', ''), o.get('建议值', ''), o.get('依据', ''), o.get('最近数据', ''), o.get('状态', ''), o.get('执行日期'), o.get('备注', ''), o.get('来源', '')]
        for ci_, v in enumerate(vals, 1):
            c = wsO.cell(row=r, column=ci_, value=v); c.font = BF; c.border = BORDER; c.alignment = Alignment(wrap_text=True, vertical='top')
            if ci_ in (2, 14): c.number_format = FMT['date']
        pr = str(o.get('优先级', ''))[:2]
        wsO.cell(row=r, column=3).fill = PatternFill('solid', fgColor={'P1': 'F8CBAD', 'P2': 'C6EFCE', 'P3': 'DDEBF7'}.get(pr, 'FFFFFF'))
        st_ = o.get('状态', '')
        wsO.cell(row=r, column=13).fill = {'待处理': YELLOW, '已执行': GREENF, '观察中': AMBER}.get(st_, PatternFill())
        r += 1
    widths(wsO, [7, 11, 10, 8, 7, 30, 7, 30, 34, 12, 46, 40, 8, 11, 24, 6])
    rN = r + 2
    wsO.cell(row=rN, column=1, value='重点注意').font = Font(name='Arial', bold=True, size=12)
    NOTES = [
     ('1', '只用成熟期下结论', f'成熟期 = {START.strftime("%m/%d")}–{MEND.strftime("%m/%d")} 五个完整周。新鲜周的订单和 ACOS 还会往上走（1 天归因通常只占最终订单的七成左右），新鲜期只看花费 / 点击 / CPC 和 1d 口径，不要拿"新鲜期 ACOS(未熟)"做决定。'),
     ('2', '样本量', '落地灯转化率约 1.5–2%，一个投放组 / 搜索词要 ≥100 次点击才能判断（参数表可改）。自动建议里的暂停动作都是大样本零单或 ACOS 翻倍的；小样本零单只降价不暂停。'),
     ('3', '改动节奏', '一次改动后至少等 14 天再评估（7 天归因 + 7 天成熟）；同一个活动别在两周内叠加改第二次，否则分不清是哪个改动起的作用。在飞书优化记录里填执行日期，两周后的周报会把执行前后的成熟数据对上。'),
     ('4', '否词的层级', '自动活动的否定加在"活动级"（对所有投放组生效）；否定精准只挡完全一样的词，否定词组会挡所有包含该词的搜索词——大词如 "floor lamp" 别用否定词组。ASIN 类搜索词（b0xxxx）用"否定商品"。'),
     ('5', '自家 ASIN 互相付费', '搜索词里出现自家 ASIN，说明自动广告把你的一个产品打在另一个产品的详情页上。转化好的可以留，零单的否掉。'),
     ('6', '大词的内部竞价', '同一个大词在多个活动同时跑，活动之间互相抬价。原则：一个大词只留 1–2 个转化最好的活动跑自动，其余否定；真正要主攻的词单独开精准活动控价。优化记录里 P3结构 类会列出跨活动的重复词。'),
     ('7', '长尾花费', '否词只管得到点击≥3次的词；四成左右的花费散在几千个 1–2 次点击的长尾词上，只能靠投放组出价和广告位溢价控——所以降组出价往往比否词省的钱更多。'),
     ('8', '预算不是预期', '日预算是上限。看活动配置表：跑满且达标才加预算；跑满且超标先降出价；用不完且达标加出价；用不完且超标降预算或合并。预算刚改过的活动"预算使用率"会失真，以"跑满天数"为准。'),
     ('9', '目标 ACOS', '各系列目标 ACOS 在飞书 SKU映射 表的 目标ACOS% 列填，留空按默认 30% 判断。客单价不同的系列盈亏平衡 ACOS 差别很大，这个值填了判断才准。'),
    ]
    header(wsO, ['#', '事项', '说明'], row=rN + 1)
    for j, (a, b, c) in enumerate(NOTES):
        for ci_, v in enumerate([a, b, c], 1):
            x = wsO.cell(row=rN + 2 + j, column=ci_, value=v); x.font = BF; x.border = BORDER; x.alignment = Alignment(wrap_text=True, vertical='top')
        wsO.merge_cells(start_row=rN + 2 + j, start_column=3, end_row=rN + 2 + j, end_column=11); wsO.row_dimensions[rN + 2 + j].height = 48
    wsO.freeze_panes = 'A5'

    # ================= 说明 =================
    ws = wb.create_sheet('说明', 0)
    notes = [['SP广告周报：整体 → 系列 → SKU → 活动配置 / 广告位 → 活动·投放 → 搜索词 → 优化方案'],
     ['这张表的逻辑', '从上往下逐级下钻，每一级回答一个问题，最后落到动作。① 整体（店铺 × 周）：广告整体在变好还是变坏；② 系列：钱投在哪个系列、哪个系列拖后腿；③ SKU：同一系列里哪个 SKU 效率高；④ 活动配置 / 广告位：预算、竞价策略、溢价和三个广告位的效果；⑤ 活动·投放：具体到活动和投放组，出价、预算、暂停在这一层定；⑥ 搜索词：否词、收割、内部竞价在这一层定；⑦ 优化方案：规则自动生成 + 你在飞书里维护状态的动作清单。每一级都同时看成熟期（下结论）和新鲜期（监控）。'],
     ['Sheet 分类', '【看板层】整体 / 系列 / SKU / 活动配置 / 广告位 / 活动·投放 / 搜索词 —— 全部是 SUMIFS 公式，改原始数据或参数会重算。【结论层】优化方案 —— 来自飞书优化记录表。【参数层，可改】参数 / SKU映射 / 活动映射 —— 黄色单元格可改；长期修改请改飞书对应表。【原始层，不用看】活动日报 / 广告商品日报 / 投放周报 / 搜索词周报 / 广告位周报 —— 飞书表原样落地。'],
     ['数据', f'飞书多维表格「领星SP广告数据」（领星 OpenAPI 每天 06:00 同步），{START} ~ {END}（6 个完整周：5 周成熟 + 1 周新鲜），HLZ-US + YYK-US，仅 Sponsored Products；本文件生成于 {SNAP}，每周三自动生成新文件，上周文件不改。'],
     ['成熟/新鲜', f'成熟期 = 最近 5 个完整周（周一~周日）{START} ~ {MEND}，共 35 天，每周期末都已过 7 天归因 + 回传延迟，可下结论。{FRESH_A} ~ {END} 为新鲜周，只看花费/点击/CPC 和 1d 口径。'],
     ['新鲜期列怎么读', '"比周均"三个量变列（展示/点击/花费）= 新鲜周 ÷ 成熟期周平均 − 1，正数只代表比平均多，本身不分好坏：花费↑点击↑CPC↑=买了更多更贵的流量；花费↑点击↓CPC↑=同样的流量更贵了（先查投放组和广告位）；花费↓点击↓=量缩了（查预算是否跑满、出价是否被降）。CPC 正数=变贵=坏；CTR 正数（百分点）=好；1dACOS变化 负数=好；预估ACOS 绿=达标 红=超标。'],
     ['CPC涨了怎么查', '按 系列 → 活动 → 投放组 → 广告位 → 搜索词 逐层下钻：①活动·投放表看哪个活动 CPC 涨，展开它的投放组；②广告位表看该活动首页顶部是否突然放大；③搜索词周报筛该活动 + 新鲜周，按花费排序，对比同一个词成熟期的 CPC。'],
     ['判断规则', '点击 < 最低样本（默认100）→ 样本不足；ACOS ≤ 目标 → 达标；≤ 目标×1.3 → 偏高；否则超标；成熟期零单 → 零单。目标ACOS按系列在参数表（来源飞书 SKU映射）。'],
     ['为什么用 5 周而不是 30 天', '周末效应——30 天窗口里周末个数会在 8–10 之间变，5 个完整周每次都是 5 个周末，环比可比。']]
    for i, row in enumerate(notes, 1):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=i, column=j, value=v); c.font = Font(name='Arial', bold=(i == 1 or j == 1)); c.alignment = Alignment(wrap_text=True, vertical='top')
    widths(ws, [16, 120])
    wb.move_sheet('参数', offset=-(wb.index(wb['参数']) - 2))
    wb.save(out_path)
    stats_ = {'camp': n_d, 'pad': n_a, 'tgt': n_t, 'qry': n_q, 'pl': n_pl, 'sku': n_s, 'campaigns': n_c, 'opt': len(opt), 'sheets': list(wb.sheetnames)}
    wb.close(); del wb
    return stats_
    return {'camp': n_d, 'pad': n_a, 'tgt': n_t, 'qry': n_q, 'pl': n_pl, 'sku': n_s, 'campaigns': n_c, 'opt': len(opt), 'sheets': wb.sheetnames}
