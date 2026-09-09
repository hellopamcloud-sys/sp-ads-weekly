"""飞书 open-api 轻量客户端：tenant token、多维表格读写、云文档上传、群机器人卡片"""
import json, os, time, datetime
import requests

BASE = 'https://open.feishu.cn/open-apis'
APP_ID = os.environ.get('FS_APP_ID', '')
APP_SECRET = os.environ.get('FS_APP_SECRET', '')
_tok = {'v': None, 'exp': 0}


def token():
    if _tok['v'] and time.time() < _tok['exp'] - 60:
        return _tok['v']
    r = requests.post(f'{BASE}/auth/v3/tenant_access_token/internal', json={'app_id': APP_ID, 'app_secret': APP_SECRET}, timeout=30).json()
    if not r.get('tenant_access_token'):
        raise RuntimeError(f'feishu token failed: {r}')
    _tok['v'] = r['tenant_access_token']; _tok['exp'] = time.time() + r.get('expire', 7200)
    return _tok['v']


def call(method, path, body=None, query=None, retries=3, **kw):
    for i in range(retries):
        r = requests.request(method, BASE + path, json=body, params=query, headers={'Authorization': 'Bearer ' + token()}, timeout=60, **kw)
        try:
            j = r.json()
        except Exception:
            j = {'code': r.status_code, 'msg': r.text[:300]}
        if j.get('code') == 0:
            return j
        if i < retries - 1 and (j.get('code') in (99991400, 1254291, 1255002, 1254290) or r.status_code >= 500):
            time.sleep(1.5 * (i + 1)); continue
        raise RuntimeError(f'feishu {method} {path} failed: {json.dumps(j, ensure_ascii=False)[:500]}')
    return j


# ---------- 多维表格 ----------
def txt(v):
    """把飞书 search 返回的字段值转成纯值"""
    if isinstance(v, list):
        if v and isinstance(v[0], dict) and 'text' in v[0]:
            return ''.join(x.get('text', '') for x in v)
        return v
    if isinstance(v, dict) and 'text' in v:
        return v['text']
    return v


_fields_cache = {}


def field_names_of(app, table):
    if table not in _fields_cache:
        names, pt = [], None
        while True:
            r = call('GET', f'/bitable/v1/apps/{app}/tables/{table}/fields', None, dict({'page_size': 100}, **({'page_token': pt} if pt else {})))
            d = r.get('data') or {}; names += [f['field_name'] for f in d.get('items', [])]
            pt = d.get('page_token') if d.get('has_more') else None
            if not pt: break
        _fields_cache[table] = names
    return _fields_cache[table]


def read_all(app, table, field_names=None, filter_=None, page_size=200):
    """分页读取全部记录；飞书对单页返回体大小有限制（code 1254002），遇到时自动减小 page_size 重读"""
    body = {'field_names': field_names or field_names_of(app, table)}  # search 接口不传 field_names 时可能不返回记录
    if filter_: body['filter'] = filter_
    while True:
        out, pt = [], None
        try:
            while True:
                q = {'page_size': page_size}
                if pt: q['page_token'] = pt
                r = call('POST', f'/bitable/v1/apps/{app}/tables/{table}/records/search', body, q)
                d = r.get('data') or {}
                for it in d.get('items', []):
                    f = {k: txt(v) for k, v in (it.get('fields') or {}).items()}
                    f['_rid'] = it.get('record_id'); out.append(f)
                pt = d.get('page_token') if d.get('has_more') else None
                if not pt: return out
        except RuntimeError as e:
            if '1254002' in str(e) and page_size > 25:
                page_size //= 2; continue
            raise


def batch_create(app, table, records):
    n = 0
    for i in range(0, len(records), 500):
        r = call('POST', f'/bitable/v1/apps/{app}/tables/{table}/records/batch_create', {'records': records[i:i + 500]})
        n += len((r.get('data') or {}).get('records') or [])
    return n


def batch_update(app, table, records):
    n = 0
    for i in range(0, len(records), 500):
        call('POST', f'/bitable/v1/apps/{app}/tables/{table}/records/batch_update', {'records': records[i:i + 500]}); n += len(records[i:i + 500])
    return n


def date_ms(d):
    """date -> 飞书日期字段值（UTC 正午毫秒，避免时区错天）"""
    if isinstance(d, str): d = datetime.date.fromisoformat(d)
    return int(datetime.datetime(d.year, d.month, d.day, 12, tzinfo=datetime.timezone.utc).timestamp() * 1000)


def ms_date(ms):
    if ms in (None, ''): return None
    return datetime.datetime.fromtimestamp(int(ms) / 1000, tz=datetime.timezone.utc).date()


# ---------- 云文档 ----------
def upload_file(path, folder_token, file_name=None):
    file_name = file_name or os.path.basename(path)
    size = os.path.getsize(path)
    with open(path, 'rb') as fh:
        r = requests.post(f'{BASE}/drive/v1/files/upload_all', headers={'Authorization': 'Bearer ' + token()},
                          data={'file_name': file_name, 'parent_type': 'explorer', 'parent_node': folder_token, 'size': str(size)},
                          files={'file': (file_name, fh, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}, timeout=300).json()
    if r.get('code') != 0: raise RuntimeError(f'upload failed: {r}')
    return r['data']['file_token']


def share_link(file_token, typ='file', entity='tenant_editable'):
    try:
        call('PATCH', f'/drive/v1/permissions/{file_token}/public', {'link_share_entity': entity, 'external_access_entity': 'open'}, {'type': typ})
    except Exception as e:
        try: call('PATCH', f'/drive/v1/permissions/{file_token}/public', {'link_share_entity': entity}, {'type': typ})
        except Exception as e2: print('share_link warn', e2)
    return f'https://kcn46c3x2v7r.feishu.cn/file/{file_token}'


# ---------- 群机器人 ----------
def send_card(webhook, title, lines, link=None, link_text='打开周报', color='blue'):
    elems = [{'tag': 'div', 'text': {'tag': 'lark_md', 'content': '\n'.join(lines)}}]
    if link:
        elems.append({'tag': 'action', 'actions': [{'tag': 'button', 'text': {'tag': 'plain_text', 'content': link_text}, 'type': 'primary', 'url': link}]})
    card = {'msg_type': 'interactive', 'card': {'config': {'wide_screen_mode': True}, 'header': {'title': {'tag': 'plain_text', 'content': title}, 'template': color}, 'elements': elems}}
    r = requests.post(webhook, json=card, timeout=30)
    return r.json() if r.content else {'status': r.status_code}
