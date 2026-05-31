"""
LTE Analyze for Asekapp (Mega analyze AI Edition 2026)
- Multi-file parsing (.nmf)
- Worst Cells & Network Anomalies Finder
- Dynamic EARFCN & KPI Filters
- Network Analysis + Recommendations
- Excel Export (6 sheets, conditional formatting)
- Synchronized Timeline Chart
- AI Assistant Илгиз Искендерович — powered by Gemini
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import requests
import json
from io import BytesIO

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
    XLSX_OK = True
except ImportError:
    XLSX_OK = False

# ──────────────────────────────────────────────────────────────────────────────
# PARSER
# ──────────────────────────────────────────────────────────────────────────────
def _f(v):
    try: return float(v) if v and v.strip() else None
    except: return None

def _i(v):
    try: return int(v) if v and v.strip() else None
    except: return None

def parse_nmf_lines(lines: list[str]) -> dict:
    """
    Field positions (0-indexed, comma-separated):
      GPS      : [1]=time [3]=lon [4]=lat [6]=gps_fix [7]=sats
      CELLMEAS : [1]=time [3]=rat [8]=earfcn [10]=pci
                 [11]=rsrp_dBm  [12]=rsrq×10→dB  [15]=sinr×10→dB  (rat=7 LTE)
      PPPRATE  : [1]=time [3]=dl_kbps [4]=ul_kbps
      DEVI     : [1]=time [3]=battery_pct [8]=speed_kmh
    """
    meta = {}
    gps_r, cell_r, ppp_r, devi_r = [], [], [], []

    for raw in lines:
        line = raw.strip()
        if not line or line.startswith('//'): continue

        if line.startswith('#'):
            p = line.split(',')
            k = p[0]
            def hv(i=3, _p=p): return _p[i].strip('"') if len(_p) > i else ''
            if k == '#PRODUCT':  meta['product']     = hv(3); meta['version'] = hv(4)
            elif k == '#DN':     meta['device']      = hv()
            elif k == '#EI':     meta['imei']        = hv()
            elif k == '#TS':     meta['test_script'] = hv()
            elif k == '#BF':     meta['map_file']    = hv()
            elif k == '#START':
                meta['start_time'] = p[1] if len(p) > 1 else ''
                meta['start_date'] = hv()
            elif k == '#STOP':   meta['stop_time']   = p[1] if len(p) > 1 else ''
            continue

        p = line.split(',')
        msg = p[0]

        if msg == 'GPS' and len(p) > 7:
            gps_r.append({'time': p[1], 'lon': _f(p[3]), 'lat': _f(p[4]),
                          'gps_fix': _i(p[6]), 'sats': _i(p[7])})

        elif msg == 'CELLMEAS' and len(p) > 15 and _i(p[3]) == 7:
            rq = _f(p[12]); sn = _f(p[15])
            cell_r.append({'time': p[1], 'earfcn': _i(p[8]), 'pci': _i(p[10]),
                           'rsrp': _f(p[11]),
                           'rsrq': rq / 10 if rq is not None else None,
                           'sinr': sn / 10 if sn is not None else None})

        elif msg == 'PPPRATE' and len(p) > 4:
            ppp_r.append({'time': p[1], 'dl_kbps': _f(p[3]), 'ul_kbps': _f(p[4])})

        elif msg == 'DEVI' and len(p) > 8:
            devi_r.append({'time': p[1], 'battery_pct': _f(p[3]),
                           'speed_kmh': _f(p[8])})

    def tdf(rows):
        df = pd.DataFrame(rows)
        if df.empty: return df
        df['time'] = pd.to_datetime(df['time'], format='%H:%M:%S.%f',
                                    errors='coerce').dt.time
        return df.dropna(subset=['time']).reset_index(drop=True)

    return {'meta': meta, 'gps': tdf(gps_r), 'cell': tdf(cell_r),
            'ppp': tdf(ppp_r), 'devi': tdf(devi_r)}


# ──────────────────────────────────────────────────────────────────────────────
# NETWORK ANALYSIS
# ──────────────────────────────────────────────────────────────────────────────
def analyze_network(df_cell, df_ppp, df_devi, threshold_rsrp, threshold_sinr):
    findings, recs = [], []
    stats = {}
    penalty = 0

    # RSRP
    if not df_cell.empty:
        rsrp = df_cell['rsrp'].dropna()
        stats.update({'rsrp_mean': round(rsrp.mean(),1), 'rsrp_min': round(rsrp.min(),1),
                      'rsrp_max': round(rsrp.max(),1),
                      'rsrp_good_pct': round((rsrp >= -90).sum()/len(rsrp)*100, 1)})
        if rsrp.mean() >= -80:
            findings.append(('✅','RSRP', f'Отличный сигнал {rsrp.mean():.1f} dBm'))
        elif rsrp.mean() >= -90:
            findings.append(('🟡','RSRP', f'Умеренный сигнал {rsrp.mean():.1f} dBm — ниже оптимума −80 dBm'))
            penalty += 10
            recs.append({'priority':'Средний','category':'Покрытие',
                         'issue': f'RSRP {rsrp.mean():.1f} dBm < −80 dBm',
                         'action':'Увеличить мощность TX или добавить small cell на маршруте'})
        else:
            findings.append(('🔴','RSRP', f'Слабый сигнал {rsrp.mean():.1f} dBm'))
            penalty += 30

        if rsrp.min() < threshold_rsrp:
            findings.append(('⚠️','RSRP мин', f'Критическое значение {rsrp.min():.1f} dBm на маршруте'))
            penalty += 15
            recs.append({'priority':'Высокий','category':'Покрытие',
                         'issue': f'Минимальный RSRP {rsrp.min():.1f} dBm',
                         'action':'RF-аудит маршрута, установить repeater/small cell в "белых пятнах"'})

    # RSRQ
    if not df_cell.empty:
        rsrq = df_cell['rsrq'].dropna()
        stats.update({'rsrq_mean': round(rsrq.mean(),1), 'rsrq_min': round(rsrq.min(),1),
                      'rsrq_max': round(rsrq.max(),1)})
        if rsrq.mean() >= -10:
            findings.append(('✅','RSRQ', f'Хорошее качество канала {rsrq.mean():.1f} dB'))
        elif rsrq.mean() >= -15:
            findings.append(('🟡','RSRQ', f'Умеренная интерференция RSRQ {rsrq.mean():.1f} dB'))
            penalty += 10
            recs.append({'priority':'Средний','category':'Интерференция',
                         'issue': f'RSRQ {rsrq.mean():.1f} dB (норма > −10 dB)',
                         'action':'Оптимизировать tilt антенн и план PCI reuse для EARFCN'})
        else:
            findings.append(('🔴','RSRQ', f'Высокая интерференция RSRQ {rsrq.mean():.1f} dB'))
            penalty += 25

    # SINR
    if not df_cell.empty:
        sinr = df_cell['sinr'].dropna()
        stats.update({'sinr_mean': round(sinr.mean(),1), 'sinr_min': round(sinr.min(),1),
                      'sinr_max': round(sinr.max(),1),
                      'sinr_good_pct': round((sinr >= 10).sum()/len(sinr)*100, 1)})
        if sinr.mean() >= 20:
            findings.append(('✅','SINR', f'Отличный SINR {sinr.mean():.1f} dB'))
        elif sinr.mean() >= 10:
            findings.append(('✅','SINR', f'Хороший SINR {sinr.mean():.1f} dB — 64QAM доступен'))
        elif sinr.mean() >= threshold_sinr:
            findings.append(('🟡','SINR', f'Посредственный SINR {sinr.mean():.1f} dB — ограничение модуляции'))
            penalty += 15
        else:
            findings.append(('🔴','SINR', f'Плохой SINR {sinr.mean():.1f} dB'))
            penalty += 25

    # Throughput
    if not df_ppp.empty:
        dl = df_ppp['dl_kbps'].dropna()
        zero_pct = round((dl == 0).sum()/len(dl)*100, 1)
        stats.update({'dl_mean_mbps': round(dl.mean()/1000, 2),
                      'dl_max_mbps':  round(dl.max()/1000, 2),
                      'dl_p95_mbps':  round(dl.quantile(0.95)/1000, 2),
                      'dl_zero_pct':  zero_pct})
        if dl.mean() >= 50000:
            findings.append(('✅','DL скорость', f'Высокая скорость {dl.mean()/1000:.1f} Mbps'))
        elif dl.mean() >= 10000:
            findings.append(('🟡','DL скорость', f'Средняя скорость {dl.mean()/1000:.1f} Mbps'))
            penalty += 10
        else:
            findings.append(('🔴','DL скорость', f'Низкая скорость {dl.mean()/1000:.1f} Mbps'))
            penalty += 25

        if zero_pct > 10:
            findings.append(('🔴','Обрывы DL', f'DL = 0 в {zero_pct}% измерений'))
            penalty += 20
            recs.append({'priority':'Высокий','category':'Стабильность',
                         'issue': f'DL = 0 kbps в {zero_pct}% выборок',
                         'action':'Проверить RLC retransmission, DRX-параметры, backhaul packet loss'})

    # Handovers
    if not df_cell.empty:
        pci_s = df_cell['pci'].dropna()
        ho = int((pci_s != pci_s.shift()).sum() - 1)
        stats['handover_count'] = ho
        if ho > 15:
            findings.append(('🔴','Хэндоверы', f'{ho} переключений соты — нестабильная привязка'))
            penalty += 20
            recs.append({'priority':'Высокий','category':'Мобильность',
                         'issue': f'{ho} хэндоверов',
                         'action':'Откалибровать TTT и A3-offset; проверить границы секторов'})
        elif ho > 5:
            findings.append(('🟡','Хэндоверы', f'{ho} переключений — умеренная мобильность'))
        else:
            findings.append(('✅','Хэндоверы', f'Только {ho} переключений — стабильная привязка'))

    if not recs:
        recs.append({'priority':'Низкий','category':'Общее',
                     'issue':'Критических проблем не выявлено',
                     'action':'Продолжать регулярный drive-test мониторинг'})

    score = max(0, 100 - penalty)
    grade = ('A','🟢','Хорошо') if score>=80 else \
            ('B','🟡','Удовлетворительно') if score>=60 else \
            ('C','🟠','Требует внимания') if score>=40 else ('D','🔴','Критично')

    return {'stats': stats, 'findings': findings, 'recs': recs,
            'score': score, 'grade': grade}


# ──────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ──────────────────────────────────────────────────────────────────────────────
def _bd():
    s = Side(border_style='thin', color='CCCCCC')
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_c): return PatternFill('solid', fgColor=hex_c)

def build_excel(df_cell, df_ppp, df_gps, df_devi, analysis, all_meta) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    stats = analysis['stats']

    C_DARK='1E3A5F'; C_MID='2E6DA4'; C_ACC='00A878'
    C_LIGHT='EAF4FB'; C_WARN='FFF3CD'; C_BAD='F8D7DA'; C_GOOD='D4EDDA'

    def hdr(ws, r, c, v, bg=None, fc='FFFFFF', bold=True, sz=10):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(bold=bold, color=fc, size=sz, name='Arial')
        cell.fill = _fill(bg or C_MID)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = _bd()
        return cell

    def dc(ws, r, c, v, fmt=None, bg=None):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name='Arial', size=9)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = _bd()
        if fmt: cell.number_format = fmt
        if bg:  cell.fill = _fill(bg)
        return cell

    # ── Sheet 1: Сводка ──────────────────────────────────────────────────────
    ws1 = wb.create_sheet('📋 Сводка')
    ws1.sheet_view.showGridLines = False
    ws1.merge_cells('A1:G1')
    t = ws1['A1']
    t.value = 'LTE Analyze for Asekapp — Отчёт Drive Test'
    t.font = Font(bold=True, size=15, color='FFFFFF', name='Arial')
    t.fill = _fill(C_DARK); t.alignment = Alignment(horizontal='center', vertical='center')
    ws1.row_dimensions[1].height = 34

    # Score
    score = analysis['score']
    g_icon, g_text = analysis['grade'][1], analysis['grade'][2]
    ws1.merge_cells('A3:B4')
    sc = ws1['A3']
    sc.value = f'{g_icon} {score}/100\n{g_text}'
    sc.font = Font(bold=True, size=13, name='Arial')
    sc.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    sc.fill = _fill(C_GOOD if score>=80 else (C_WARN if score>=60 else C_BAD))
    sc.border = _bd()
    ws1.row_dimensions[3].height = 22; ws1.row_dimensions[4].height = 22

    # KPI table
    ws1.merge_cells('A6:G6')
    kh = ws1['A6']
    kh.value = 'Ключевые KPI'
    kh.font = Font(bold=True, size=11, color='FFFFFF', name='Arial')
    kh.fill = _fill(C_MID); kh.alignment = Alignment(horizontal='center')

    for ci, h in enumerate(['KPI','Среднее','Мин','Макс','Оценка'], 1):
        hdr(ws1, 7, ci, h, bg=C_ACC); ws1.row_dimensions[7].height = 16

    def rsrp_g(v): return '✅ Хорошо' if v and v>=-90 else ('🟡 Норма' if v and v>=-100 else '🔴 Слабо')
    def rsrq_g(v): return '✅ Хорошо' if v and v>=-10 else ('🟡 Норма' if v and v>=-15 else '🔴 Плохо')
    def sinr_g(v): return '✅ Хорошо' if v and v>=10 else ('🟡 Норма' if v and v>=0 else '🔴 Плохо')
    def dl_g(v):   return '✅ Высокая' if v and v>=50 else ('🟡 Средняя' if v and v>=10 else '🔴 Низкая')

    kpi_rows = [
        ('RSRP (dBm)',       stats.get('rsrp_mean'), stats.get('rsrp_min'), stats.get('rsrp_max'), rsrp_g(stats.get('rsrp_mean'))),
        ('RSRQ (dB)',        stats.get('rsrq_mean'), stats.get('rsrq_min'), stats.get('rsrq_max'), rsrq_g(stats.get('rsrq_mean'))),
        ('SINR (dB)',        stats.get('sinr_mean'), stats.get('sinr_min'), stats.get('sinr_max'), sinr_g(stats.get('sinr_mean'))),
        ('DL средн. (Mbps)',stats.get('dl_mean_mbps'), None, stats.get('dl_max_mbps'), dl_g(stats.get('dl_mean_mbps'))),
        ('DL P95 (Mbps)',    stats.get('dl_p95_mbps'), None, None, ''),
        ('Нулевой DL (%)',   stats.get('dl_zero_pct'), None, None,
         '🔴 Проблема' if (stats.get('dl_zero_pct') or 0)>10 else '✅ ОК'),
        ('Хэндоверы',        stats.get('handover_count'), None, None,
         '🔴 Много' if (stats.get('handover_count') or 0)>15 else '✅ ОК'),
    ]
    for ri, row in enumerate(kpi_rows, 8):
        for ci, v in enumerate(row, 1):
            bg = None
            if ci == 2 and ri == 8:
                rm = stats.get('rsrp_mean')
                bg = C_GOOD if rm and rm>=-90 else (C_WARN if rm and rm>=-100 else C_BAD)
            dc(ws1, ri, ci, v, bg=bg)
        ws1.row_dimensions[ri].height = 15

    # Findings
    r = 8 + len(kpi_rows) + 2
    ws1.merge_cells(f'A{r}:G{r}')
    fh = ws1[f'A{r}']
    fh.value = '📊 Выводы по сети'
    fh.font = Font(bold=True, size=11, color='FFFFFF', name='Arial')
    fh.fill = _fill(C_MID); fh.alignment = Alignment(horizontal='center')
    ws1.row_dimensions[r].height = 18
    for icon, cat, text in analysis['findings']:
        r += 1
        ws1.cell(row=r, column=1, value=icon).border = _bd()
        c2 = ws1.cell(row=r, column=2, value=cat)
        c2.font = Font(bold=True, name='Arial', size=9); c2.border = _bd()
        ws1.merge_cells(f'C{r}:G{r}')
        c3 = ws1.cell(row=r, column=3, value=text)
        c3.font = Font(name='Arial', size=9)
        c3.alignment = Alignment(wrap_text=True); c3.border = _bd()
        ws1.row_dimensions[r].height = 15

    for col, w in zip('ABCDEFG', [5, 18, 40, 12, 12, 12, 18]):
        ws1.column_dimensions[col].width = w

    # ── Sheet 2: Рекомендации ─────────────────────────────────────────────────
    ws2 = wb.create_sheet('💡 Рекомендации')
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells('A1:E1')
    t2 = ws2['A1']
    t2.value = 'Рекомендации по оптимизации сети'
    t2.font = Font(bold=True, size=13, color='FFFFFF', name='Arial')
    t2.fill = _fill(C_DARK); t2.alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[1].height = 28

    for ci, h in enumerate(['#','Приоритет','Категория','Проблема','Действие'], 1):
        hdr(ws2, 2, ci, h, bg=C_ACC)
    ws2.row_dimensions[2].height = 16

    pri_c = {'Высокий': C_BAD, 'Средний': C_WARN, 'Низкий': C_GOOD, 'Информационно': C_LIGHT}
    for ri, rec in enumerate(analysis['recs'], 3):
        ws2.cell(row=ri, column=1, value=ri-2).border = _bd()
        pc = ws2.cell(row=ri, column=2, value=rec['priority'])
        pc.fill = _fill(pri_c.get(rec['priority'], 'FFFFFF'))
        pc.font = Font(bold=True, name='Arial', size=9)
        pc.alignment = Alignment(horizontal='center'); pc.border = _bd()
        for ci, k in enumerate(['category','issue','action'], 3):
            c = ws2.cell(row=ri, column=ci, value=rec[k])
            c.font = Font(name='Arial', size=9)
            c.alignment = Alignment(wrap_text=True, vertical='top'); c.border = _bd()
        ws2.row_dimensions[ri].height = 38
    for col, w in zip('ABCDE', [5, 13, 15, 48, 52]):
        ws2.column_dimensions[col].width = w

    # ── Sheet 3: CELLMEAS ─────────────────────────────────────────────────────
    ws3 = wb.create_sheet('📶 Signal')
    if not df_cell.empty:
        for ci, h in enumerate(['Время','EARFCN','PCI','RSRP (dBm)','RSRQ (dB)','SINR (dB)'], 1):
            hdr(ws3, 1, ci, h)
        cols3 = ['time','earfcn','pci','rsrp','rsrq','sinr']
        for ri, row in enumerate(df_cell[cols3].itertuples(index=False), 2):
            for ci, v in enumerate(row, 1):
                dc(ws3, ri, ci, str(v) if ci==1 else v)
        n = len(df_cell)
        ws3.conditional_formatting.add(f'D2:D{n+1}',
            ColorScaleRule(start_type='num', start_value=-100, start_color='F8D7DA',
                           mid_type='num',   mid_value=-90,    mid_color='FFF3CD',
                           end_type='num',   end_value=-80,    end_color='D4EDDA'))
        ws3.freeze_panes = 'A2'
        for col, w in zip('ABCDEF', [13,9,7,13,11,11]):
            ws3.column_dimensions[col].width = w

    # ── Sheet 4: Throughput ───────────────────────────────────────────────────
    ws4 = wb.create_sheet('📈 Throughput')
    if not df_ppp.empty:
        for ci, h in enumerate(['Время','DL (kbps)','UL (kbps)'], 1):
            hdr(ws4, 1, ci, h)
        cols4 = ['time','dl_kbps','ul_kbps']
        for ri, row in enumerate(df_ppp[cols4].itertuples(index=False), 2):
            for ci, v in enumerate(row, 1):
                dc(ws4, ri, ci, str(v) if ci==1 else v)
        n = len(df_ppp)
        ws4.conditional_formatting.add(f'B2:B{n+1}',
            DataBarRule(start_type='min', start_value=0,
                        end_type='max', end_value=None, color='2E6DA4'))
        ws4.freeze_panes = 'A2'
        for col, w in zip('ABC', [13,13,13]):
            ws4.column_dimensions[col].width = w

    # ── Sheet 5: GPS ──────────────────────────────────────────────────────────
    ws5 = wb.create_sheet('🗺️ GPS')
    if not df_gps.empty:
        for ci, h in enumerate(['Время','Долгота','Широта','GPS Fix','Спутников'], 1):
            hdr(ws5, 1, ci, h)
        cols5 = ['time','lon','lat','gps_fix','sats']
        existing = [c for c in cols5 if c in df_gps.columns]
        for ri, row in enumerate(df_gps[existing].itertuples(index=False), 2):
            for ci, v in enumerate(row, 1):
                dc(ws5, ri, ci, str(v) if ci==1 else v)
        ws5.freeze_panes = 'A2'
        for col, w in zip('ABCDE', [13,14,14,10,12]):
            ws5.column_dimensions[col].width = w

    # ── Sheet 6: Метаданные ───────────────────────────────────────────────────
    ws6 = wb.create_sheet('ℹ️ Метаданные')
    ws6.sheet_view.showGridLines = False
    ws6.merge_cells('A1:C1')
    mh = ws6['A1']
    mh.value = 'Метаданные сессии'
    mh.font = Font(bold=True, size=12, color='FFFFFF', name='Arial')
    mh.fill = _fill(C_DARK); mh.alignment = Alignment(horizontal='center')
    ws6.row_dimensions[1].height = 24

    for ri, (k, v) in enumerate(all_meta.items(), 2):
        ck = ws6.cell(row=ri, column=1, value=k)
        ck.font = Font(bold=True, name='Arial', size=10)
        ck.fill = _fill(C_LIGHT); ck.border = _bd()
        cv = ws6.cell(row=ri, column=2, value=v)
        cv.font = Font(name='Arial', size=10); cv.border = _bd()
        ws6.merge_cells(f'B{ri}:C{ri}')
        ws6.row_dimensions[ri].height = 15
    ws6.column_dimensions['A'].width = 20
    ws6.column_dimensions['B'].width = 40

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────────────────────
# GEMINI CALL
# ──────────────────────────────────────────────────────────────────────────────
# ──────────────────────────────────────────────────────────────────────────────
# AI API CALLS (Gemini + DeepSeek + Anthropic)
# ──────────────────────────────────────────────────────────────────────────────

def call_gemini(api_key: str, messages: list[dict], system_prompt: str) -> str:
    """Вызов Gemini API"""
    if not api_key:
        return '❌ Не указан API-ключ Gemini'
    
    url = f'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}'
    
    contents = []
    for m in messages:
        role = 'user' if m['role'] == 'user' else 'model'
        contents.append({'role': role, 'parts': [{'text': m['content']}]})
    
    payload = {
        'system_instruction': {'parts': [{'text': system_prompt}]},
        'contents': contents,
        'generationConfig': {'temperature': 0.7, 'maxOutputTokens': 1024},
    }
    
    try:
        r = requests.post(url, json=payload, timeout=30)
        r.raise_for_status()
        return r.json()['candidates'][0]['content']['parts'][0]['text']
    except Exception as e:
        return f'❌ Gemini ошибка: {str(e)[:100]}'


def call_deepseek(api_key: str, messages: list[dict], system_prompt: str) -> str:
    """Вызов DeepSeek API"""
    if not api_key:
        return '❌ Не указан API-ключ DeepSeek'
    
    url = "https://api.deepseek.com/v1/chat/completions"
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    formatted_messages = [{"role": "system", "content": system_prompt}]
    for msg in messages:
        formatted_messages.append({
            "role": msg["role"],
            "content": msg["content"]
        })
    
    payload = {
        "model": "deepseek-chat",
        "messages": formatted_messages,
        "temperature": 0.7,
        "max_tokens": 1024
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=30)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f'❌ DeepSeek ошибка: {str(e)[:100]}'


def call_anthropic(api_key: str, messages: list[dict], system_prompt: str) -> str:
    """Вызов Anthropic Claude API (исправленная версия)"""
    if not api_key:
        return '❌ Не указан API-ключ Anthropic'
    
    url = "https://api.anthropic.com/v1/messages"
    
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    
    # Преобразуем историю в формат Anthropic (только user и assistant)
    anthropic_messages = []
    for m in messages:
        # Пропускаем system сообщения — они идут отдельно
        if m['role'] == 'system':
            continue
        # Конвертируем 'assistant' в ассистента, остальное в user
        role = 'assistant' if m['role'] == 'assistant' else 'user'
        anthropic_messages.append({
            "role": role,
            "content": m['content']
        })
    
    # Если сообщений нет, добавляем хотя бы одно user message
    if not anthropic_messages:
        anthropic_messages = [{"role": "user", "content": "Привет"}]
    
    # Убеждаемся, что последнее сообщение от user (требование Claude API)
    if anthropic_messages and anthropic_messages[-1]['role'] != 'user':
        anthropic_messages.append({"role": "user", "content": "Продолжим?"})
    
    # Формируем запрос
    payload = {
        "model": "claude-3-sonnet-20240229",  # Правильное имя модели
        "max_tokens": 1024,
        "temperature": 0.7,
        "system": system_prompt,  # system — отдельный параметр!
        "messages": anthropic_messages
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=30)
        response.raise_for_status()
        return response.json()["content"][0]["text"]
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 400:
            # Покажем подробности ошибки для диагностики
            error_detail = e.response.text
            return f'❌ Ошибка 400: Неверный запрос к Anthropic.\nДетали: {error_detail[:300]}'
        elif e.response.status_code == 401:
            return '❌ Неверный API-ключ Anthropic. Проверьте ключ в настройках.'
        elif e.response.status_code == 429:
            return '❌ Превышен лимит запросов к Anthropic. Попробуйте позже.'
        else:
            return f'❌ Anthropic ошибка: {e.response.status_code} — {e.response.text[:200]}'
    except Exception as e:
        return f'❌ Ошибка: {str(e)[:100]}'


# ──────────────────────────────────────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title='LTE Analyze for Asekapp', page_icon='📡', layout='wide')

st.markdown("""
<div style="background:#1E293B;padding:20px;border-radius:10px;margin-bottom:20px;border-left:8px solid #0EA5E9">
  <h1 style="color:white;margin:0;padding-bottom:4px;">
    📡 LTE Analyze for Asekapp
    <span style="font-size:13px;color:#94A3B8;font-weight:normal;font-family:monospace;margin-left:10px;">
      Mega AI Edition 2026 · AsekDzhientaev inc.
    </span>
  </h1>
  <p style="color:#94A3B8;margin:0;font-size:15px;">
    Экспертная аналитическая среда · Сетевой анализ · Excel-отчёт · ИИ-помощник Илгиз Искендерович
  </p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
st.sidebar.markdown('### 🎛️ Критерии KPI')
threshold_rsrp = st.sidebar.slider('Критический RSRP (дБм)', -125, -90, -110, 1)
threshold_sinr = st.sidebar.slider('Критический SINR (дБ)',   -10,  10,    0, 1)
worst_cell_limit = st.sidebar.slider('Мин. плохих отсчётов для worst cell', 5, 100, 15)

st.sidebar.markdown('---')
st.sidebar.markdown('### 🤖 AI-ассистенты')
st.sidebar.markdown('Введите API-ключи для активации:')

gemini_key = st.sidebar.text_input('🔑 Gemini API Key', type='password',
                                   placeholder='AIza...',
                                   help='aistudio.google.com/apikey')

deepseek_key = st.sidebar.text_input('🔑 DeepSeek API Key', type='password',
                                     placeholder='sk-...',
                                     help='platform.deepseek.com/api_keys')

anthropic_key = st.sidebar.text_input('🔑 Anthropic Claude API Key', type='password',
                                      placeholder='sk-ant-...',
                                      help='console.anthropic.com')
st.sidebar.markdown('---')
st.sidebar.markdown('### 📂 Загрузка логов')
uploaded_files = st.sidebar.file_uploader('Перетащите .nmf файлы:',
                                          type=['nmf'], accept_multiple_files=True)

if not uploaded_files:
    st.info('💡 Загрузите файлы `.nmf` в панели слева.')
    st.stop()

# ── Parse all files ───────────────────────────────────────────────────────────
all_gps, all_cell, all_ppp, all_devi = [], [], [], []
combined_meta = {}
for f in uploaded_files:
    text = f.read().decode('utf-8', errors='replace')
    res = parse_nmf_lines(text.splitlines())
    combined_meta.update(res['meta'])
    combined_meta['filename'] = f.name
    if not res['gps'].empty:   all_gps.append(res['gps'])
    if not res['cell'].empty:  all_cell.append(res['cell'])
    if not res['ppp'].empty:   all_ppp.append(res['ppp'])
    if not res['devi'].empty:  all_devi.append(res['devi'])

df_gps  = pd.concat(all_gps,  ignore_index=True) if all_gps  else pd.DataFrame()
df_cell = pd.concat(all_cell, ignore_index=True) if all_cell else pd.DataFrame()
df_ppp  = pd.concat(all_ppp,  ignore_index=True) if all_ppp  else pd.DataFrame()
df_devi = pd.concat(all_devi, ignore_index=True) if all_devi else pd.DataFrame()

if df_cell.empty:
    st.error('Не найдено CELLMEAS LTE данных.'); st.stop()

# ── EARFCN filter ─────────────────────────────────────────────────────────────
available_earfcns = sorted(df_cell['earfcn'].dropna().unique().astype(int).tolist())
st.sidebar.markdown('---')
selected_earfcn = st.sidebar.selectbox('EARFCN:', ['Все частоты'] + available_earfcns)
if selected_earfcn != 'Все частоты':
    df_cell = df_cell[df_cell['earfcn'] == selected_earfcn].reset_index(drop=True)

# ── Analysis ──────────────────────────────────────────────────────────────────
analysis = analyze_network(df_cell, df_ppp, df_devi, threshold_rsrp, threshold_sinr)
stats    = analysis['stats']

# ── KPI cards ─────────────────────────────────────────────────────────────────
avg_rsrp = stats.get('rsrp_mean', 0)
avg_sinr = stats.get('sinr_mean', 0)
dl_mean  = stats.get('dl_mean_mbps', 0)
dl_max   = stats.get('dl_max_mbps',  0)

m1, m2, m3, m4 = st.columns(4)
m1.markdown(f"<div style='background:#F8FAFC;padding:15px;border-radius:8px;border-bottom:4px solid #6366F1;text-align:center'><span style='color:#64748B;font-size:13px;font-weight:bold'>СРЕДНИЙ RSRP</span><h2 style='margin:5px 0;color:#1E293B'>{avg_rsrp} дБм</h2></div>", unsafe_allow_html=True)
m2.markdown(f"<div style='background:#F8FAFC;padding:15px;border-radius:8px;border-bottom:4px solid #10B981;text-align:center'><span style='color:#64748B;font-size:13px;font-weight:bold'>СРЕДНИЙ SINR</span><h2 style='margin:5px 0;color:#1E293B'>{avg_sinr} дБ</h2></div>", unsafe_allow_html=True)
m3.markdown(f"<div style='background:#F8FAFC;padding:15px;border-radius:8px;border-bottom:4px solid #F59E0B;text-align:center'><span style='color:#64748B;font-size:13px;font-weight:bold'>СРЕДНИЙ DL</span><h2 style='margin:5px 0;color:#1E293B'>{dl_mean} Mbps</h2></div>", unsafe_allow_html=True)
m4.markdown(f"<div style='background:#F8FAFC;padding:15px;border-radius:8px;border-bottom:4px solid #EF4444;text-align:center'><span style='color:#64748B;font-size:13px;font-weight:bold'>ПИКОВЫЙ DL</span><h2 style='margin:5px 0;color:#1E293B'>{dl_max} Mbps</h2></div>", unsafe_allow_html=True)

# ── Worst cells ───────────────────────────────────────────────────────────────
df_cell['is_bad_sinr'] = df_cell['sinr'] < threshold_sinr
worst_cells = df_cell.groupby(['earfcn','pci']).agg(
    total_samples=('sinr','count'), bad_sinr_samples=('is_bad_sinr','sum'),
    avg_rsrp=('rsrp','mean'), avg_sinr=('sinr','mean')).reset_index()
worst_cells_df = worst_cells[worst_cells['bad_sinr_samples'] >= worst_cell_limit]\
                    .sort_values('bad_sinr_samples', ascending=False)

# ── Tabs ──────────────────────────────────────────────────────────────────────
t1, t2, t3, t4, t5 = st.tabs([
    '🌐 Анализ сети', '📶 Таймлайн KPI', '🗺️ Карта покрытия',
    '🗄️ Данные', '📥 Excel-отчёт'
])

# ══ Tab 1: Анализ сети ════════════════════════════════════════════════════════
with t1:
    score = analysis['score']
    g_icon, g_text = analysis['grade'][1], analysis['grade'][2]
    score_color = '#2E7D32' if score>=80 else ('#F57C00' if score>=60 else '#C62828')

    col_sc, col_info = st.columns([1, 3])
    with col_sc:
        st.markdown(f"""
        <div style="text-align:center;padding:1.5rem;background:#f8f9fa;
                    border-radius:12px;border:2px solid {score_color}">
          <div style="font-size:2.5rem">{g_icon}</div>
          <div style="font-size:2.2rem;font-weight:700;color:{score_color}">{score}/100</div>
          <div style="color:#666">{g_text}</div>
          <div style="font-size:0.8rem;color:#999;margin-top:4px">Оценка сети</div>
        </div>""", unsafe_allow_html=True)

    with col_info:
        k1,k2,k3,k4 = st.columns(4)
        k1.metric('RSRP', f"{stats.get('rsrp_mean','—')} dBm", f"мин {stats.get('rsrp_min','—')}")
        k2.metric('RSRQ', f"{stats.get('rsrq_mean','—')} dB")
        k3.metric('SINR', f"{stats.get('sinr_mean','—')} dB", f"макс {stats.get('sinr_max','—')}")
        k4.metric('DL', f"{stats.get('dl_mean_mbps','—')} Mbps", f"пик {stats.get('dl_max_mbps','—')}")
        k5,k6,k7,k8 = st.columns(4)
        k5.metric('Нулевой DL', f"{stats.get('dl_zero_pct','—')}%", delta_color='inverse')
        k6.metric('Хэндоверы', stats.get('handover_count','—'))
        k7.metric('RSRP хор.', f"{stats.get('rsrp_good_pct','—')}%")
        k8.metric('SINR хор.', f"{stats.get('sinr_good_pct','—')}%")

    st.subheader('📊 Выводы')
    for icon, cat, text in analysis['findings']:
        ci, cc, ct = st.columns([0.4,2,8])
        ci.write(icon); cc.markdown(f'**{cat}**'); ct.write(text)

    st.subheader('💡 Рекомендации')
    pri_icons = {'Высокий':'🔴','Средний':'🟡','Низкий':'🟢','Информационно':'🔵'}
    for rec in analysis['recs']:
        with st.expander(f"{pri_icons.get(rec['priority'],'•')} [{rec['priority']}] {rec['category']}"):
            st.markdown(f"**Проблема:** {rec['issue']}")
            st.markdown(f"**Действие:** {rec['action']}")

    if not worst_cells_df.empty:
        st.subheader('🚨 Worst Cells')
        st.dataframe(worst_cells_df.round(2), use_container_width=True)

# ══ Tab 2: Таймлайн ══════════════════════════════════════════════════════════
with t2:
    fig = make_subplots(rows=3, cols=1, shared_xaxes=True,
                        subplot_titles=('RSRP (дБм)','SINR (дБ)','DL (Mbps)'),
                        vertical_spacing=0.08)
    fig.add_trace(go.Scatter(x=df_cell['time'], y=df_cell['rsrp'],
                             name='RSRP', line=dict(color='#6366F1', width=1.5)), row=1, col=1)
    fig.add_hline(y=-90, line_dash='dot', line_color='gray', row=1, col=1)
    fig.add_trace(go.Scatter(x=df_cell['time'], y=df_cell['sinr'],
                             name='SINR', line=dict(color='#10B981', width=1.5)), row=2, col=1)
    fig.add_hline(y=10, line_dash='dot', line_color='gray', row=2, col=1)
    if not df_ppp.empty:
        fig.add_trace(go.Scatter(x=df_ppp['time'], y=df_ppp['dl_kbps']/1000,
                                 name='DL', line=dict(color='#F59E0B', width=1.5)), row=3, col=1)
    fig.update_layout(height=520, template='plotly_white',
                      showlegend=False, hovermode='x unified')
    st.plotly_chart(fig, use_container_width=True)

# ══ Tab 3: Карта ══════════════════════════════════════════════════════════════
with t3:
    if not df_gps.empty:
        kpi_choice = st.selectbox('KPI для раскраски', ['RSRP','RSRQ','SINR','DL kbps'])
        kpi_src = {'RSRP':(df_cell,'rsrp'),'RSRQ':(df_cell,'rsrq'),
                   'SINR':(df_cell,'sinr'),'DL kbps':(df_ppp,'dl_kbps')}
        src_df, src_col = kpi_src[kpi_choice]
        map_gps = df_gps.dropna(subset=['lat','lon']).reset_index(drop=True)
        if not src_df.empty and src_col in src_df.columns:
            idx = np.linspace(0, len(src_df)-1, num=len(map_gps), dtype=int)
            map_gps[src_col] = src_df[src_col].iloc[idx].values
        fig_m = px.scatter_mapbox(map_gps, lat='lat', lon='lon', color=src_col,
                                  color_continuous_scale=['#EF4444','#F59E0B','#10B981'],
                                  zoom=14, mapbox_style='open-street-map', height=500)
        fig_m.update_layout(margin=dict(l=0,r=0,t=0,b=0))
        st.plotly_chart(fig_m, use_container_width=True)
    else:
        st.warning('GPS-данные не найдены.')

# ══ Tab 4: Данные ══════════════════════════════════════════════════════════════
with t4:
    ds = st.selectbox('Набор данных', ['CELLMEAS','PPPRATE','GPS','DEVI'])
    df_show = {'CELLMEAS':df_cell,'PPPRATE':df_ppp,'GPS':df_gps,'DEVI':df_devi}[ds]
    if df_show.empty:
        st.info('Нет данных.')
    else:
        st.dataframe(df_show, use_container_width=True, height=400)
        st.download_button(f'⬇️ {ds}.csv',
                           data=df_show.to_csv(index=False).encode('utf-8'),
                           file_name=f'{ds.lower()}.csv', mime='text/csv')

# ══ Tab 5: Excel ══════════════════════════════════════════════════════════════
with t5:
    st.subheader('📥 Excel-отчёт')
    st.markdown("""
    **6 листов:** 📋 Сводка · 💡 Рекомендации · 📶 Signal · 📈 Throughput · 🗺️ GPS · ℹ️ Метаданные
    """)
    if not XLSX_OK:
        st.error('Установите openpyxl: `pip install openpyxl`')
    elif st.button('🔄 Сформировать отчёт', type='primary'):
        with st.spinner('Формирую…'):
            xlsx = build_excel(df_cell, df_ppp, df_gps, df_devi, analysis, combined_meta)
        fname = f"LTE_Report_{combined_meta.get('start_date','').replace('.','')}.xlsx"
        st.download_button('⬇️ Скачать .xlsx', data=xlsx, file_name=fname,
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        st.success(f'Готово: {fname}')

# ──────────────────────────────────────────────────────────────────────────────
# 🤖 ИИ-АССИСТЕНТ ИЛГИЗ ИСКЕНДЕРОВИЧ (Gemini)
# ──────────────────────────────────────────────────────────────────────────────
# ──────────────────────────────────────────────────────────────────────────────
# 🤖 ТРОЙНОЙ ИИ-АССИСТЕНТ (Gemini + DeepSeek + Claude)
# ──────────────────────────────────────────────────────────────────────────────
st.divider()
st.subheader('🤖 AI-ассистенты: Илгиз Искендерович (Трио)')

# Статистика для всех AI
worst_summary = ''
if 'worst_cells_df' in dir() and not worst_cells_df.empty:
    worst_summary = 'Проблемные секторы: ' + ', '.join(
        [f"EARFCN {r['earfcn']}/PCI {r['pci']} ({r['bad_sinr_samples']:.0f} плохих точек)"
         for _, r in worst_cells_df.head(3).iterrows()])

system_prompt = (
    f"Тебя зовут Илгиз Искендерович. Ты — эксперт по LTE-сетям (инженер L3).\n\n"
    f"Данные лога:\n"
    f"• RSRP: {avg_rsrp} дБм (мин {stats.get('rsrp_min', '?')}, макс {stats.get('rsrp_max', '?')})\n"
    f"• RSRQ: {stats.get('rsrq_mean', '?')} dB\n"
    f"• SINR: {avg_sinr} дБ (макс {stats.get('sinr_max', '?')})\n"
    f"• DL скорость: {dl_mean:.1f} Мбит/с (пик {dl_max:.1f})\n"
    f"• Нулевой DL: {stats.get('dl_zero_pct', '?')}%\n"
    f"• Хэндоверы: {stats.get('handover_count', '?')}\n"
    f"• Оценка сети: {analysis['score']}/100\n"
    f"{worst_summary}\n\n"
    f"Отвечай кратко, по делу, с инженерной аргументацией, на русском языке."
)

# Выбор режима
ai_mode = st.radio(
    "🎮 Режим работы:",
    ["💬 Один ассистент", "🎭 Сравнить ответы (все три)"],
    horizontal=True
)

if ai_mode == "💬 Один ассистент":
    selected_ai = st.selectbox(
        "Выберите ассистента:",
        ["🧠 DeepSeek (рекомендую)", "🤖 Gemini", "🎨 Claude (Anthropic)"]
    )
    
    ai_map = {
        "🧠 DeepSeek (рекомендую)": ("deepseek", deepseek_key, call_deepseek),
        "🤖 Gemini": ("gemini", gemini_key, call_gemini),
        "🎨 Claude (Anthropic)": ("anthropic", anthropic_key, call_anthropic)
    }
    
    ai_name, ai_key, ai_func = ai_map[selected_ai]
    
    if not ai_key:
        st.warning(f"⚠️ Введите API-ключ для {selected_ai} в боковой панели")
    else:
        if 'messages_single' not in st.session_state:
            st.session_state.messages_single = [{'role': 'assistant',
                'content': f'Здравствуйте! Я {selected_ai} в роли Илгиза Искендеровича. Задавайте вопросы по LTE!'}]
        
        for msg in st.session_state.messages_single:
            with st.chat_message(msg['role']):
                st.markdown(msg['content'])
        
        if user_query := st.chat_input('Задайте вопрос...'):
            st.session_state.messages_single.append({'role': 'user', 'content': user_query})
            with st.chat_message('user'):
                st.markdown(user_query)
            
            with st.chat_message('assistant'):
                with st.spinner(f'{selected_ai} анализирует...'):
                    response = ai_func(ai_key, st.session_state.messages_single, system_prompt)
                st.markdown(response)
            st.session_state.messages_single.append({'role': 'assistant', 'content': response})

else:
    st.markdown("💡 **Задайте вопрос — получите ответы от всех трёх AI одновременно!**")
    
    if 'messages_compare' not in st.session_state:
        st.session_state.messages_compare = []
    
    for msg in st.session_state.messages_compare:
        with st.chat_message(msg['role']):
            st.markdown(msg['content'])
    
    if user_query := st.chat_input('Задайте вопрос всем трём ассистентам...'):
        st.session_state.messages_compare.append({'role': 'user', 'content': user_query})
        with st.chat_message('user'):
            st.markdown(user_query)
        
        col_g, col_d, col_c = st.columns(3)
        
        with col_g:
            st.markdown("**🤖 Gemini**")
            if gemini_key:
                with st.spinner("Gemini думает..."):
                    resp_g = call_gemini(gemini_key, st.session_state.messages_compare, system_prompt)
                st.markdown(resp_g)
            else:
                st.warning("Нет ключа Gemini")
        
        with col_d:
            st.markdown("**🧠 DeepSeek**")
            if deepseek_key:
                with st.spinner("DeepSeek думает..."):
                    resp_d = call_deepseek(deepseek_key, st.session_state.messages_compare, system_prompt)
                st.markdown(resp_d)
            else:
                st.warning("Нет ключа DeepSeek")
        
        with col_c:
            st.markdown("**🎨 Claude**")
            if anthropic_key:
                with st.spinner("Claude думает..."):
                    resp_c = call_anthropic(anthropic_key, st.session_state.messages_compare, system_prompt)
                st.markdown(resp_c)
            else:
                st.warning("Нет ключа Anthropic")
        
        combined_response = f"**Gemini:** {resp_g[:200]}...\n\n**DeepSeek:** {resp_d[:200]}...\n\n**Claude:** {resp_c[:200]}..."
        st.session_state.messages_compare.append({'role': 'assistant', 'content': combined_response})