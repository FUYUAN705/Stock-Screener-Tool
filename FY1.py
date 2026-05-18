#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
黄牛选股器 v2.0 - 黄牛剑客完整版
融合35条炒股原则 + 黄牛形态 + 均线偏离度选股
包含：七步选股 + 实时监控提醒 + 交易记录 + 自动热点排序 + 行业缓存 + 我的盈利模式
数据源：baostock + 东方财富（备用）
修复版本：修复数据源获取问题
"""
import sys
import os
import baostock as bs
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
import threading
import queue
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import time
import warnings
import hashlib
import pickle
from concurrent.futures import ThreadPoolExecutor, as_completed
import numpy as np
import re

# 尝试导入akshare作为备用数据源
try:
    import akshare as ak
    AKSHARE_AVAILABLE = True
    print("✓ akshare已加载")
except ImportError:
    AKSHARE_AVAILABLE = False
    print("⚠ akshare未安装，东方财富数据源不可用。安装命令: pip install akshare")

warnings.filterwarnings('ignore')

# ==============================
# 全局配置
# ==============================
bs_lock = threading.Lock()
logged_in = False
stop_scan_flag = False
CONFIG_FILE = os.path.join(os.path.expanduser("~"), "黄牛选股器配置.json")
CACHE_DIR = os.path.join(os.path.expanduser("~"), "黄牛选股器缓存")
RECORD_DIR = os.path.join(os.path.expanduser("~"), "黄牛选股器交易记录")
INDUSTRY_CACHE_FILE = os.path.join(CACHE_DIR, "industry_cache.json")
INDUSTRY_CACHE_EXPIRE = 24 * 60 * 60

for d in [CACHE_DIR, RECORD_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

CACHE_EXPIRE_SECONDS = 24 * 60 * 60

_industry_stocks_cache = None
_industry_stocks_cache_time = 0

# 数据源选择
DATA_SOURCE = "auto"  # auto, baostock, eastmoney

# ==============================
# 缓存工具函数
# ==============================
def get_cache_key(code, start_date, end_date):
    key_str = f"{code}_{start_date}_{end_date}"
    return hashlib.md5(key_str.encode()).hexdigest()

def get_cached_data(code, start_date, end_date):
    cache_key = get_cache_key(code, start_date, end_date)
    cache_file = os.path.join(CACHE_DIR, f"{cache_key}.pkl")
    if not os.path.exists(cache_file):
        return None
    try:
        with open(cache_file, 'rb') as f:
            cached = pickle.load(f)
        if isinstance(cached, tuple) and len(cached) == 2:
            cached_time, cached_df = cached
        else:
            cached_df = cached
            cached_time = os.path.getmtime(cache_file)
        if time.time() - cached_time > CACHE_EXPIRE_SECONDS:
            os.remove(cache_file)
            return None
        return cached_df
    except:
        return None

def save_to_cache(code, start_date, end_date, df):
    if df is None or df.empty:
        return
    cache_key = get_cache_key(code, start_date, end_date)
    cache_file = os.path.join(CACHE_DIR, f"{cache_key}.pkl")
    try:
        with open(cache_file, 'wb') as f:
            pickle.dump((time.time(), df), f)
    except:
        pass

def clear_cache():
    count = 0
    for f in os.listdir(CACHE_DIR):
        if f.endswith('.pkl'):
            os.remove(os.path.join(CACHE_DIR, f))
            count += 1
    return count

# ==============================
# 登录登出
# ==============================
def ensure_login():
    global logged_in
    if logged_in:
        return True
    with bs_lock:
        if logged_in:
            return True
        try:
            lg = bs.login()
            if lg.error_code == '0':
                logged_in = True
                print(f"baostock登录成功")
                return True
            else:
                print(f"baostock登录失败: {lg.error_msg}")
                return False
        except Exception as e:
            print(f"baostock登录异常: {e}")
            return False

def bs_logout():
    global logged_in
    with bs_lock:
        if logged_in:
            bs.logout()
            logged_in = False

# ==============================
# 修复后的数据获取函数（去掉市值相关）
# ==============================

def fetch_stock_data_baostock_with_retry(code, start_date, end_date, retry=2):
    """修复后的baostock数据获取 - 去掉市值管理"""
    for attempt in range(retry):
        try:
            if not ensure_login():
                if attempt == 0:
                    time.sleep(1)
                    continue
                return None
            
            with bs_lock:
                # 只使用确认存在的字段
                rs = bs.query_history_k_data_plus(
                    code, 
                    "date,code,open,close,high,low,volume,turn,pctChg",
                    start_date=start_date, 
                    end_date=end_date, 
                    frequency="d", 
                    adjustflag="2"  # 2表示前复权
                )
                
                if rs.error_code != '0':
                    if attempt == retry - 1:
                        print(f"baostock查询失败 {code}: {rs.error_msg}")
                    return None
                
                data_list = []
                while rs.next():
                    data_list.append(rs.get_row_data())
            
            if not data_list:
                return None
            
            df = pd.DataFrame(data_list, columns=rs.fields)
            
            # 转换数据类型
            numeric_cols = ['open', 'close', 'high', 'low', 'turn', 'pctChg', 'volume']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # 删除空值行
            df = df.dropna(subset=['close', 'high', 'low', 'open'])
            
            if df.empty:
                return None
            
            # 添加空字段以保持接口一致
            df['pbMRQ'] = None
            df['peTTM'] = None
            
            # 去掉市值计算，直接设为None
            df['market_cap_yi'] = None
            
            return df
            
        except Exception as e:
            print(f"baostock获取失败 {code} (尝试 {attempt+1}/{retry}): {e}")
            if attempt == retry - 1:
                return None
            time.sleep(2)
    
    return None


def fetch_stock_data_eastmoney(code, start_date, end_date):
    """
    修复后的东方财富数据获取函数 - 去掉市值管理
    """
    if not AKSHARE_AVAILABLE:
        return None
    
    try:
        # 转换代码格式：sh.600000 -> 600000
        symbol = code.replace('sh.', '').replace('sz.', '')
        
        # 转换日期格式
        start = start_date.replace('-', '')
        end = end_date.replace('-', '')
        
        # 尝试多个akshare接口
        df = None
        
        # 主方法：stock_zh_a_hist
        try:
            df = ak.stock_zh_a_hist(symbol=symbol, period='daily',
                                     start_date=start, end_date=end,
                                     adjust='qfq')
            if df is not None and not df.empty:
                pass  # 成功获取
        except Exception as e:
            print(f"stock_zh_a_hist失败 {symbol}: {e}")
        
        # 备用方法：如果失败，尝试获取最近数据
        if df is None or df.empty:
            try:
                # 尝试不指定结束日期，获取最近数据
                end_date_obj = datetime.now()
                start_date_obj = end_date_obj - timedelta(days=400)
                start = start_date_obj.strftime('%Y%m%d')
                end = end_date_obj.strftime('%Y%m%d')
                
                df = ak.stock_zh_a_hist(symbol=symbol, period='daily',
                                         start_date=start, end_date=end,
                                         adjust='qfq')
            except Exception as e:
                print(f"备用方法也失败 {symbol}: {e}")
                return None
        
        if df is None or df.empty:
            return None
        
        # 统一列名格式（适配不同的akshare版本）
        column_mapping = {
            '日期': 'date',
            '开盘': 'open',
            '收盘': 'close',
            '最高': 'high',
            '最低': 'low',
            '成交量': 'volume',
            '成交额': 'amount',
            '振幅': 'amplitude',
            '涨跌幅': 'pctChg',
            '涨跌额': 'change',
            '换手率': 'turn'
        }
        
        # 重命名列
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df.rename(columns={old_name: new_name}, inplace=True)
        
        # 添加代码列
        df['code'] = code
        
        # 确保必要列存在
        if 'pctChg' not in df.columns:
            if 'close' in df.columns:
                df['pctChg'] = df['close'].pct_change() * 100
            else:
                df['pctChg'] = 0
        df['pctChg'] = df['pctChg'].fillna(0)
        
        if 'turn' not in df.columns:
            df['turn'] = 0
        
        # 添加其他必要字段，市值设为None
        df['pbMRQ'] = None
        df['peTTM'] = None
        df['market_cap_yi'] = None  # 去掉市值计算
        
        # 转换数据类型
        numeric_cols = ['open', 'close', 'high', 'low', 'volume', 'pctChg']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # 删除无效数据
        df = df.dropna(subset=['close'])
        
        if df.empty:
            return None
        
        # 按日期排序
        if 'date' in df.columns:
            df['date'] = pd.to_datetime(df['date'])
            df = df.sort_values('date')
        
        return df
        
    except Exception as e:
        print(f"东方财富获取失败 {code}: {e}")
        return None


def fetch_stock_data_with_fallback(code, start_date, end_date, use_cache=True, force_source=None):
    """
    带备用数据源的股票数据获取
    优先使用缓存 -> 根据选择的数据源获取数据
    """
    # 1. 尝试从缓存读取
    if use_cache:
        cached = get_cached_data(code, start_date, end_date)
        if cached is not None and len(cached) > 0:
            return cached
    
    source = force_source if force_source else DATA_SOURCE
    
    df = None
    
    # 2. 根据数据源选择获取数据
    if source == "baostock":
        df = fetch_stock_data_baostock_with_retry(code, start_date, end_date)
    elif source == "eastmoney":
        if AKSHARE_AVAILABLE:
            df = fetch_stock_data_eastmoney(code, start_date, end_date)
        else:
            print(f"东方财富数据源不可用，请安装akshare: pip install akshare")
    else:  # auto模式：先尝试baostock，失败则用东方财富
        df = fetch_stock_data_baostock_with_retry(code, start_date, end_date)
        if df is None or df.empty:
            if AKSHARE_AVAILABLE:
                df = fetch_stock_data_eastmoney(code, start_date, end_date)
    
    # 3. 保存到缓存
    if df is not None and not df.empty:
        save_to_cache(code, start_date, end_date, df)
    
    return df


def check_data_source_health():
    """检查数据源可用性"""
    test_codes = ['sh.600000', 'sz.000001']
    results = {'baostock': False, 'eastmoney': False}
    
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.now() - timedelta(days=5)).strftime('%Y-%m-%d')
    
    # 测试baostock
    try:
        if ensure_login():
            df = fetch_stock_data_baostock_with_retry(test_codes[0], start_date, end_date)
            results['baostock'] = df is not None and len(df) > 0
            if results['baostock']:
                print(f"baostock测试成功，获取{len(df)}条数据")
            else:
                print("baostock测试失败")
    except Exception as e:
        print(f"baostock测试异常: {e}")
        results['baostock'] = False
    
    # 测试东方财富
    if AKSHARE_AVAILABLE:
        try:
            df = fetch_stock_data_eastmoney(test_codes[0], start_date, end_date)
            results['eastmoney'] = df is not None and len(df) > 0
            if results['eastmoney']:
                print(f"东方财富测试成功，获取{len(df)}条数据")
            else:
                print("东方财富测试失败")
        except Exception as e:
            print(f"东方财富测试异常: {e}")
            results['eastmoney'] = False
    else:
        print("akshare未安装，跳过东方财富测试")
    
    return results

def fetch_stock_data(code, start_date, end_date):
    """兼容原接口"""
    return fetch_stock_data_with_fallback(code, start_date, end_date)

# ==============================
# 行业数据获取（修复版）
# ==============================
def get_industry_stocks(force_refresh=False):
    """修复后的行业数据获取"""
    global _industry_stocks_cache, _industry_stocks_cache_time
    
    if not force_refresh and _industry_stocks_cache is not None:
        if time.time() - _industry_stocks_cache_time < 3600:
            return _industry_stocks_cache
    
    if not force_refresh and os.path.exists(INDUSTRY_CACHE_FILE):
        try:
            with open(INDUSTRY_CACHE_FILE, 'r', encoding='utf-8') as f:
                cache = json.load(f)
            if time.time() - cache.get('timestamp', 0) < INDUSTRY_CACHE_EXPIRE:
                df = pd.DataFrame(cache['data'])
                _industry_stocks_cache = df
                _industry_stocks_cache_time = time.time()
                return df
        except:
            pass
    
    # 尝试从baostock获取
    try:
        if not ensure_login():
            print("无法登录baostock获取行业数据")
            return pd.DataFrame()
        
        with bs_lock:
            rs = bs.query_stock_industry()
            if rs.error_code != '0':
                print(f"查询行业数据失败: {rs.error_msg}")
                return pd.DataFrame()

        data = []
        while rs.next():
            row = rs.get_row_data()
            if row and len(row) >= 4:
                # 过滤掉ST股票和退市股票
                stock_name = row[2]
                if 'ST' not in stock_name and '退' not in stock_name and '暂停' not in stock_name:
                    data.append({
                        'code': row[1], 
                        'name': stock_name, 
                        'industry': row[3]
                    })
        
        if not data:
            print("未获取到行业数据")
            return pd.DataFrame()
            
        df = pd.DataFrame(data)
        
        # 保存缓存
        try:
            with open(INDUSTRY_CACHE_FILE, 'w', encoding='utf-8') as f:
                json.dump({'timestamp': time.time(), 'data': df.to_dict('records')}, f, ensure_ascii=False)
        except Exception as e:
            print(f"保存行业缓存失败: {e}")
        
        _industry_stocks_cache = df
        _industry_stocks_cache_time = time.time()
        
        print(f"成功获取{len(df)}只股票的行业信息")
        return df
        
    except Exception as e:
        print(f"获取行业数据异常: {e}")
        
        # 如果baostock失败，尝试从缓存读取旧数据
        if os.path.exists(INDUSTRY_CACHE_FILE):
            try:
                with open(INDUSTRY_CACHE_FILE, 'r', encoding='utf-8') as f:
                    cache = json.load(f)
                df = pd.DataFrame(cache['data'])
                _industry_stocks_cache = df
                _industry_stocks_cache_time = time.time()
                print(f"从缓存加载行业数据: {len(df)}条")
                return df
            except:
                pass
        
        return pd.DataFrame()

def safe_round(value, digits=2):
    """安全地四舍五入，处理None和NaN"""
    if value is None or pd.isna(value):
        return 0.0
    try:
        return round(float(value), digits)
    except:
        return 0.0

# ==============================
# 核心技术指标（保持不变）
# ==============================
def calc_area_above_below(df, ma_col, start_idx, end_idx):
    if end_idx - start_idx < 2:
        return 0.0
    area = 0.0
    for i in range(start_idx, end_idx):
        price = df.iloc[i]['close']
        ma = df.iloc[i][ma_col]
        height = price - ma
        area += height
    return area

def calc_huangniu_fade(df, ma_period=30, fade_ratio=0.75):
    """重构后的黄牛衰竭形态"""
    if len(df) < ma_period + 20:
        return 0, 0, "数据不足"

    ma_col = f'MA{ma_period}'
    if ma_col not in df.columns:
        df[ma_col] = df['close'].rolling(ma_period, min_periods=1).mean()

    crosses = []
    i = ma_period
    while i < len(df) - 5:
        prev = df.iloc[i-1]
        curr = df.iloc[i]
        prev_ma = prev[ma_col]
        curr_ma = curr[ma_col]
        prev_close = prev['close']
        curr_close = curr['close']

        if prev_close <= prev_ma and curr_close > curr_ma:
            j = i + 1
            while j < len(df) and df.iloc[j]['close'] > df.iloc[j][ma_col]:
                j += 1
            up_area = calc_area_above_below(df, ma_col, i, j)
            strength = abs(curr_close - prev_ma) / (prev_ma + 1e-5)
            crosses.append((i, 'UP', up_area, strength, j))
            i = j
            continue

        elif prev_close >= prev_ma and curr_close < curr_ma:
            j = i + 1
            while j < len(df) and df.iloc[j]['close'] < df.iloc[j][ma_col]:
                j += 1
            down_area = abs(calc_area_above_below(df, ma_col, i, j))
            strength = abs(curr_close - prev_ma) / (prev_ma + 1e-5)
            crosses.append((i, 'DOWN', down_area, strength, j))
            i = j
            continue
        i += 1

    if len(crosses) < 3:
        return 0, 0, "交互次数不足"

    recent_crosses = crosses[-5:]

    up_areas = [c[2] for c in recent_crosses if c[1] == 'UP']
    down_areas = [c[2] for c in recent_crosses if c[1] == 'DOWN']
    up_strengths = [c[3] for c in recent_crosses if c[1] == 'UP']

    if not up_areas or not down_areas:
        return 0, 0, "缺少多空波段"

    down_shrink = False
    if len(down_areas) >= 2:
        if down_areas[-1] < down_areas[-2] * 0.85:
            down_shrink = True

    multi_force_gt_empty = False
    if up_areas and down_areas:
        last_up_area = up_areas[-1]
        last_down_area = down_areas[-1]
        if last_up_area > last_down_area * 1.1:
            multi_force_gt_empty = True

    force_fade = 1.0
    if len(up_strengths) >= 2:
        force_fade = up_strengths[-1] / (up_strengths[-2] + 1e-5)

    reverse_fade = False
    if len(down_areas) >= 2 and down_areas[-1] < down_areas[-2] * fade_ratio:
        reverse_fade = True

    if down_shrink and multi_force_gt_empty and force_fade > 0.9:
        return -1, round(force_fade, 3), "多方穿空+下降面积缩小+多方力度占优（买入）"

    if len(up_areas) >= 2 and force_fade <= fade_ratio and reverse_fade:
        return 1, round(force_fade, 3), "多方反弹力度递减+反向水漂（卖出）"

    return 0, 0, "无明确衰竭信号"

def check_my_profit_pattern(df, debug=False):
    """
    用户自定义盈利模式：前2天小跌 + 前1天小涨（反弹修复）
    只做形态识别
    返回: (是否符合, T-2跌幅, T-1涨幅, 修复比例, 描述)
    """
    if len(df) < 10:
        return False, 0, 0, 0, "数据不足"
    
    try:
        prev1 = df.iloc[-2]  # T-1（前一天）
        prev2 = df.iloc[-3]  # T-2（前两天）
        
        pct_prev1 = prev1.get('pctChg', 0)
        pct_prev2 = prev2.get('pctChg', 0)
        
        # 1. 前2天小跌（-0.5% 到 -7%）
        if not (-7.0 <= pct_prev2 <= -0.3):
            return False, pct_prev2, pct_prev1, 0, f"T-2跌幅{pct_prev2:.2f}%"
        
        # 2. 前1天小涨（+0.1% 到 +6%）
        if not (0.1 <= pct_prev1 <= 6.0):
            return False, pct_prev2, pct_prev1, 0, f"T-1涨幅{pct_prev1:.2f}%"
        
        # 计算修复比例：T-1涨幅 / |T-2跌幅|
        repair_ratio = pct_prev1 / (abs(pct_prev2) + 0.01)
        
        desc = f"T-2:{pct_prev2:.2f}% | T-1:{pct_prev1:.2f}% | 修复:{repair_ratio:.2f}"
        
        return True, pct_prev2, pct_prev1, repair_ratio, desc
        
    except Exception as e:
        return False, 0, 0, 0, f"错误"

def check_limit_up_quality(df, days_back=30):
    """检查涨停质量：返回(是否有涨停, 涨停次数, 最近涨停天数, 封板力度)"""
    if len(df) < days_back:
        return False, 0, 999, 0
    limit_count = 0
    best_strength = 0
    last_limit_days_ago = 999
    for i in range(-min(days_back, len(df)), 0):
        pct = df.iloc[i].get('pctChg', 0)
        if pd.isna(pct):
            continue
        if pct >= 9.5:
            limit_count += 1
            days_ago = abs(i)
            if days_ago < last_limit_days_ago:
                last_limit_days_ago = days_ago
            high = df.iloc[i].get('high', 0)
            close = df.iloc[i].get('close', 0)
            if high and high > 0:
                strength = 100 - abs((close - high) / high * 100)
                best_strength = max(best_strength, strength)
    return limit_count >= 1, limit_count, last_limit_days_ago, best_strength

def check_volume_stability(df, days_back=10, max_ratio=3.0):
    """检查量能稳定性"""
    if len(df) < days_back:
        return True, 1.0, 0
    df['VOL_MA5'] = df['volume'].rolling(5, min_periods=1).mean()
    abnormal = 0
    ratios = []
    for i in range(-min(days_back, len(df)), 0):
        vol = df.iloc[i]['volume']
        ma5 = df.iloc[i]['VOL_MA5']
        if pd.isna(ma5) or ma5 == 0:
            continue
        ratio = vol / ma5
        ratios.append(ratio)
        if ratio > max_ratio or ratio < 0.2:
            abnormal += 1
    current_ratio = ratios[-1] if ratios else 1.0
    return abnormal <= 5, current_ratio, abnormal

def calc_indicators(df, ma_periods):
    for period in ma_periods:
        if len(df) >= period:
            df[f'MA{period}'] = df['close'].rolling(period, min_periods=1).mean()
            df[f'BIAS{period}'] = (df['close'] - df[f'MA{period}']) / df[f'MA{period}'] * 100
    df['TURN_MA5'] = df['turn'].rolling(5, min_periods=1).mean()
   
    huangniu_signal, fade_degree, desc = calc_huangniu_fade(df)
    df['HUANGNIU_SIGNAL'] = huangniu_signal
    df['HUANGNIU_FADE_DEGREE'] = fade_degree
    df['HUANGNIU_DESC'] = desc
   
    has_limit, limit_count, limit_days_ago, limit_strength = check_limit_up_quality(df, 30)
    df['HAS_LIMIT'] = has_limit
    df['LIMIT_COUNT'] = limit_count
    df['LIMIT_DAYS_AGO'] = limit_days_ago
    df['LIMIT_STRENGTH'] = limit_strength
   
    vol_stable, vol_ratio, vol_abnormal = check_volume_stability(df)
    df['VOL_STABLE'] = vol_stable
    df['VOL_RATIO'] = vol_ratio
    df['VOL_ABNORMAL'] = vol_abnormal
   
    return df

def get_buy_points(prev_day, current_day):
    """获取右侧买点（基于前一天的K线）"""
    try:
        high = prev_day['high']
        low = prev_day['low']
        open_price = prev_day['open']
        close = prev_day['close']
        if pd.isna(high) or pd.isna(low):
            return None, None, None
        # 右侧买点(上)：前一天的最高价
        # 右侧买点(下)：前一天的最低价
        upper = round(float(high), 2)
        lower = round(float(low), 2)
        is_red = close > open_price
        return upper, lower, is_red
    except:
        return None, None, None

def get_industries_list():
    df = get_industry_stocks()
    return sorted(df['industry'].unique()) if not df.empty else []

# ==============================
# 交易记录管理（保持不变）
# ==============================
class TradeRecorder:
    def __init__(self):
        self.today = datetime.now().strftime('%Y-%m-%d')
        self.record_file = os.path.join(RECORD_DIR, f"交易记录_{self.today}.json")
        self.trades = self.load()
   
    def load(self):
        if os.path.exists(self.record_file):
            try:
                with open(self.record_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {"trades": [], "statistics": {"success_rate": 0, "sell_accuracy": 0, "daily_pnl": 0}}
        return {"trades": [], "statistics": {"success_rate": 0, "sell_accuracy": 0, "daily_pnl": 0}}
   
    def save(self):
        with open(self.record_file, 'w', encoding='utf-8') as f:
            json.dump(self.trades, f, ensure_ascii=False, indent=2)
   
    def add_trade(self, trade):
        self.trades["trades"].append(trade)
        self.update_statistics()
        self.save()
   
    def update_statistics(self):
        trades = self.trades["trades"]
        if not trades:
            return
        wins = [t for t in trades if t.get('profit_pct', 0) > 0]
        self.trades["statistics"]["success_rate"] = len(wins) / len(trades) * 100 if trades else 0
        good_sells = [t for t in trades if t.get('sell_position', 0) >= 90]
        self.trades["statistics"]["sell_accuracy"] = len(good_sells) / len(trades) * 100 if trades else 0
        self.trades["statistics"]["daily_pnl"] = sum(t.get('profit_pct', 0) for t in trades)
   
    def get_stats_text(self):
        s = self.trades["statistics"]
        return f"今日统计: 选股成功率: {s['success_rate']:.1f}% | 卖点准确率: {s['sell_accuracy']:.1f}% | 累计盈亏: {s['daily_pnl']:.2f}%"

# ==============================
# 监控提醒模块（简化版，保持不变）
# ==============================
class MonitorPanel:
    def __init__(self, parent, log_func):
        self.parent = parent
        self.log = log_func
        self.watch_stocks = []
        self.positions = {}
        self.recorder = TradeRecorder()
        self.running = True
        self.setup_ui()
        self.start_monitor()
    
    def setup_ui(self):
        self.frame = tk.LabelFrame(self.parent, text="实时监控与提醒",
                                   font=("微软雅黑", 12, "bold"), bg='#f0f0f0', fg="#2c3e50", padx=10, pady=10)
        self.frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)
        
        self.frame.grid_rowconfigure(0, weight=0)
        self.frame.grid_rowconfigure(1, weight=1)
        self.frame.grid_rowconfigure(2, weight=0)
        self.frame.grid_rowconfigure(3, weight=1)
        self.frame.grid_rowconfigure(4, weight=0)
        self.frame.grid_rowconfigure(5, weight=2)
        self.frame.grid_rowconfigure(6, weight=0)
        self.frame.grid_rowconfigure(7, weight=0)
        self.frame.grid_rowconfigure(8, weight=0)
        self.frame.grid_columnconfigure(0, weight=1)
        
        # 监控股票池
        watch_frame = tk.LabelFrame(self.frame, text="监控股票池", font=("微软雅黑", 10, "bold"), bg='#f0f0f0')
        watch_frame.grid(row=0, column=0, sticky="nsew", pady=2)
        
        cols = ("code", "name", "buy_upper", "buy_lower", "t2_pct", "t1_pct", "status")
        self.watch_tree = ttk.Treeview(watch_frame, columns=cols, show="headings", height=6)
        self.watch_tree.heading("code", text="代码")
        self.watch_tree.heading("name", text="名称")
        self.watch_tree.heading("buy_upper", text="买点(上)")
        self.watch_tree.heading("buy_lower", text="买点(下)")
        self.watch_tree.heading("t2_pct", text="T-2跌幅")
        self.watch_tree.heading("t1_pct", text="T-1涨幅")
        self.watch_tree.heading("status", text="状态")
        
        self.watch_tree.column("code", width=90, anchor="center")
        self.watch_tree.column("name", width=100, anchor="center")
        self.watch_tree.column("buy_upper", width=80, anchor="center")
        self.watch_tree.column("buy_lower", width=80, anchor="center")
        self.watch_tree.column("t2_pct", width=80, anchor="center")
        self.watch_tree.column("t1_pct", width=80, anchor="center")
        self.watch_tree.column("status", width=80, anchor="center")
        
        self.watch_tree.pack(fill=tk.BOTH, expand=True, pady=2)
        
        btn_frame = tk.Frame(watch_frame, bg='#f0f0f0')
        btn_frame.pack(fill=tk.X, pady=2)
        tk.Button(btn_frame, text="导入选股结果", command=self.import_selected, bg="#4F81BD", fg="white", font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="清空监控池", command=self.clear_watch, bg="#e74c3c", fg="white", font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=2)
        
        # 当前持仓
        pos_frame = tk.LabelFrame(self.frame, text="当前持仓", font=("微软雅黑", 10, "bold"), bg='#f0f0f0')
        pos_frame.grid(row=2, column=0, sticky="nsew", pady=2)
        
        pos_cols = ("code", "name", "buy_price", "current", "profit", "buy_date", "action")
        self.position_tree = ttk.Treeview(pos_frame, columns=pos_cols, show="headings", height=3)
        self.position_tree.heading("code", text="代码")
        self.position_tree.heading("name", text="名称")
        self.position_tree.heading("buy_price", text="买入价")
        self.position_tree.heading("current", text="当前价")
        self.position_tree.heading("profit", text="盈亏%")
        self.position_tree.heading("buy_date", text="买入日期")
        self.position_tree.heading("action", text="操作")
        for col in pos_cols:
            self.position_tree.column(col, width=85)
        self.position_tree.pack(fill=tk.BOTH, expand=True, pady=2)
        
        action_frame = tk.Frame(pos_frame, bg='#f0f0f0')
        action_frame.pack(fill=tk.X, pady=2)
        tk.Button(action_frame, text="记录买入", command=self.record_buy, bg="#27ae60", fg="white", font=("微软雅黑", 9), width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(action_frame, text="记录卖出", command=self.record_sell, bg="#e74c3c", fg="white", font=("微软雅黑", 9), width=10).pack(side=tk.LEFT, padx=2)
        
        # 实时信号提醒
        signal_frame = tk.LabelFrame(self.frame, text="实时信号提醒", font=("微软雅黑", 10, "bold"), bg='#f0f0f0')
        signal_frame.grid(row=4, column=0, sticky="nsew", pady=2)
        
        self.signal_text = tk.Text(signal_frame, font=("Consolas", 9), height=6, bg='#1a1a2e', fg='#00ff88')
        self.signal_text.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)
        
        time_frame = tk.Frame(self.frame, bg='#2c3e50', height=32)
        time_frame.grid(row=5, column=0, sticky="ew", pady=2)
        time_frame.grid_propagate(False)
        
        self.time_label = tk.Label(time_frame, text="当前时段: 等待开盘", 
                                   font=("微软雅黑", 10, "bold"), 
                                   bg='#2c3e50', fg='#FFD700')
        self.time_label.pack(expand=True)
        
        stats_frame = tk.Frame(self.frame, bg='#f0f0f0', height=28)
        stats_frame.grid(row=6, column=0, sticky="ew", pady=2)
        stats_frame.grid_propagate(False)
        self.stats_label = tk.Label(stats_frame, text=self.recorder.get_stats_text(), font=("微软雅黑", 9), bg='#f0f0f0')
        self.stats_label.pack(expand=True)
        
        record_frame = tk.Frame(self.frame, bg='#f0f0f0', height=32)
        record_frame.grid(row=7, column=0, sticky="ew", pady=2)
        record_frame.grid_propagate(False)
        tk.Button(record_frame, text="查看今日交易记录", command=self.show_records, bg="#3498db", fg="white", font=("微软雅黑", 9), width=16).pack(side=tk.LEFT, padx=2)
        tk.Button(record_frame, text="记录每日总结", command=self.record_daily_summary, bg="#9b59b6", fg="white", font=("微软雅黑", 9), width=14).pack(side=tk.LEFT, padx=2)
        tk.Button(record_frame, text="查看历史记录", command=self.show_history, bg="#f39c12", fg="white", font=("微软雅黑", 9), width=14).pack(side=tk.LEFT, padx=2)
    
    def import_selected(self):
        """导入选股结果到监控池"""
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        files = [f for f in os.listdir(desktop) if f.startswith("选股结果_") and f.endswith(".xlsx")]
        if not files:
            self.log("没有找到选股结果文件，请先运行选股")
            return
        latest = max(files, key=lambda x: os.path.getctime(os.path.join(desktop, x)))
        filepath = os.path.join(desktop, latest)
        try:
            df = pd.read_excel(filepath)
            self.log(f"找到选股结果文件: {latest}")
        except Exception as e:
            self.log(f"读取文件失败: {e}")
            return
        
        self.watch_stocks = []
        for _, row in df.iterrows():
            try:
                buy_upper = row.get('右侧买点(上)', '-')
                buy_lower = row.get('右侧买点(下)', '-')
                t2_pct = row.get('T-2跌幅%', '-')
                t1_pct = row.get('T-1涨幅%', '-')
                
                self.watch_stocks.append({
                    'code': row['代码'], 
                    'name': row['名称'],
                    'buy_upper': buy_upper if buy_upper != '-' else '-',
                    'buy_lower': buy_lower if buy_lower != '-' else '-',
                    't2_pct': t2_pct,
                    't1_pct': t1_pct,
                    'current_price': float(row['当前价格']) if pd.notna(row['当前价格']) else 0
                })
            except Exception as e:
                self.log(f"解析行数据出错: {e}")
                continue
        
        self.refresh_watch_list()
        self.log(f"已导入 {len(self.watch_stocks)} 只股票到监控池")
    
    def clear_watch(self):
        self.watch_stocks = []
        self.refresh_watch_list()
        self.log("监控池已清空")
    
    def refresh_watch_list(self):
        for item in self.watch_tree.get_children():
            self.watch_tree.delete(item)
        for s in self.watch_stocks:
            try:
                status = "观察中"
                self.watch_tree.insert("", tk.END, values=(
                    s['code'], s['name'],
                    s['buy_upper'], s['buy_lower'],
                    s['t2_pct'], s['t1_pct'], status
                ))
            except:
                pass
    
    def record_buy(self):
        selected = self.watch_tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请先在监控池中选择要买入的股票")
            return
        
        item = self.watch_tree.item(selected[0])
        values = item['values']
        if len(values) < 2:
            return
        code, name = values[0], values[1]
        
        dialog = tk.Toplevel(self.parent)
        dialog.title(f"记录买入 - {name}")
        dialog.geometry("300x280")
        
        tk.Label(dialog, text=f"股票: {name}({code})", font=("微软雅黑", 10, "bold")).pack(pady=10)
        
        tk.Label(dialog, text="买入价格:").pack()
        price_entry = tk.Entry(dialog, width=15)
        price_entry.pack(pady=5)
        
        for s in self.watch_stocks:
            if s['code'] == code:
                price_entry.insert(0, str(s['current_price']))
                break
        
        tk.Label(dialog, text="仓位(%):").pack()
        size_entry = tk.Entry(dialog, width=15)
        size_entry.pack(pady=5)
        size_entry.insert(0, "10")
        
        def confirm():
            try:
                price = float(price_entry.get())
                size = float(size_entry.get())
                stop_loss = round(price * 0.97, 2)
                self.positions[code] = {
                    'name': name, 'buy_price': price, 'buy_date': datetime.now().strftime('%Y-%m-%d %H:%M'),
                    'size_pct': size, 'stop_loss': stop_loss
                }
                self.refresh_positions()
                self.add_signal(f"买入: {name}({code}) 价格:{price} 仓位:{size}% 止损:{stop_loss}")
                dialog.destroy()
            except:
                messagebox.showerror("错误", "请输入有效数字")
        
        tk.Button(dialog, text="确认", command=confirm, bg="#27ae60", fg="white").pack(pady=20)
    
    def record_sell(self):
        if not self.positions:
            messagebox.showwarning("提示", "当前没有持仓")
            return
        
        dialog = tk.Toplevel(self.parent)
        dialog.title("记录卖出")
        dialog.geometry("300x250")
        
        tk.Label(dialog, text="选择股票:", font=("微软雅黑", 10)).pack(pady=5)
        code_var = tk.StringVar()
        code_combo = ttk.Combobox(dialog, textvariable=code_var, values=list(self.positions.keys()), width=15)
        code_combo.pack(pady=5)
        
        tk.Label(dialog, text="卖出价格:").pack()
        price_entry = tk.Entry(dialog, width=15)
        price_entry.pack(pady=5)
        
        tk.Label(dialog, text="卖出原因:").pack()
        reason_combo = ttk.Combobox(dialog, values=["止盈", "止损-3%", "3日未启动", "其他"], width=15)
        reason_combo.pack(pady=5)
        reason_combo.set("止盈")
        
        def confirm():
            code = code_var.get()
            if code not in self.positions:
                return
            try:
                price = float(price_entry.get())
                reason = reason_combo.get()
                self.sell_position(code, price, reason)
                dialog.destroy()
            except:
                messagebox.showerror("错误", "请输入有效价格")
        
        tk.Button(dialog, text="确认", command=confirm, bg="#e74c3c", fg="white").pack(pady=20)
    
    def add_position(self, code, name, price, size_pct):
        stop_loss = round(price * 0.97, 2)
        self.positions[code] = {
            'name': name, 'buy_price': price, 'buy_date': datetime.now().strftime('%Y-%m-%d %H:%M'),
            'size_pct': size_pct, 'stop_loss': stop_loss
        }
        self.refresh_positions()
        self.add_signal(f"买入信号: {name}({code}) 价格:{price} 仓位:{size_pct}% 止损:{stop_loss}")
    
    def sell_position(self, code, sell_price, reason):
        if code in self.positions:
            pos = self.positions[code]
            profit = (sell_price - pos['buy_price']) / pos['buy_price'] * 100
            self.recorder.add_trade({
                'code': code, 'name': pos['name'], 'buy_price': pos['buy_price'], 'sell_price': sell_price,
                'profit_pct': profit, 'reason': reason, 'sell_time': datetime.now().strftime('%Y-%m-%d %H:%M')
            })
            self.add_signal(f"卖出: {pos['name']}({code}) 盈亏:{profit:.2f}% 原因:{reason}")
            del self.positions[code]
            self.refresh_positions()
            self.stats_label.config(text=self.recorder.get_stats_text())
    
    def refresh_positions(self):
        for item in self.position_tree.get_children():
            self.position_tree.delete(item)
        for code, pos in self.positions.items():
            current = pos['buy_price']
            profit = 0
            self.position_tree.insert("", tk.END, values=(code, pos['name'], pos['buy_price'], 
                                                          current, f"{profit:.2f}", 
                                                          pos['buy_date'], "卖出"))
    
    def add_signal(self, msg):
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.signal_text.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.signal_text.see(tk.END)
        self.log(msg)
    
    def show_records(self):
        records = self.recorder.trades.get("trades", [])
        if not records:
            messagebox.showinfo("交易记录", "暂无交易记录")
            return
        text = "\n".join([f"{r['sell_time']} {r['name']} 盈亏:{r['profit_pct']:.2f}% {r['reason']}" for r in records])
        messagebox.showinfo("今日交易记录", text)
    
    def show_history(self):
        files = [f for f in os.listdir(RECORD_DIR) if f.startswith("交易记录_") and f.endswith(".json")]
        if not files:
            messagebox.showinfo("历史记录", "暂无历史记录")
            return
        
        dialog = tk.Toplevel(self.parent)
        dialog.title("历史交易记录")
        dialog.geometry("600x400")
        
        text_area = tk.Text(dialog, font=("Consolas", 10))
        text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        for f in sorted(files, reverse=True):
            filepath = os.path.join(RECORD_DIR, f)
            with open(filepath, 'r', encoding='utf-8') as file:
                data = json.load(file)
                text_area.insert(tk.END, f"\n{'='*50}\n")
                text_area.insert(tk.END, f"日期: {f.replace('交易记录_', '').replace('.json', '')}\n")
                text_area.insert(tk.END, f"成功率: {data['statistics']['success_rate']:.1f}%\n")
                text_area.insert(tk.END, f"总盈亏: {data['statistics']['daily_pnl']:.2f}%\n")
                for t in data['trades']:
                    text_area.insert(tk.END, f"  {t['name']}: {t['profit_pct']:.2f}% ({t['reason']})\n")
    
    def record_daily_summary(self):
        def save():
            summary = {
                'date': datetime.now().strftime('%Y-%m-%d'),
                'follow_rules': follow_var.get(),
                'stop_loss_executed': stop_var.get(),
                'mistake': mistake_text.get("1.0", tk.END).strip(),
                'improvement': improve_text.get("1.0", tk.END).strip()
            }
            filepath = os.path.join(RECORD_DIR, f"每日总结_{summary['date']}.json")
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(summary, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存成功", "今日总结已保存")
            dialog.destroy()
        
        dialog = tk.Toplevel(self.parent)
        dialog.title("每日交易总结")
        dialog.geometry("500x450")
        
        follow_var = tk.BooleanVar(value=True)
        stop_var = tk.BooleanVar(value=True)
        
        tk.Label(dialog, text="是否遵守选股规则？", font=("微软雅黑", 10)).pack(anchor="w", padx=10, pady=5)
        tk.Radiobutton(dialog, text="是", variable=follow_var, value=True).pack(anchor="w", padx=30)
        tk.Radiobutton(dialog, text="否", variable=follow_var, value=False).pack(anchor="w", padx=30)
        
        tk.Label(dialog, text="是否在止损位果断离场？", font=("微软雅黑", 10)).pack(anchor="w", padx=10, pady=5)
        tk.Radiobutton(dialog, text="是", variable=stop_var, value=True).pack(anchor="w", padx=30)
        tk.Radiobutton(dialog, text="否", variable=stop_var, value=False).pack(anchor="w", padx=30)
        
        tk.Label(dialog, text="今日失误总结：", font=("微软雅黑", 10)).pack(anchor="w", padx=10, pady=5)
        mistake_text = tk.Text(dialog, height=4, width=55)
        mistake_text.pack(padx=10, pady=5)
        
        tk.Label(dialog, text="明日改进计划：", font=("微软雅黑", 10)).pack(anchor="w", padx=10, pady=5)
        improve_text = tk.Text(dialog, height=4, width=55)
        improve_text.pack(padx=10, pady=5)
        
        tk.Button(dialog, text="保存", command=save, bg="#27ae60", fg="white", font=("微软雅黑", 10)).pack(pady=10)
    
    def update_time_reminder(self):
        now = datetime.now()
        hour = now.hour
        minute = now.minute
        time_str = f"{hour:02d}:{minute:02d}"
        
        if 9 <= hour < 15:
            if hour == 9 and minute < 15:
                self.time_label.config(text=f"集合竞价 ({time_str})", fg="#FFD700")
            elif 9 <= hour < 10:
                self.time_label.config(text=f"黄金半小时 ({time_str})", fg="#90EE90")
            elif 10 <= hour < 14:
                self.time_label.config(text=f"持仓观望 ({time_str})", fg="#87CEEB")
            elif 14 <= hour < 15:
                self.time_label.config(text=f"尾盘决策 ({time_str})", fg="#FFA500")
            else:
                self.time_label.config(text=f"休市 ({time_str})", fg="#CCCCCC")
        else:
            self.time_label.config(text=f"休市 ({time_str})", fg="#CCCCCC")
        
        self.parent.after(60000, self.update_time_reminder)
    
    def start_monitor(self):
        self.update_time_reminder()
    
    def stop(self):
        self.running = False

# ==============================
# 扫描主函数（去掉市值过滤）
# ==============================
def scan_stocks(selected_boards, selected_industries, **params):
    global stop_scan_flag
   
   
# ==============================
# GUI主界面（去掉市值相关控件）
# ==============================
class StockPickerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("黄牛选股器 v2.0 无市值限制版")
        self.root.geometry("1400x900")
        self.root.configure(bg='#f0f0f0')
        
        self.q = queue.Queue()
        self.all_industries = []
        self.monitor = None
        self.scan_thread = None
        self.hot_ranking = []
        self.industry_hot_score = {}
        
        self.setup_ui()
        self.load_industries()
        self.load_config()
        self.update_ui()
        
        # 启动时测试数据源
        self.test_data_sources()
    
    def test_data_sources(self):
        """测试数据源并在日志中显示结果"""
        def test():
            self.log("正在测试数据源...")
            results = check_data_source_health()
            self.log(f"数据源状态 - baostock: {'✓ 可用' if results['baostock'] else '✗ 不可用'}, "
                    f"东方财富: {'✓ 可用' if results['eastmoney'] else '✗ 不可用'}")
            if not results['baostock'] and not results['eastmoney']:
                self.log("⚠ 警告: 所有数据源均不可用，请检查网络连接")
            elif not results['baostock'] and results['eastmoney']:
                self.log("ℹ 提示: baostock不可用，将使用东方财富数据源")
            elif results['baostock'] and not results['eastmoney']:
                self.log("ℹ 提示: 东方财富不可用，将使用baostock数据源")
        
        threading.Thread(target=test, daemon=True).start()
    
    def setup_ui(self):
        # 顶部标题栏
        title = tk.Frame(self.root, bg="#2c3e50", height=70)
        title.pack(fill=tk.X)
        title.pack_propagate(False)
        tk.Label(title, text="黄牛选股器 v2.0 (无市值限制)", 
                 font=("微软雅黑", 20, "bold"), fg="white", bg="#2c3e50").pack(expand=True)
        
        # 控制栏
        control = tk.Frame(self.root, bg='#f0f0f0', height=45)
        control.pack(fill=tk.X, pady=(10, 5))
        control.pack_propagate(False)
        
        self.progress = ttk.Progressbar(control, length=600, mode="determinate")
        self.progress.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.progress_label = tk.Label(control, text="0%", font=("微软雅黑", 10, "bold"), bg='#f0f0f0', width=6)
        self.progress_label.pack(side=tk.LEFT, padx=5)
        
        self.start_btn = tk.Button(control, text="开始选股", command=self.start_scan, 
                                   bg="#27ae60", fg="white", font=("微软雅黑", 10, "bold"), width=12)
        self.start_btn.pack(side=tk.RIGHT, padx=5)
        self.cancel_btn = tk.Button(control, text="取消", command=self.cancel_scan, 
                                    bg="#e74c3c", fg="white", font=("微软雅黑", 10), width=8, state="disabled")
        self.cancel_btn.pack(side=tk.RIGHT, padx=5)
        tk.Button(control, text="重置", command=self.reset_config, 
                  bg="#95a5a6", fg="white", width=8).pack(side=tk.RIGHT, padx=5)
        tk.Button(control, text="清空缓存", command=self.clear_cache, 
                  bg="#e67e22", fg="white", width=10).pack(side=tk.RIGHT, padx=5)
        tk.Button(control, text="退出", command=self.exit_app, 
                  bg="#7f8c8d", fg="white", width=8).pack(side=tk.RIGHT, padx=5)
        
        # 主内容：左右分屏
        paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # ========== 左侧选股面板 ==========
        left = tk.Frame(paned, bg='#f0f0f0')
        paned.add(left, weight=2)
        left_main = tk.Frame(left, bg='#f0f0f0')
        left_main.pack(fill=tk.BOTH, expand=True)
        
        # 1. 板块选择
        board_frame = tk.LabelFrame(left_main, text="1. 选择板块", font=("微软雅黑", 11, "bold"), 
                                    bg='#f0f0f0', padx=10, pady=5)
        board_frame.pack(fill=tk.X, pady=5, padx=10)
        board_inner = tk.Frame(board_frame, bg='#f0f0f0')
        board_inner.pack()
        self.var600 = tk.BooleanVar(value=True)
        self.var300 = tk.BooleanVar(value=True)
        self.var000 = tk.BooleanVar(value=True)
        self.var688 = tk.BooleanVar(value=False)
        for text, var in [("主板(60)", self.var600), ("创业板(30)", self.var300), 
                          ("中小板(00)", self.var000), ("科创板(688)", self.var688)]:
            tk.Checkbutton(board_inner, text=text, variable=var, bg='#f0f0f0', 
                          font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=15)
        
        # 2. 行业选择
        industry_frame = tk.LabelFrame(left_main, text="2. 选择行业", 
                                       font=("微软雅黑", 11, "bold"), bg='#f0f0f0', padx=10, pady=5)
        industry_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)
        
        hot_toolbar = tk.Frame(industry_frame, bg='#f0f0f0')
        hot_toolbar.pack(fill=tk.X, pady=(0, 5))
        self.refresh_hot_btn = tk.Button(hot_toolbar, text="刷新热点排序", command=self.refresh_hot_ranking,
                                          bg="#e67e22", fg="white", font=("微软雅黑", 9))
        self.refresh_hot_btn.pack(side=tk.LEFT, padx=2)
        self.hot_status_label = tk.Label(hot_toolbar, text="正在加载行业...", font=("微软雅黑", 8), 
                                          bg='#f0f0f0', fg="#27ae60")
        self.hot_status_label.pack(side=tk.LEFT, padx=10)
        
        # 数据源选择
        source_frame = tk.Frame(industry_frame, bg='#f0f0f0')
        source_frame.pack(fill=tk.X, pady=(0, 5))
        tk.Label(source_frame, text="数据源:", font=("微软雅黑", 9), bg='#f0f0f0').pack(side=tk.LEFT, padx=5)
        self.data_source_var = tk.StringVar(value="auto")
        tk.Radiobutton(source_frame, text="自动", variable=self.data_source_var, 
                       value="auto", bg='#f0f0f0', command=self.change_data_source).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(source_frame, text="baostock", variable=self.data_source_var, 
                       value="baostock", bg='#f0f0f0', command=self.change_data_source).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(source_frame, text="东方财富", variable=self.data_source_var, 
                       value="eastmoney", bg='#f0f0f0', command=self.change_data_source).pack(side=tk.LEFT, padx=5)
        
        # Treeview行业列表
        columns = ("select", "rank", "industry", "score", "limit", "turn", "up")
        self.industry_tree = ttk.Treeview(industry_frame, columns=columns, show="headings", height=12, selectmode="extended")
        
        self.industry_tree.heading("select", text="☑")
        self.industry_tree.heading("rank", text="排名")
        self.industry_tree.heading("industry", text="行业")
        self.industry_tree.heading("score", text="热度")
        self.industry_tree.heading("limit", text="涨停")
        self.industry_tree.heading("turn", text="换手%")
        self.industry_tree.heading("up", text="上涨%")
        
        self.industry_tree.column("select", width=40, anchor="center")
        self.industry_tree.column("rank", width=50, anchor="center")
        self.industry_tree.column("industry", width=120, anchor="w")
        self.industry_tree.column("score", width=60, anchor="center")
        self.industry_tree.column("limit", width=50, anchor="center")
        self.industry_tree.column("turn", width=60, anchor="center")
        self.industry_tree.column("up", width=60, anchor="center")
        
        tree_scrollbar = tk.Scrollbar(industry_frame, orient="vertical", command=self.industry_tree.yview)
        self.industry_tree.configure(yscrollcommand=tree_scrollbar.set)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.industry_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        select_btn_frame = tk.Frame(industry_frame, bg='#f0f0f0')
        select_btn_frame.pack(fill=tk.X, pady=5)
        tk.Button(select_btn_frame, text="全选", command=self.select_all_industries_tree,
                  bg="#3498db", fg="white", width=8).pack(side=tk.LEFT, padx=2)
        tk.Button(select_btn_frame, text="清空", command=self.clear_industries_tree,
                  bg="#95a5a6", fg="white", width=8).pack(side=tk.LEFT, padx=2)
        tk.Button(select_btn_frame, text="选择前5热点", command=self.select_top_hot_industries_tree,
                  bg="#e67e22", fg="white", width=12).pack(side=tk.LEFT, padx=2)
        
        # 3. 均线偏离度选股
        ma_frame = tk.LabelFrame(left_main, text="3. 均线偏离度选股", 
                                 font=("微软雅黑", 11, "bold"), bg='#f0f0f0', padx=10, pady=10)
        ma_frame.pack(fill=tk.X, pady=5, padx=10)
        
        line1 = tk.Frame(ma_frame, bg='#f0f0f0')
        line1.pack(fill=tk.X, pady=5)
        tk.Label(line1, text="参考均线:", font=("微软雅黑", 10), bg='#f0f0f0', width=8).pack(side=tk.LEFT)
        self.ma_selected = tk.StringVar(value="30")
        ma_menu = ttk.Combobox(line1, textvariable=self.ma_selected, values=[5,10,20,30,60,120,250], width=8)
        ma_menu.pack(side=tk.LEFT, padx=5)
        tk.Label(line1, text="日均线", font=("微软雅黑", 10), bg='#f0f0f0').pack(side=tk.LEFT)
        
        line2 = tk.Frame(ma_frame, bg='#f0f0f0')
        line2.pack(fill=tk.X, pady=5)
        tk.Label(line2, text="偏离范围:", font=("微软雅黑", 10), bg='#f0f0f0', width=8).pack(side=tk.LEFT)
        self.ma_deviation_lower = tk.Entry(line2, width=6, font=("微软雅黑", 10))
        self.ma_deviation_lower.insert(0, "-20")
        self.ma_deviation_lower.pack(side=tk.LEFT, padx=2)
        tk.Label(line2, text="% ~", font=("微软雅黑", 10), bg='#f0f0f0').pack(side=tk.LEFT)
        self.ma_deviation_upper = tk.Entry(line2, width=6, font=("微软雅黑", 10))
        self.ma_deviation_upper.insert(0, "20")
        self.ma_deviation_upper.pack(side=tk.LEFT, padx=2)
        tk.Label(line2, text="%", font=("微软雅黑", 10), bg='#f0f0f0').pack(side=tk.LEFT)
        
        # 4. 参数设置（去掉市值控件）
        param_frame = tk.LabelFrame(left_main, text="4. 参数设置", 
                                    font=("微软雅黑", 11, "bold"), bg='#f0f0f0', padx=10, pady=10)
        param_frame.pack(fill=tk.X, pady=5, padx=10)
        
        param_inner = tk.Frame(param_frame, bg='#f0f0f0', height=200)
        param_inner.pack(fill=tk.X, pady=5)
        param_inner.pack_propagate(False)
        
        # 第1行 - 换手率
        row1 = tk.Frame(param_inner, bg='#f0f0f0')
        row1.pack(fill=tk.X, pady=3)
        tk.Label(row1, text="换手率%:", bg='#f0f0f0', font=("微软雅黑", 9), width=8).pack(side=tk.LEFT)
        self.entry_min_turn = tk.Entry(row1, width=6)
        self.entry_min_turn.insert(0, "0.5")
        self.entry_min_turn.pack(side=tk.LEFT, padx=2)
        tk.Label(row1, text="~", bg='#f0f0f0').pack(side=tk.LEFT)
        self.entry_max_turn = tk.Entry(row1, width=6)
        self.entry_max_turn.insert(0, "50")
        self.entry_max_turn.pack(side=tk.LEFT, padx=2)
        
        # 第2行 - 高位阈值
        row2 = tk.Frame(param_inner, bg='#f0f0f0')
        row2.pack(fill=tk.X, pady=3)
        tk.Label(row2, text="高位阈值:", bg='#f0f0f0', font=("微软雅黑", 9), width=8).pack(side=tk.LEFT)
        self.entry_high_pos = tk.Entry(row2, width=6)
        self.entry_high_pos.insert(0, "0.98")
        self.entry_high_pos.pack(side=tk.LEFT, padx=2)
        
        # 第3行 - 形态识别开关
        row3 = tk.Frame(param_inner, bg='#f0f0f0')
        row3.pack(fill=tk.X, pady=3)
        self.var_enable_my_pattern = tk.BooleanVar(value=True)
        tk.Checkbutton(row3, text="启用形态识别", 
                       variable=self.var_enable_my_pattern,
                       bg='#f0f0f0', font=("微软雅黑", 9)).pack(side=tk.LEFT)
        
        # 第4行 - 右侧买点开关
        row4 = tk.Frame(param_inner, bg='#f0f0f0')
        row4.pack(fill=tk.X, pady=3)
        self.var_enable_buy_points = tk.BooleanVar(value=True)
        tk.Checkbutton(row4, text="显示右侧买点", 
                       variable=self.var_enable_buy_points,
                       bg='#f0f0f0', font=("微软雅黑", 9)).pack(side=tk.LEFT)
        
        # 第5行 - 调试模式开关
        row5 = tk.Frame(param_inner, bg='#f0f0f0')
        row5.pack(fill=tk.X, pady=3)
        self.var_debug_mode = tk.BooleanVar(value=False)
        tk.Checkbutton(row5, text="调试模式", 
                       variable=self.var_debug_mode,
                       bg='#f0f0f0', font=("微软雅黑", 9)).pack(side=tk.LEFT)
        
        # 右侧监控面板
        right = tk.Frame(paned, bg='#f0f0f0')
        paned.add(right, weight=1)
        
        log_frame = tk.LabelFrame(right, text="运行日志", font=("微软雅黑", 11, "bold"), bg='#f0f0f0')
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        self.text_log = tk.Text(log_frame, font=("Consolas", 9), height=12, bg='#1a1a2e', fg='#00ff88')
        self.text_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.monitor = MonitorPanel(right, self.log)
    
    def change_data_source(self):
        global DATA_SOURCE
        DATA_SOURCE = self.data_source_var.get()
        self.log(f"数据源已切换为: {DATA_SOURCE}")
        
        # 显示数据源健康状态
        def check_health():
            results = check_data_source_health()
            self.log(f"数据源状态 - baostock: {'✓ 可用' if results['baostock'] else '✗ 不可用'}, "
                    f"东方财富: {'✓ 可用' if results['eastmoney'] else '✗ 不可用'}")
        
        threading.Thread(target=check_health, daemon=True).start()
    
    def select_all_industries_tree(self):
        for item in self.industry_tree.get_children():
            self.industry_tree.selection_add(item)
    
    def clear_industries_tree(self):
        self.industry_tree.selection_remove(*self.industry_tree.selection())
    
    def select_top_hot_industries_tree(self):
        if not self.hot_ranking:
            self.log("正在计算热点，请稍后再试...")
            self.refresh_hot_ranking()
            return
        
        self.industry_tree.selection_remove(*self.industry_tree.selection())
        for i, item in enumerate(self.industry_tree.get_children()):
            if i < 5:
                self.industry_tree.selection_add(item)
        self.log(f"已选择热度前5行业")
    
    def refresh_hot_ranking(self):
        if not self.all_industries:
            self.hot_status_label.config(text="请等待行业加载完成...", fg="#e74c3c")
            self.log("行业数据未加载完成，请稍后再试")
            return
        
        self.hot_status_label.config(text="正在计算热点排序...", fg="#e67e22")
        self.refresh_hot_btn.config(state="disabled")
        
        def calc_hot():
            try:
                self.q.put(("log", "开始计算热点排序..."))
                
                industries_df = get_industry_stocks()
                if industries_df.empty:
                    self.q.put(("hot_error", "获取行业数据失败"))
                    return
                
                end_date = datetime.now().strftime('%Y-%m-%d')
                start_date = (datetime.now() - timedelta(days=10)).strftime('%Y-%m-%d')
                
                industry_stats = {}
                for industry in self.all_industries:
                    industry_stats[industry] = {
                        'limit_up_count': 0,
                        'total_turnover': 0,
                        'stock_count': 0,
                        'up_count': 0,
                    }
                
                total_industries = len(self.all_industries)
                processed = 0
                
                for industry in self.all_industries:
                    processed += 1
                    if processed % 10 == 0:
                        self.q.put(("log", f"计算热点进度: {processed}/{total_industries}"))
                    
                    stocks_in_industry = industries_df[industries_df['industry'] == industry].to_dict('records')
                    industry_stats[industry]['stock_count'] = len(stocks_in_industry)
                    
                    for stock in stocks_in_industry[:20]:
                        code = stock['code']
                        df = fetch_stock_data_with_fallback(code, start_date, end_date, use_cache=True)
                        if df is None or len(df) < 5:
                            continue
                        
                        stats = industry_stats[industry]
                        last = df.iloc[-1]
                        
                        for i in range(-5, 0):
                            if i >= len(df):
                                continue
                            pct = df.iloc[i].get('pctChg', 0)
                            if not pd.isna(pct) and pct >= 9.5:
                                stats['limit_up_count'] += 1
                                break
                        
                        turn = last.get('turn', 0)
                        if not pd.isna(turn) and turn > 0:
                            stats['total_turnover'] += turn
                        
                        pct = last.get('pctChg', 0)
                        if not pd.isna(pct) and pct > 0:
                            stats['up_count'] += 1
                
                hot_list = []
                for industry, stats in industry_stats.items():
                    stock_cnt = stats['stock_count']
                    if stock_cnt == 0:
                        continue
                    
                    limit_score = min(stats['limit_up_count'] * 5, 40)
                    
                    avg_turn = stats['total_turnover'] / stock_cnt if stock_cnt > 0 else 0
                    if 3 <= avg_turn <= 25:
                        turn_score = 30
                    elif 2 <= avg_turn <= 30:
                        turn_score = 20
                    elif avg_turn > 0:
                        turn_score = 10
                    else:
                        turn_score = 0
                    
                    up_ratio = stats['up_count'] / stock_cnt if stock_cnt > 0 else 0
                    up_score = up_ratio * 30
                    
                    total_score = limit_score + turn_score + up_score
                    
                    hot_list.append({
                        'industry': industry,
                        'score': round(total_score, 1),
                        'limit_count': stats['limit_up_count'],
                        'avg_turn': round(avg_turn, 1),
                        'up_ratio': round(up_ratio * 100, 1)
                    })
                
                hot_list.sort(key=lambda x: -x['score'])
                for i, item in enumerate(hot_list):
                    item['rank'] = i + 1
                
                self.q.put(("hot_ranking", hot_list))
                self.q.put(("log", f"热点计算完成，共{len(hot_list)}个行业"))
                
            except Exception as e:
                self.q.put(("hot_error", f"计算热点失败: {str(e)}"))
            finally:
                self.q.put(("hot_done", None))
        
        threading.Thread(target=calc_hot, daemon=True).start()
    
    def update_hot_ranking_display(self, hot_list):
        for item in self.industry_tree.get_children():
            self.industry_tree.delete(item)
        
        self.hot_ranking = hot_list
        self.industry_hot_score = {}
        
        for item in hot_list:
            industry = item['industry']
            score = item['score']
            rank = item['rank']
            limit_count = item['limit_count']
            avg_turn = item['avg_turn']
            up_ratio = item['up_ratio']
            
            self.industry_hot_score[industry] = score
            
            if score >= 80:
                prefix = "🔥"
            elif score >= 60:
                prefix = "⭐"
            elif score >= 40:
                prefix = "●"
            else:
                prefix = "○"
            
            self.industry_tree.insert("", tk.END, values=(
                prefix, rank, industry, score, limit_count, avg_turn, up_ratio
            ))
        
        self.hot_status_label.config(text=f"热点计算完成，共{len(hot_list)}个行业", fg="#27ae60")
        self.refresh_hot_btn.config(state="normal")
    
    def load_config(self):
        if not os.path.exists(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, 'r') as f:
                cfg = json.load(f)
            self.var600.set(cfg.get("600", True))
            self.var300.set(cfg.get("300", True))
            self.var000.set(cfg.get("000", True))
            self.var688.set(cfg.get("688", False))
            self.ma_selected.set(cfg.get("ma_period", "30"))
            self.ma_deviation_lower.delete(0, tk.END)
            self.ma_deviation_lower.insert(0, cfg.get("ma_deviation_lower", "-20"))
            self.ma_deviation_upper.delete(0, tk.END)
            self.ma_deviation_upper.insert(0, cfg.get("ma_deviation_upper", "20"))
            self.entry_min_turn.delete(0, tk.END)
            self.entry_min_turn.insert(0, cfg.get("min_turn", "0.5"))
            self.entry_max_turn.delete(0, tk.END)
            self.entry_max_turn.insert(0, cfg.get("max_turn", "50"))
            self.entry_high_pos.delete(0, tk.END)
            self.entry_high_pos.insert(0, cfg.get("high_pos", "0.98"))
            self.var_enable_my_pattern.set(cfg.get("enable_my_pattern", True))
            self.var_enable_buy_points.set(cfg.get("enable_buy_points", True))
            self.var_debug_mode.set(cfg.get("debug_mode", False))
            self.data_source_var.set(cfg.get("data_source", "auto"))
            global DATA_SOURCE
            DATA_SOURCE = self.data_source_var.get()
        except:
            pass
    
    def save_config(self):
        cfg = {
            "600": self.var600.get(), "300": self.var300.get(),
            "000": self.var000.get(), "688": self.var688.get(),
            "ma_period": self.ma_selected.get(),
            "ma_deviation_lower": self.ma_deviation_lower.get(),
            "ma_deviation_upper": self.ma_deviation_upper.get(),
            "min_turn": self.entry_min_turn.get(),
            "max_turn": self.entry_max_turn.get(),
            "high_pos": self.entry_high_pos.get(),
            "enable_my_pattern": self.var_enable_my_pattern.get(),
            "enable_buy_points": self.var_enable_buy_points.get(),
            "debug_mode": self.var_debug_mode.get(),
            "data_source": self.data_source_var.get()
        }
        with open(CONFIG_FILE, 'w') as f:
            json.dump(cfg, f)
    
    def reset_config(self):
        self.ma_selected.set("30")
        self.ma_deviation_lower.delete(0, tk.END)
        self.ma_deviation_lower.insert(0, "-20")
        self.ma_deviation_upper.delete(0, tk.END)
        self.ma_deviation_upper.insert(0, "20")
        self.entry_min_turn.delete(0, tk.END)
        self.entry_min_turn.insert(0, "0.5")
        self.entry_max_turn.delete(0, tk.END)
        self.entry_max_turn.insert(0, "50")
        self.entry_high_pos.delete(0, tk.END)
        self.entry_high_pos.insert(0, "0.98")
        self.var_enable_my_pattern.set(True)
        self.var_enable_buy_points.set(True)
        self.var_debug_mode.set(False)
        self.log("已重置为默认配置")
    
    def clear_cache(self):
        if messagebox.askyesno("确认", "清空缓存将重新下载数据，确定吗？"):
            count = clear_cache()
            self.log(f"已清空{count}个缓存文件")
    
    def cancel_scan(self):
        global stop_scan_flag
        stop_scan_flag = True
        self.start_btn.config(state="normal", text="开始选股")
        self.cancel_btn.config(state="disabled", text="取消")
        self.log("正在取消扫描...")
    
    def exit_app(self):
        if messagebox.askyesno("确认", "确定退出吗？"):
            self.save_config()
            self.monitor.stop()
            self.root.quit()
    
    def start_scan(self):
        global stop_scan_flag
        stop_scan_flag = False
        
        boards = []
        if self.var600.get(): boards.append("600")
        if self.var300.get(): boards.append("300")
        if self.var000.get(): boards.append("000")
        if self.var688.get(): boards.append("688")
        
        if not boards:
            messagebox.showwarning("提示", "请至少选择一个板块")
            return
        
        selected_items = self.industry_tree.selection()
        clean_industries = []
        for item in selected_items:
            values = self.industry_tree.item(item, "values")
            if values and len(values) >= 3:
                clean_industries.append(values[2])
        
        if not clean_industries:
            messagebox.showwarning("提示", "请至少选择一个行业")
            return
        
        try:
            params = {
                "selected_boards": boards,
                "selected_industries": clean_industries,
                "ma_period": int(self.ma_selected.get()),
                "ma_deviation_lower": float(self.ma_deviation_lower.get()),
                "ma_deviation_upper": float(self.ma_deviation_upper.get()),
                "min_turn": float(self.entry_min_turn.get()),
                "max_turn": float(self.entry_max_turn.get()),
                "enable_limit_strength": False,
                "limit_days_max": 30,
                "enable_huangniu": False,
                "enable_my_pattern": self.var_enable_my_pattern.get(),
                "enable_buy_points": self.var_enable_buy_points.get(),
                "high_pos_threshold": float(self.entry_high_pos.get()),
                "debug_mode": self.var_debug_mode.get(),
                "log_func": self.log,
                "progress_func": self.set_progress
            }
        except ValueError as e:
            messagebox.showerror("参数错误", f"请输入正确的数值: {e}")
            return
        
        self.text_log.delete(1.0, tk.END)
        self.progress["value"] = 0
        self.start_btn.config(state="disabled", text="扫描中...")
        self.cancel_btn.config(state="normal", text="取消")
        self.save_config()
        
        self.scan_thread = threading.Thread(target=main, args=(params, self.log, self.set_progress), daemon=True)
        self.scan_thread.start()
    
    def log(self, msg):
        self.q.put(("log", msg))
    
    def set_progress(self, val):
        self.q.put(("progress", val))
    
    def update_ui(self):
        while not self.q.empty():
            item = self.q.get()
            if item[0] == "log":
                self.text_log.insert(tk.END, item[1] + "\n")
                self.text_log.see(tk.END)
            elif item[0] == "progress":
                self.progress["value"] = item[1]
                self.progress_label.config(text=f"{item[1]}%")
            elif item[0] == "industries":
                for existing in self.industry_tree.get_children():
                    self.industry_tree.delete(existing)
                if item[1]:
                    self.all_industries = item[1]
                    for ind in item[1]:
                        self.industry_tree.insert("", tk.END, values=("○", "-", ind, "-", "-", "-", "-"))
                    self.log(f"加载完成，共{len(item[1])}个行业")
                    self.hot_status_label.config(text="点击刷新热点排序计算热点", fg="#3498db")
                else:
                    self.all_industries = []
                    self.industry_tree.insert("", tk.END, values=("", "", "无行业数据，请检查网络", "", "", "", ""))
                    self.hot_status_label.config(text="行业加载失败", fg="#e74c3c")
            elif item[0] == "hot_ranking":
                self.update_hot_ranking_display(item[1])
            elif item[0] == "hot_error":
                self.hot_status_label.config(text=item[1], fg="#e74c3c")
                self.refresh_hot_btn.config(state="normal")
            elif item[0] == "hot_done":
                pass
            elif item[0] == "enable_button":
                self.start_btn.config(state="normal", text="开始选股")
                self.cancel_btn.config(state="disabled", text="取消")
        self.root.after(100, self.update_ui)
    
    def load_industries(self):
        self.hot_status_label.config(text="正在加载行业数据...", fg="#e67e22")
        
        def load():
            try:
                industries = get_industries_list()
                if industries and len(industries) > 0:
                    self.q.put(("industries", industries))
                else:
                    self.q.put(("industries", []))
                    self.q.put(("log", "警告：行业数据加载失败，请检查网络后重试"))
            except Exception as e:
                self.q.put(("log", f"加载行业失败: {str(e)}"))
                self.q.put(("industries", []))
        
        threading.Thread(target=load, daemon=True).start()
    
    def run(self):
        self.root.mainloop()
        bs_logout()

# ==============================
# 主程序入口
# ==============================
if __name__ == "__main__":
    print("=" * 60)
    print("黄牛选股器 v2.0 (无市值限制版)")
    print("=" * 60)
    
    # 测试数据源
    print("\n正在测试数据源...")
    results = check_data_source_health()
    print(f"baostock: {'✓ 可用' if results['baostock'] else '✗ 不可用'}")
    print(f"东方财富: {'✓ 可用' if results['eastmoney'] else '✗ 不可用'}")
    
    if not results['baostock'] and not results['eastmoney']:
        print("\n⚠ 警告: 所有数据源均不可用！")
        print("请检查:")
        print("1. 网络连接是否正常")
        print("2. 防火墙是否允许Python访问网络")
        print("3. 如需使用东方财富，请安装: pip install akshare --upgrade")
        print("\n程序将继续启动，但可能无法获取数据...")
        print("=" * 60)
    
    app = StockPickerGUI()
    app.run()
