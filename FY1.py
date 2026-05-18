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
   

# ==============================
# 交易记录管理（保持不变）
# ==============================
class TradeRecorder:
    def __init__(self):
        self.today = datetime.now().strftime('%Y-%m-%d')
        self.record_file = os.path.join(RECORD_DIR, f"交易记录_{self.today}.json")
        self.trades = self.load()
   
   

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
