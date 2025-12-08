import pandas as pd
import numpy as np
from collections import Counter, defaultdict
import matplotlib.pyplot as plt
from datetime import datetime
import warnings
import re
import os
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
warnings.filterwarnings('ignore')

class LotteryAnalyzer:
    def __init__(self, file_path):
        """
        初始化分析器
        Args:
            file_path: Excel文件路径
        """
        self.file_path = file_path
        self.data = None
        self.front_range = (1, 35)  # 前区范围
        self.back_range = (1, 12)   # 后区范围
        self.analysis_results = {}  # 存储分析结果
        self.recommendations = {}   # 存储推荐结果
        self.comprehensive_recommendations = []  # 综合推荐结果
        self.hot_selected_recommendations = []  # 热号精选推荐
        self.skew_trends = {}  # 斜连号趋势
        self.consecutive_appearances = {}  # 连续出现分析
        self.load_data()
        
    def load_data(self):
        """加载并预处理数据"""
        try:
            # 读取Excel文件，跳过前两行表头
            df = pd.read_excel(self.file_path, header=None)
            
            print(f"✓ 数据形状: {df.shape}")
            print("✓ 正在处理数据格式...")
            
            # 重新构建DataFrame
            processed_data = []
            
            for i in range(2, len(df)):
                row = df.iloc[i]
                if len(row) >= 9:  # 确保有9列数据
                    period = row[0]  # 期号
                    date = row[1]    # 开奖日期
                    
                    # 前区号码 (位置2-6)
                    front_numbers = []
                    for j in range(2, 7):
                        num = row[j]
                        if pd.notna(num):
                            try:
                                front_numbers.append(int(num))
                            except:
                                try:
                                    front_numbers.append(int(float(num)))
                                except:
                                    pass
                    
                    # 后区号码 (位置7-8)
                    back_numbers = []
                    for j in range(7, 9):
                        num = row[j]
                        if pd.notna(num):
                            try:
                                back_numbers.append(int(num))
                            except:
                                try:
                                    back_numbers.append(int(float(num)))
                                except:
                                    pass
                    
                    # 验证数据
                    if (len(front_numbers) == 5 and len(back_numbers) == 2 and
                        all(1 <= num <= 35 for num in front_numbers) and
                        all(1 <= num <= 12 for num in back_numbers)):
                        
                        processed_data.append({
                            '期号': period,
                            '开奖日期': date,
                            '前区': front_numbers,
                            '后区': back_numbers
                        })
            
            self.data = pd.DataFrame(processed_data)
            
            # 按期号排序
            self.data = self.data.sort_values('期号').reset_index(drop=True)
            
            print(f"✓ 成功加载 {len(self.data)} 期开奖数据")
            print(f"✓ 数据范围: 期号{self.data['期号'].min()} - {self.data['期号'].max()}")
            
            # 存储基本信息
            self.analysis_results['基本信息'] = {
                '总期数': len(self.data),
                '最早期号': int(self.data['期号'].min()),
                '最晚期号': int(self.data['期号'].max()),
                '最早日期': str(self.data['开奖日期'].min()),
                '最晚日期': str(self.data['开奖日期'].max()),
                '前区范围': f"{self.front_range[0]}-{self.front_range[1]}",
                '后区范围': f"{self.back_range[0]}-{self.back_range[1]}"
            }
            
            # 分离不同周期的数据
            self.recent_30 = self.data.tail(30).copy()
            self.recent_60 = self.data.tail(60).copy()
            self.recent_100 = self.data.tail(100).copy()
            self.all_data = self.data.copy()
            
        except Exception as e:
            print(f"✗ 数据加载失败: {e}")
            import traceback
            traceback.print_exc()
            raise
    
    def is_prime(self, n):
        """判断是否为质数"""
        if n < 2:
            return False
        for i in range(2, int(n**0.5) + 1):
            if n % i == 0:
                return False
        return True
    
    def analyze_skew_numbers(self, period_data, period_name):
        """分析斜连号趋势"""
        print(f"正在分析{period_name}斜连号趋势...")
        
        if len(period_data) < 5:
            print(f"  ✗ {period_name}数据不足，跳过斜连号分析")
            return {}
        
        # 将数据按开奖期号排序
        sorted_data = period_data.sort_values('期号')
        
        # 斜连号分析结果
        skew_analysis = {
            '正斜连': {},
            '反斜连': {},
            '连续出现': {},
            '预测号码': []
        }
        
        # 1. 正斜连分析（递增或递减的斜线）
        # 检查每个位置（假设我们按每期号码排序后分析）
        # 正斜连：号码在连续几期中呈等差数列，差值为±1
        
        # 获取最近几期的号码数据
        recent_periods = min(10, len(sorted_data))
        recent_data = sorted_data.tail(recent_periods)
        
        # 存储每期的前区号码（已排序）
        period_numbers = []
        for _, row in recent_data.iterrows():
            period_numbers.append(sorted(row['前区']))
        
        # 分析每个号码的斜连情况
        for i in range(len(period_numbers) - 1):
            current_period = period_numbers[i]
            next_period = period_numbers[i + 1]
            
            # 检查正斜连（号码相差1）
            for num1 in current_period:
                for num2 in next_period:
                    diff = num2 - num1
                    
                    # 正斜连：相差1（递增或递减）
                    if abs(diff) == 1:
                        direction = "递增" if diff > 0 else "递减"
                        key = f"{num1}→{num2}({direction})"
                        
                        if key not in skew_analysis['正斜连']:
                            skew_analysis['正斜连'][key] = {
                                '开始号码': num1,
                                '结束号码': num2,
                                '方向': direction,
                                '长度': 2,
                                '期数': [i, i+1]
                            }
                        else:
                            # 如果已经存在，检查是否可以延长
                            existing = skew_analysis['正斜连'][key]
                            if i == existing['期数'][-1]:
                                existing['长度'] += 1
                                existing['结束号码'] = num2
                                existing['期数'].append(i+1)
        
        # 2. 反斜连分析（先增后减或先减后增）
        if len(period_numbers) >= 3:
            for i in range(len(period_numbers) - 2):
                period1 = period_numbers[i]
                period2 = period_numbers[i + 1]
                period3 = period_numbers[i + 2]
                
                for num1 in period1:
                    for num2 in period2:
                        diff1 = num2 - num1
                        
                        # 检查是否是第一步斜连
                        if abs(diff1) == 1:
                            # 检查第三步是否形成反斜
                            for num3 in period3:
                                diff2 = num3 - num2
                                
                                # 反斜连：方向相反
                                if abs(diff2) == 1 and diff1 * diff2 < 0:
                                    direction = "先增后减" if diff1 > 0 else "先减后增"
                                    key = f"{num1}→{num2}→{num3}({direction})"
                                    
                                    skew_analysis['反斜连'][key] = {
                                        '开始号码': num1,
                                        '中间号码': num2,
                                        '结束号码': num3,
                                        '方向': direction,
                                        '长度': 3,
                                        '期数': [i, i+1, i+2]
                                    }
        
        # 3. 分析连续出现的号码
        # 统计每个号码连续出现的次数
        all_numbers = []
        all_periods = []
        
        for _, row in sorted_data.iterrows():
            all_numbers.append(row['前区'])
            all_periods.append(row['期号'])
        
        # 分析每个号码的连续出现情况
        for num in range(1, 36):
            consecutive_count = 0
            max_consecutive = 0
            current_streak = 0
            appearances = []  # 记录出现期数
            
            for i, numbers in enumerate(all_numbers):
                if num in numbers:
                    current_streak += 1
                    appearances.append(all_periods[i])
                else:
                    if current_streak > max_consecutive:
                        max_consecutive = current_streak
                    current_streak = 0
            
            # 检查最后一期
            if current_streak > max_consecutive:
                max_consecutive = current_streak
            
            # 计算总出现次数
            total_appearances = sum(1 for numbers in all_numbers if num in numbers)
            
            if max_consecutive >= 2:  # 至少连续出现2期才记录
                skew_analysis['连续出现'][num] = {
                    '最长连续': max_consecutive,
                    '总出现次数': total_appearances,
                    '出现期数': appearances[-max_consecutive:] if appearances else []
                }
        
        # 4. 预测可能的斜连号
        if len(period_numbers) >= 2:
            last_period = period_numbers[-1]
            second_last_period = period_numbers[-2]
            
            # 预测正斜连延续
            for num in last_period:
                # 检查递增斜连
                if num + 1 <= 35:
                    skew_analysis['预测号码'].append({
                        '号码': num + 1,
                        '类型': '正斜连延续',
                        '理由': f'延续{num}的递增斜连',
                        '权重': 1.5
                    })
                
                # 检查递减斜连
                if num - 1 >= 1:
                    skew_analysis['预测号码'].append({
                        '号码': num - 1,
                        '类型': '正斜连延续',
                        '理由': f'延续{num}的递减斜连',
                        '权重': 1.5
                    })
            
            # 检查反斜连可能
            for i in range(len(period_numbers) - 2, len(period_numbers)):
                if i >= 0 and i < len(period_numbers) - 1:
                    current = period_numbers[i]
                    prev = period_numbers[i-1] if i > 0 else []
                    
                    for num in prev:
                        if num + 1 in current:
                            # 可能形成反斜连（先增后减）
                            skew_analysis['预测号码'].append({
                                '号码': num,
                                '类型': '反斜连形成',
                                '理由': f'可能形成{num}→{num+1}→{num}的反斜连',
                                '权重': 2.0
                            })
        
        return skew_analysis
    
    def analyze_consecutive_appearances(self, period_data, period_name):
        """分析号码连续出现情况"""
        print(f"正在分析{period_name}连续出现情况...")
        
        if len(period_data) < 5:
            print(f"  ✗ {period_name}数据不足，跳过连续出现分析")
            return {}
        
        sorted_data = period_data.sort_values('期号')
        
        consecutive_analysis = {
            '当前连续': {},
            '历史连续': {},
            '热号追踪': {},
            '冷号反弹': {}
        }
        
        # 分析每个号码的连续出现情况
        for num in range(1, 36):
            current_streak = 0
            max_streak = 0
            streak_history = []
            in_streak = False
            streak_start = None
            
            # 按时间顺序检查
            for idx, (_, row) in enumerate(sorted_data.iterrows()):
                numbers = row['前区']
                
                if num in numbers:
                    if not in_streak:
                        in_streak = True
                        streak_start = row['期号']
                        current_streak = 1
                    else:
                        current_streak += 1
                else:
                    if in_streak:
                        # 记录结束的连续
                        streak_history.append({
                            '开始期': streak_start,
                            '结束期': sorted_data.iloc[idx-1]['期号'],
                            '长度': current_streak
                        })
                        if current_streak > max_streak:
                            max_streak = current_streak
                        in_streak = False
                        current_streak = 0
            
            # 检查最后一期是否在连续中
            if in_streak:
                streak_history.append({
                    '开始期': streak_start,
                    '结束期': sorted_data.iloc[-1]['期号'],
                    '长度': current_streak
                })
                if current_streak > max_streak:
                    max_streak = current_streak
                
                # 记录当前连续
                consecutive_analysis['当前连续'][num] = current_streak
        
        # 分析热号和冷号
        all_numbers = []
        for _, row in sorted_data.iterrows():
            all_numbers.extend(row['前区'])
        
        number_counts = Counter(all_numbers)
        total_periods = len(sorted_data)
        
        # 热号：出现频率高的号码
        hot_numbers = number_counts.most_common(10)
        for num, count in hot_numbers:
            consecutive_analysis['热号追踪'][num] = {
                '出现次数': count,
                '出现频率': f"{count/total_periods*100:.1f}%",
                '最近出现': '是' if num in sorted_data.iloc[-1]['前区'] else '否'
            }
        
        # 冷号：出现频率低的号码
        cold_numbers = sorted(number_counts.items(), key=lambda x: x[1])[:10]
        for num, count in cold_numbers:
            # 检查是否近期有出现
            recent_appearance = False
            recent_periods = min(10, total_periods)
            for i in range(1, recent_periods + 1):
                if num in sorted_data.iloc[-i]['前区']:
                    recent_appearance = True
                    break
            
            consecutive_analysis['冷号反弹'][num] = {
                '出现次数': count,
                '出现频率': f"{count/total_periods*100:.1f}%",
                '近期出现': '是' if recent_appearance else '否',
                '遗漏期数': self.calculate_missing(sorted_data, num)
            }
        
        return consecutive_analysis
    
    def calculate_missing(self, period_data, num):
        """计算某个号码的遗漏期数"""
        missing = 0
        for i in range(len(period_data) - 1, -1, -1):
            if num in period_data.iloc[i]['前区']:
                break
            missing += 1
        return missing
    
    def analyze_period(self, period_data, period_name):
        """分析指定周期的数据"""
        print(f"\n正在分析{period_name}数据...")
        
        analysis = {
            '周期名称': period_name,
            '数据期数': len(period_data),
            '开始期号': int(period_data['期号'].min()),
            '结束期号': int(period_data['期号'].max()),
            '热门号码': {},
            '冷门号码': {},
            '遗漏分析': {},
            '模式分析': {},
            '斜连号分析': {},
            '连续出现分析': {},
            '推荐依据': []
        }
        
        # 1. 热门号码分析
        all_front = []
        all_back = []
        
        for _, row in period_data.iterrows():
            all_front.extend(row['前区'])
            all_back.extend(row['后区'])
        
        front_freq = Counter(all_front)
        back_freq = Counter(all_back)
        
        analysis['热门号码']['前区热门'] = dict(front_freq.most_common(10))
        analysis['热门号码']['后区热门'] = dict(back_freq.most_common(5))
        
        # 冷门号码
        analysis['冷门号码']['前区冷门'] = dict(sorted(front_freq.items(), key=lambda x: x[1])[:10])
        
        # 2. 模式分析
        stats = {
            '奇偶比': [],
            '大小比': [],
            '和值': [],
            '跨度': [],
            '连号次数': 0,
            '重复前区': 0,
            '重复后区': 0
        }
        
        for i in range(len(period_data)):
            row = period_data.iloc[i]
            front = row['前区']
            back = row['后区']
            
            # 奇偶比
            odd_count = sum(1 for num in front if num % 2 == 1)
            stats['奇偶比'].append(f"{odd_count}:{5-odd_count}")
            
            # 大小比
            big_count = sum(1 for num in front if num > 18)
            stats['大小比'].append(f"{big_count}:{5-big_count}")
            
            # 和值
            stats['和值'].append(sum(front))
            
            # 跨度
            stats['跨度'].append(max(front) - min(front))
            
            # 连号检测
            sorted_front = sorted(front)
            for j in range(4):
                if sorted_front[j+1] - sorted_front[j] == 1:
                    stats['连号次数'] += 1
                    break
            
            # 重复号检测
            if i > 0:
                prev_front = period_data.iloc[i-1]['前区']
                prev_back = period_data.iloc[i-1]['后区']
                stats['重复前区'] += len(set(front) & set(prev_front))
                stats['重复后区'] += len(set(back) & set(prev_back))
        
        analysis['模式分析'] = stats
        
        # 3. 遗漏分析
        front_missing = {num: 0 for num in range(1, 36)}
        back_missing = {num: 0 for num in range(1, 13)}
        
        sorted_period = period_data.sort_values('期号')
        
        for _, row in sorted_period.iterrows():
            front = row['前区']
            back = row['后区']
            
            for num in front_missing:
                if num in front:
                    front_missing[num] = 0
                else:
                    front_missing[num] += 1
            
            for num in back_missing:
                if num in back:
                    back_missing[num] = 0
                else:
                    back_missing[num] += 1
        
        analysis['遗漏分析']['前区热遗漏'] = dict(sorted(front_missing.items(), key=lambda x: x[1])[:10])
        analysis['遗漏分析']['前区冷遗漏'] = dict(sorted(front_missing.items(), key=lambda x: x[1], reverse=True)[:10])
        analysis['遗漏分析']['后区遗漏'] = back_missing
        
        # 4. 斜连号分析
        analysis['斜连号分析'] = self.analyze_skew_numbers(period_data, period_name)
        
        # 5. 连续出现分析
        analysis['连续出现分析'] = self.analyze_consecutive_appearances(period_data, period_name)
        
        return analysis, front_freq, back_freq, front_missing, back_missing
    
    def generate_recommendations_for_period(self, period_data, period_name, front_freq, back_freq, front_missing, back_missing, skew_analysis, consecutive_analysis):
        """为指定周期生成推荐号码"""
        
        # 斜连号权重设置
        skew_weights = {
            '2连斜': 1.2,    # 2期斜连
            '3连斜': 1.5,    # 3期斜连
            '4连斜': 2.0,    # 4期斜连
            '5连斜以上': 2.5, # 5期以上斜连
            '正斜连': 1.3,   # 正斜连
            '反斜连': 1.8,   # 反斜连
            '连续出现2期': 1.2, # 连续出现2期
            '连续出现3期': 1.5, # 连续出现3期
            '连续出现4期以上': 2.0 # 连续出现4期以上
        }
        
        # 综合评分算法
        def score_front(num):
            """前区号码综合评分"""
            # 频率得分 (权重30%)
            max_front_freq = max(front_freq.values()) if front_freq else 1
            freq_score = (front_freq.get(num, 0) / max_front_freq * 30)
            
            # 遗漏得分 (权重25%)
            avg_missing = np.mean(list(front_missing.values())) if front_missing else 1
            missing_norm = front_missing.get(num, 0) / avg_missing if avg_missing > 0 else 1
            
            if missing_norm < 0.5:  # 过热
                missing_score = 10
            elif missing_norm < 1.0:  # 温热
                missing_score = 20
            elif missing_norm < 2.0:  # 正常
                missing_score = 25
            else:  # 冷门
                missing_score = 15
            
            # 奇偶趋势得分 (权重15%)
            recent_odd_ratio = 0.5  # 默认
            if len(period_data) > 0:
                odd_count = sum(1 for row in period_data.iterrows() for n in row[1]['前区'] if n % 2 == 1)
                total_count = len(period_data) * 5
                recent_odd_ratio = odd_count / total_count if total_count > 0 else 0.5
            
            if num % 2 == 1 and recent_odd_ratio < 0.5:
                odd_even_score = 15
            elif num % 2 == 0 and recent_odd_ratio > 0.5:
                odd_even_score = 15
            else:
                odd_even_score = 10
            
            # 大小趋势得分 (权重10%)
            recent_big_ratio = 0.5  # 默认
            if len(period_data) > 0:
                big_count = sum(1 for row in period_data.iterrows() for n in row[1]['前区'] if n > 18)
                total_count = len(period_data) * 5
                recent_big_ratio = big_count / total_count if total_count > 0 else 0.5
            
            if num > 18 and recent_big_ratio < 0.5:
                size_score = 10
            elif num <= 18 and recent_big_ratio > 0.5:
                size_score = 10
            else:
                size_score = 5
            
            # 斜连号趋势得分 (权重15%)
            skew_score = 0
            
            # 检查斜连号预测
            for prediction in skew_analysis.get('预测号码', []):
                if prediction['号码'] == num:
                    skew_score += prediction['权重'] * 5
            
            # 检查正斜连
            for skew_key, skew_info in skew_analysis.get('正斜连', {}).items():
                if skew_info['结束号码'] == num:
                    length = skew_info['长度']
                    if length == 2:
                        skew_score += skew_weights['2连斜'] * 3
                    elif length == 3:
                        skew_score += skew_weights['3连斜'] * 4
                    elif length == 4:
                        skew_score += skew_weights['4连斜'] * 5
                    elif length >= 5:
                        skew_score += skew_weights['5连斜以上'] * 6
            
            # 检查反斜连
            for skew_key, skew_info in skew_analysis.get('反斜连', {}).items():
                if skew_info['结束号码'] == num:
                    skew_score += skew_weights['反斜连'] * 4
            
            # 连续出现得分 (权重5%)
            consecutive_score = 0
            if num in consecutive_analysis.get('当前连续', {}):
                consecutive_length = consecutive_analysis['当前连续'][num]
                if consecutive_length == 2:
                    consecutive_score = skew_weights['连续出现2期'] * 2
                elif consecutive_length == 3:
                    consecutive_score = skew_weights['连续出现3期'] * 3
                elif consecutive_length >= 4:
                    consecutive_score = skew_weights['连续出现4期以上'] * 4
            
            # 热号追踪得分
            if num in consecutive_analysis.get('热号追踪', {}):
                hot_info = consecutive_analysis['热号追踪'][num]
                if hot_info['最近出现'] == '是':
                    consecutive_score += 2
            
            return freq_score + missing_score + odd_even_score + size_score + skew_score + consecutive_score
        
        def score_back(num):
            """后区号码综合评分"""
            # 频率得分 (权重40%)
            max_back_freq = max(back_freq.values()) if back_freq else 1
            freq_score = (back_freq.get(num, 0) / max_back_freq * 40)
            
            # 遗漏得分 (权重35%)
            avg_back_missing = np.mean(list(back_missing.values())) if back_missing else 1
            missing_norm = back_missing.get(num, 0) / avg_back_missing if avg_back_missing > 0 else 1
            
            if missing_norm < 0.5:
                missing_score = 20
            elif missing_norm < 1.0:
                missing_score = 30
            elif missing_norm < 2.0:
                missing_score = 35
            else:
                missing_score = 25
            
            # 奇偶平衡 (权重25%)
            odd_even_score = 12.5
            
            return freq_score + missing_score + odd_even_score
        
        # 为所有号码评分
        front_scores = {num: score_front(num) for num in range(1, 36)}
        back_scores = {num: score_back(num) for num in range(1, 13)}
        
        # 生成推荐组合
        recommendations = []
        selection_log = []
        
        def is_valid_combination(front_nums, back_nums):
            """检查号码组合是否合理"""
            if len(front_nums) != 5 or len(back_nums) != 2:
                return False
            
            if len(set(front_nums)) != 5 or len(set(back_nums)) != 2:
                return False
            
            # 奇偶比
            odd_count = sum(1 for num in front_nums if num % 2 == 1)
            if not (2 <= odd_count <= 3):
                return False
            
            # 大小比
            big_count = sum(1 for num in front_nums if num > 18)
            if not (2 <= big_count <= 3):
                return False
            
            # 跨度
            span = max(front_nums) - min(front_nums)
            if span < 15 or span > 30:
                return False
            
            # 和值
            total = sum(front_nums)
            if total < 80 or total > 130:
                return False
            
            return True
        
        # 生成5组推荐
        for group_num in range(5):
            attempts = 0
            max_attempts = 1000
            
            while attempts < max_attempts:
                attempts += 1
                
                # 使用加权随机选择前区号码
                front_candidates = list(front_scores.items())
                front_weights = [score for _, score in front_candidates]
                total_weight = sum(front_weights)
                
                if total_weight == 0:
                    probabilities = None
                else:
                    probabilities = [w/total_weight for w in front_weights]
                
                front_nums = []
                while len(front_nums) < 5:
                    if probabilities:
                        chosen_idx = np.random.choice(len(front_candidates), p=probabilities)
                    else:
                        chosen_idx = np.random.choice(len(front_candidates))
                    chosen_num = front_candidates[chosen_idx][0]
                    
                    if chosen_num not in front_nums:
                        front_nums.append(chosen_num)
                
                front_nums.sort()
                
                # 选择后区号码
                back_candidates = list(back_scores.items())
                back_weights = [score for _, score in back_candidates]
                back_total_weight = sum(back_weights)
                
                if back_total_weight == 0:
                    back_probabilities = None
                else:
                    back_probabilities = [w/back_total_weight for w in back_weights]
                
                back_nums = []
                while len(back_nums) < 2:
                    if back_probabilities:
                        chosen_idx = np.random.choice(len(back_candidates), p=back_probabilities)
                    else:
                        chosen_idx = np.random.choice(len(back_candidates))
                    chosen_num = back_candidates[chosen_idx][0]
                    
                    if chosen_num not in back_nums:
                        back_nums.append(chosen_num)
                
                back_nums.sort()
                
                # 验证组合
                if is_valid_combination(front_nums, back_nums):
                    # 计算组合评分
                    front_score = sum(front_scores[num] for num in front_nums)
                    back_score = sum(back_scores[num] for num in back_nums)
                    total_score = front_score + back_score
                    
                    # 分析组合特点
                    odd_count = sum(1 for num in front_nums if num % 2 == 1)
                    big_count = sum(1 for num in front_nums if num > 18)
                    
                    # 检查斜连号和连续出现情况
                    skew_features = []
                    for num in front_nums:
                        # 检查是否是斜连号
                        for prediction in skew_analysis.get('预测号码', []):
                            if prediction['号码'] == num:
                                skew_features.append(f"{num}({prediction['类型']})")
                        
                        # 检查是否连续出现
                        if num in consecutive_analysis.get('当前连续', {}):
                            length = consecutive_analysis['当前连续'][num]
                            skew_features.append(f"{num}(连续{length}期)")
                    
                    # 记录选号依据
                    selection_reason = {
                        '组号': group_num + 1,
                        '前区': front_nums,
                        '后区': back_nums,
                        '总分': round(total_score, 2),
                        '选号依据': [
                            f"前区包含热门号码: {[num for num in front_nums if num in dict(front_freq.most_common(10))]}",
                            f"前区考虑遗漏: {[f'{num}({front_missing.get(num, 0)}期)' for num in front_nums]}",
                            f"奇偶比: {odd_count}:{5-odd_count}",
                            f"大小比: {big_count}:{5-big_count}",
                            f"和值: {sum(front_nums)}",
                            f"跨度: {max(front_nums) - min(front_nums)}",
                            f"斜连/连续特征: {', '.join(skew_features) if skew_features else '无'}"
                        ]
                    }
                    
                    recommendations.append({
                        '组号': group_num + 1,
                        '前区': front_nums,
                        '后区': back_nums,
                        '总分': total_score
                    })
                    
                    selection_log.append(selection_reason)
                    break
        
        return recommendations, selection_log, front_scores, back_scores
    
    def generate_hot_selected_recommendations(self):
        """生成热号精选推荐（从各周期热号中精选）"""
        print("\n" + "=" * 80)
        print("正在生成热号精选推荐...")
        print("=" * 80)
        
        # 收集各周期的热号
        period_hot_numbers = {
            '最近30期': [],
            '最近60期': [],
            '最近100期': []
        }
        
        period_hot_back_numbers = {
            '最近30期': [],
            '最近60期': [],
            '最近100期': []
        }
        
        # 获取各周期的热号（前区前3名，后区前2名）
        for period_name in period_hot_numbers.keys():
            if period_name in self.analysis_results:
                analysis = self.analysis_results[period_name]
                
                # 前区热号（取前3名）
                front_hot = list(analysis['热门号码']['前区热门'].keys())[:3]
                period_hot_numbers[period_name] = front_hot
                
                # 后区热号（取前2名）
                back_hot = list(analysis['热门号码']['后区热门'].keys())[:2]
                period_hot_back_numbers[period_name] = back_hot
        
        # 显示各周期热号
        print("\n各周期热号统计：")
        for period_name in period_hot_numbers.keys():
            front_hot = period_hot_numbers[period_name]
            back_hot = period_hot_back_numbers[period_name]
            print(f"{period_name}: 前区热号{front_hot} | 后区热号{back_hot}")
        
        # 合并所有热号
        all_hot_numbers = []
        for numbers in period_hot_numbers.values():
            all_hot_numbers.extend(numbers)
        
        all_hot_back_numbers = []
        for numbers in period_hot_back_numbers.values():
            all_hot_back_numbers.extend(numbers)
        
        # 统计热号出现次数（在几个周期中是热号）
        hot_counter = Counter(all_hot_numbers)
        back_hot_counter = Counter(all_hot_back_numbers)
        
        print(f"\n热号合并结果：")
        print(f"前区热号（去重后）: {sorted(set(all_hot_numbers))}")
        print(f"后区热号（去重后）: {sorted(set(all_hot_back_numbers))}")
        print(f"热号出现次数统计: {dict(hot_counter)}")
        print(f"后区热号出现次数统计: {dict(back_hot_counter)}")
        
        # 生成热号精选推荐
        hot_selected_recs = []
        
        # 策略1：选择出现次数最多的热号
        if hot_counter:
            # 前区选择出现次数最多的5个号码
            most_common_front = [num for num, _ in hot_counter.most_common(5)]
            
            # 如果不足5个，从其他热号中补充
            if len(most_common_front) < 5:
                remaining_hot = [num for num in sorted(set(all_hot_numbers)) if num not in most_common_front]
                most_common_front.extend(remaining_hot[:5-len(most_common_front)])
            
            # 后区选择出现次数最多的2个号码
            most_common_back = [num for num, _ in back_hot_counter.most_common(2)]
            
            # 如果不足2个，从其他后区热号中补充
            if len(most_common_back) < 2:
                remaining_back = [num for num in sorted(set(all_hot_back_numbers)) if num not in most_common_back]
                most_common_back.extend(remaining_back[:2-len(most_common_back)])
            
            # 验证组合
            if self.validate_combination(most_common_front, most_common_back):
                hot_selected_recs.append({
                    '组号': 1,
                    '前区': sorted(most_common_front),
                    '后区': sorted(most_common_back),
                    '策略': '热号最多出现次数',
                    '特点': [
                        f"热号出现次数: {[f'{num}({hot_counter[num]}次)' for num in most_common_front]}",
                        f"后区热号出现次数: {[f'{num}({back_hot_counter[num]}次)' for num in most_common_back] if most_common_back else '无'}",
                        f"奇偶比: {sum(1 for num in most_common_front if num % 2 == 1)}:{5-sum(1 for num in most_common_front if num % 2 == 1)}",
                        f"大小比: {sum(1 for num in most_common_front if num > 18)}:{5-sum(1 for num in most_common_front if num > 18)}"
                    ]
                })
        
        # 策略2：加权随机选择（权重为出现次数）
        for group_num in range(2, 6):  # 生成第2-5组
            attempts = 0
            max_attempts = 500
            
            while attempts < max_attempts:
                attempts += 1
                
                # 前区加权随机选择
                front_candidates = list(hot_counter.items())
                if not front_candidates:
                    break
                
                front_weights = [count for _, count in front_candidates]
                total_weight = sum(front_weights)
                
                if total_weight == 0:
                    probabilities = [1/len(front_candidates)] * len(front_candidates)
                else:
                    probabilities = [w/total_weight for w in front_weights]
                
                front_nums = []
                while len(front_nums) < 5:
                    chosen_idx = np.random.choice(len(front_candidates), p=probabilities)
                    chosen_num = front_candidates[chosen_idx][0]
                    
                    if chosen_num not in front_nums:
                        front_nums.append(chosen_num)
                
                front_nums.sort()
                
                # 后区加权随机选择
                back_candidates = list(back_hot_counter.items())
                if not back_candidates:
                    # 如果没有后区热号，从1-12中随机选择
                    back_nums = list(np.random.choice(range(1, 13), size=2, replace=False))
                else:
                    back_weights = [count for _, count in back_candidates]
                    back_total_weight = sum(back_weights)
                    
                    if back_total_weight == 0:
                        back_probabilities = [1/len(back_candidates)] * len(back_candidates)
                    else:
                        back_probabilities = [w/back_total_weight for w in back_weights]
                    
                    back_nums = []
                    while len(back_nums) < 2:
                        chosen_idx = np.random.choice(len(back_candidates), p=back_probabilities)
                        chosen_num = back_candidates[chosen_idx][0]
                        
                        if chosen_num not in back_nums:
                            back_nums.append(chosen_num)
                
                back_nums.sort()
                
                # 验证组合
                if self.validate_combination(front_nums, back_nums):
                    # 计算热号得分（基于出现次数）
                    hot_score = sum(hot_counter.get(num, 0) for num in front_nums)
                    back_hot_score = sum(back_hot_counter.get(num, 0) for num in back_nums)
                    total_hot_score = hot_score + back_hot_score
                    
                    hot_selected_recs.append({
                        '组号': group_num,
                        '前区': front_nums,
                        '后区': back_nums,
                        '策略': '热号加权随机',
                        '热号得分': total_hot_score,
                        '特点': [
                            f"热号出现次数: {[f'{num}({hot_counter[num]}次)' for num in front_nums if num in hot_counter]}",
                            f"热号占比: {len([num for num in front_nums if num in hot_counter])}/5",
                            f"奇偶比: {sum(1 for num in front_nums if num % 2 == 1)}:{5-sum(1 for num in front_nums if num % 2 == 1)}",
                            f"大小比: {sum(1 for num in front_nums if num > 18)}:{5-sum(1 for num in front_nums if num > 18)}"
                        ]
                    })
                    break
        
        self.hot_selected_recommendations = hot_selected_recs
        
        print("\n" + "=" * 80)
        print("热号精选推荐结果")
        print("=" * 80)
        
        for rec in hot_selected_recs:
            front_str = ' '.join(f"{num:2d}" for num in rec['前区'])
            back_str = ' '.join(f"{num:2d}" for num in rec['后区'])
            print(f"第{rec['组号']}组: 前区 [{front_str}] | 后区 [{back_str}]")
            print(f"      策略: {rec['策略']}")
            for feature in rec['特点'][:2]:
                print(f"      {feature}")
            print()
        
        return hot_selected_recs
    
    def validate_combination(self, front_nums, back_nums):
        """验证号码组合是否合理"""
        if len(front_nums) != 5 or len(back_nums) != 2:
            return False
        
        # 检查重复
        if len(set(front_nums)) != 5 or len(set(back_nums)) != 2:
            return False
        
        # 检查号码范围
        if not all(1 <= num <= 35 for num in front_nums):
            return False
        if not all(1 <= num <= 12 for num in back_nums):
            return False
        
        # 奇偶比 (2:3或3:2)
        odd_count = sum(1 for num in front_nums if num % 2 == 1)
        if not (2 <= odd_count <= 3):
            return False
        
        # 大小比 (2:3或3:2)
        big_count = sum(1 for num in front_nums if num > 18)
        if not (2 <= big_count <= 3):
            return False
        
        # 跨度
        span = max(front_nums) - min(front_nums)
        if span < 15 or span > 30:
            return False
        
        # 和值
        total = sum(front_nums)
        if total < 80 or total > 130:
            return False
        
        # 尾数分布
        tails = [num % 10 for num in front_nums]
        if len(set(tails)) < 3:
            return False
        
        # 后区避免相邻
        if abs(back_nums[1] - back_nums[0]) == 1:
            return False
        
        return True
    
    def perform_complete_analysis(self):
        """执行完整分析"""
        print("=" * 80)
        print("大乐透智能分析系统")
        print("=" * 80)
        print(f"分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 存储所有分析结果
        all_recommendations = {}
        all_selection_logs = {}
        
        # 分析不同周期
        periods = [
            ('最近30期', self.recent_30),
            ('最近60期', self.recent_60),
            ('最近100期', self.recent_100),
            ('全部数据', self.all_data)
        ]
        
        for period_name, period_data in periods:
            if len(period_data) == 0:
                print(f"✗ {period_name}数据为空，跳过分析")
                continue
            
            print(f"\n正在分析{period_name}...")
            
            # 分析该周期数据
            analysis, front_freq, back_freq, front_missing, back_missing = self.analyze_period(
                period_data, period_name
            )
            
            # 获取斜连号和连续出现分析
            skew_analysis = analysis.get('斜连号分析', {})
            consecutive_analysis = analysis.get('连续出现分析', {})
            
            # 生成推荐
            recommendations, selection_log, front_scores, back_scores = self.generate_recommendations_for_period(
                period_data, period_name, front_freq, back_freq, front_missing, back_missing,
                skew_analysis, consecutive_analysis
            )
            
            # 存储结果
            self.analysis_results[period_name] = analysis
            all_recommendations[period_name] = recommendations
            all_selection_logs[period_name] = selection_log
            
            print(f"✓ {period_name}分析完成，生成{len(recommendations)}组推荐")
            
            # 显示斜连号分析结果
            if skew_analysis:
                print(f"  - 发现{len(skew_analysis.get('正斜连', {}))}个正斜连模式")
                print(f"  - 发现{len(skew_analysis.get('反斜连', {}))}个反斜连模式")
                print(f"  - 预测{len(skew_analysis.get('预测号码', []))}个斜连号")
            
            # 显示连续出现分析结果
            if consecutive_analysis.get('当前连续'):
                print(f"  - {len(consecutive_analysis['当前连续'])}个号码当前连续出现")
        
        self.recommendations = all_recommendations
        self.selection_logs = all_selection_logs
        
        # 生成综合推荐
        self.generate_comprehensive_recommendations()
        
        # 生成热号精选推荐
        self.generate_hot_selected_recommendations()
        
        return all_recommendations
    
    def generate_comprehensive_recommendations(self):
        """综合最近30、60、100期推荐，生成5组最优号码"""
        print("\n" + "=" * 80)
        print("正在生成综合推荐号码...")
        print("=" * 80)
        
        # 收集所有周期的推荐
        all_recommendations = []
        
        # 从各周期取前3名（权重：30期 > 60期 > 100期）
        for period_name, weight in [('最近30期', 3.0), ('最近60期', 2.0), ('最近100期', 1.0)]:
            if period_name in self.recommendations:
                recs = self.recommendations[period_name]
                # 取前3名并添加权重
                for i, rec in enumerate(recs[:3]):
                    weighted_rec = rec.copy()
                    weighted_rec['总分'] = rec['总分'] * weight
                    weighted_rec['来源'] = f"{period_name}第{rec['组号']}组"
                    weighted_rec['原始分'] = rec['总分']
                    weighted_rec['权重'] = weight
                    all_recommendations.append(weighted_rec)
        
        if not all_recommendations:
            print("✗ 没有找到任何周期的推荐数据")
            return []
        
        # 按加权总分排序
        all_recommendations.sort(key=lambda x: x['总分'], reverse=True)
        
        # 提取所有号码的频率
        front_counter = Counter()
        back_counter = Counter()
        
        for rec in all_recommendations:
            for num in rec['前区']:
                front_counter[num] += rec['总分']
            for num in rec['后区']:
                back_counter[num] += rec['总分']
        
        # 基于加权频率重新生成5组最优组合
        comprehensive_recs = []
        
        def get_weighted_random(weights_dict):
            """加权随机选择"""
            items = list(weights_dict.items())
            total_weight = sum(weights_dict.values())
            if total_weight == 0:
                return np.random.choice([item[0] for item in items])
            
            rand_val = np.random.random() * total_weight
            cumulative = 0
            for num, weight in items:
                cumulative += weight
                if rand_val <= cumulative:
                    return num
            return items[-1][0]
        
        for group_num in range(5):
            attempts = 0
            max_attempts = 500
            
            while attempts < max_attempts:
                attempts += 1
                
                # 选择前区号码（使用分层加权随机）
                front_nums = []
                front_weights = front_counter.copy()
                
                while len(front_nums) < 5:
                    if not front_weights:
                        # 如果没有权重，随机选择
                        available_nums = [num for num in range(1, 36) if num not in front_nums]
                        if not available_nums:
                            break
                        chosen = np.random.choice(available_nums)
                    else:
                        # 移除已选号码
                        for num in front_nums:
                            if num in front_weights:
                                del front_weights[num]
                        
                        if not front_weights:
                            available_nums = [num for num in range(1, 36) if num not in front_nums]
                            if not available_nums:
                                break
                            chosen = np.random.choice(available_nums)
                        else:
                            chosen = get_weighted_random(front_weights)
                    
                    if chosen not in front_nums:
                        front_nums.append(chosen)
                
                front_nums.sort()
                
                # 选择后区号码
                back_nums = []
                back_weights = back_counter.copy()
                
                while len(back_nums) < 2:
                    if not back_weights:
                        available_back = [num for num in range(1, 13) if num not in back_nums]
                        if not available_back:
                            break
                        chosen = np.random.choice(available_back)
                    else:
                        # 移除已选号码
                        for num in back_nums:
                            if num in back_weights:
                                del back_weights[num]
                        
                        if not back_weights:
                            available_back = [num for num in range(1, 13) if num not in back_nums]
                            if not available_back:
                                break
                            chosen = np.random.choice(available_back)
                        else:
                            chosen = get_weighted_random(back_weights)
                    
                    if chosen not in back_nums:
                        back_nums.append(chosen)
                
                back_nums.sort()
                
                # 验证组合
                if self.validate_combination(front_nums, back_nums):
                    # 计算组合得分（基于号码在原始推荐中的加权频率）
                    front_score = sum(front_counter.get(num, 0) for num in front_nums)
                    back_score = sum(back_counter.get(num, 0) for num in back_nums)
                    total_score = front_score + back_score
                    
                    # 分析组合特点
                    odd_count = sum(1 for num in front_nums if num % 2 == 1)
                    big_count = sum(1 for num in front_nums if num > 18)
                    
                    # 找出号码的来源
                    source_counts = Counter()
                    for rec in all_recommendations:
                        front_match = len(set(rec['前区']) & set(front_nums))
                        back_match = len(set(rec['后区']) & set(back_nums))
                        if front_match >= 2 or back_match >= 1:
                            source_counts[rec['来源']] += 1
                    
                    main_source = source_counts.most_common(1)[0][0] if source_counts else "综合生成"
                    
                    comprehensive_recs.append({
                        '组号': group_num + 1,
                        '前区': front_nums,
                        '后区': back_nums,
                        '总分': round(total_score, 2),
                        '特点': [
                            f"奇偶比: {odd_count}:{5-odd_count}",
                            f"大小比: {big_count}:{5-big_count}",
                            f"和值: {sum(front_nums)}",
                            f"跨度: {max(front_nums) - min(front_nums)}",
                            f"热门号码: {[num for num in front_nums if num in front_counter and front_counter[num] > np.percentile(list(front_counter.values()), 70)]}",
                            f"来源参考: {main_source}"
                        ]
                    })
                    break
        
        # 按总分排序
        comprehensive_recs.sort(key=lambda x: x['总分'], reverse=True)
        
        self.comprehensive_recommendations = comprehensive_recs[:5]  # 取前5组
        
        print("✓ 综合推荐生成完成")
        print(f"✓ 生成{len(self.comprehensive_recommendations)}组综合推荐号码")
        
        # 打印推荐结果
        print("\n" + "=" * 80)
        print("综合推荐结果（基于最近30、60、100期数据）")
        print("=" * 80)
        
        for rec in self.comprehensive_recommendations:
            front_str = ' '.join(f"{num:2d}" for num in rec['前区'])
            back_str = ' '.join(f"{num:2d}" for num in rec['后区'])
            print(f"第{rec['组号']}组: 前区 [{front_str}] | 后区 [{back_str}] | 综合评分: {rec['总分']:.1f}")
            for feature in rec['特点'][:3]:  # 只显示前3个特点
                print(f"       {feature}")
            print()
        
        return self.comprehensive_recommendations
    
    def generate_analysis_document(self):
        """生成分析文档"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        doc_filename = f"大乐透分析报告_{timestamp}.md"
        
        print(f"\n正在生成分析文档: {doc_filename}")
        
        with open(doc_filename, 'w', encoding='utf-8') as f:
            # 标题
            f.write("# 大乐透智能分析报告\n\n")
            f.write(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            
            # 基本信息
            f.write("## 一、基本信息\n")
            f.write(f"- **分析数据总期数**: {self.analysis_results['基本信息']['总期数']}期\n")
            f.write(f"- **数据范围**: 期号{self.analysis_results['基本信息']['最早期号']} - {self.analysis_results['基本信息']['最晚期号']}\n")
            f.write(f"- **时间范围**: {self.analysis_results['基本信息']['最早日期']} - {self.analysis_results['基本信息']['最晚日期']}\n")
            f.write(f"- **前区范围**: {self.analysis_results['基本信息']['前区范围']}\n")
            f.write(f"- **后区范围**: {self.analysis_results['基本信息']['后区范围']}\n\n")
            
            # 各周期分析结果
            for period_name in ['最近30期', '最近60期', '最近100期']:
                if period_name not in self.analysis_results:
                    continue
                    
                analysis = self.analysis_results[period_name]
                
                f.write(f"## 二、{period_name}分析结果\n")
                f.write(f"- **分析期数**: {analysis['数据期数']}期\n")
                f.write(f"- **期号范围**: {analysis['开始期号']} - {analysis['结束期号']}\n\n")
                
                # 热门号码
                f.write(f"### 1. {period_name}热门号码\n")
                f.write("**前区热门号码(TOP10)**:\n")
                for num, count in analysis['热门号码']['前区热门'].items():
                    percentage = count / analysis['数据期数'] * 100
                    f.write(f"  - 号码{num:2d}: 出现{count:3d}次 ({percentage:5.1f}%)\n")
                
                f.write("\n**后区热门号码(TOP5)**:\n")
                for num, count in analysis['热门号码']['后区热门'].items():
                    percentage = count / analysis['数据期数'] * 100
                    f.write(f"  - 号码{num:2d}: 出现{count:3d}次 ({percentage:5.1f}%)\n")
                
                # 模式分析
                f.write(f"\n### 2. {period_name}模式分析\n")
                stats = analysis['模式分析']
                
                if stats['奇偶比']:
                    odd_ratio_dist = Counter(stats['奇偶比'])
                    f.write("**奇偶比分布**:\n")
                    for ratio, count in odd_ratio_dist.most_common():
                        percentage = count / len(stats['奇偶比']) * 100
                        f.write(f"  - {ratio}: {count:3d}次 ({percentage:5.1f}%)\n")
                
                if stats['大小比']:
                    size_ratio_dist = Counter(stats['大小比'])
                    f.write("\n**大小比分布**:\n")
                    for ratio, count in size_ratio_dist.most_common():
                        percentage = count / len(stats['大小比']) * 100
                        f.write(f"  - {ratio}: {count:3d}次 ({percentage:5.1f}%)\n")
                
                if stats['和值']:
                    f.write(f"\n**和值统计**:\n")
                    f.write(f"  - 平均和值: {np.mean(stats['和值']):.1f}\n")
                    f.write(f"  - 最小和值: {min(stats['和值'])}\n")
                    f.write(f"  - 最大和值: {max(stats['和值'])}\n")
                
                f.write(f"\n**出现连号的期数**: {stats['连号次数']}次 ({stats['连号次数']/analysis['数据期数']*100:.1f}%)\n")
                
                # 斜连号分析
                f.write(f"\n### 3. {period_name}斜连号分析\n")
                skew_analysis = analysis.get('斜连号分析', {})
                
                if skew_analysis.get('正斜连'):
                    f.write("**正斜连模式**:\n")
                    for key, info in skew_analysis['正斜连'].items():
                        f.write(f"  - {key}: {info['长度']}期斜连\n")
                
                if skew_analysis.get('反斜连'):
                    f.write("\n**反斜连模式**:\n")
                    for key, info in skew_analysis['反斜连'].items():
                        f.write(f"  - {key}: {info['长度']}期反斜连\n")
                
                if skew_analysis.get('预测号码'):
                    f.write("\n**斜连号预测**:\n")
                    for pred in skew_analysis['预测号码']:
                        f.write(f"  - 号码{pred['号码']:2d}: {pred['类型']} ({pred['理由']})\n")
                
                # 连续出现分析
                f.write(f"\n### 4. {period_name}连续出现分析\n")
                consecutive_analysis = analysis.get('连续出现分析', {})
                
                if consecutive_analysis.get('当前连续'):
                    f.write("**当前连续出现的号码**:\n")
                    for num, length in consecutive_analysis['当前连续'].items():
                        f.write(f"  - 号码{num:2d}: 连续出现{length}期\n")
                
                f.write("\n")
            
            # 综合推荐
            f.write("## 三、综合推荐结果\n")
            f.write("基于最近30期、60期、100期的分析结果，综合加权推荐以下5组号码：\n\n")
            
            if self.comprehensive_recommendations:
                f.write("| 组号 | 前区号码 | 后区号码 | 综合评分 | 主要特点 |\n")
                f.write("|------|----------|----------|----------|----------|\n")
                
                for rec in self.comprehensive_recommendations:
                    front_str = ' '.join(f"{num:2d}" for num in rec['前区'])
                    back_str = ' '.join(f"{num:2d}" for num in rec['后区'])
                    f.write(f"| {rec['组号']} | {front_str} | {back_str} | {rec['总分']:.1f} | {rec['特点'][0].replace('奇偶比: ', '')} |\n")
                
                f.write("\n### 详细选号依据：\n")
                for rec in self.comprehensive_recommendations:
                    f.write(f"\n**第{rec['组号']}组**:\n")
                    for feature in rec['特点']:
                        f.write(f"  - {feature}\n")
            
            # 热号精选推荐
            f.write("\n## 四、热号精选推荐\n")
            f.write("基于最近30期、60期、100期的热门号码（各取前3名）精选生成的推荐：\n\n")
            
            if self.hot_selected_recommendations:
                f.write("| 组号 | 前区号码 | 后区号码 | 选择策略 | 主要特点 |\n")
                f.write("|------|----------|----------|----------|----------|\n")
                
                for rec in self.hot_selected_recommendations:
                    front_str = ' '.join(f"{num:2d}" for num in rec['前区'])
                    back_str = ' '.join(f"{num:2d}" for num in rec['后区'])
                    f.write(f"| {rec['组号']} | {front_str} | {back_str} | {rec['策略']} | {rec['特点'][0] if rec['特点'] else '热号精选'} |\n")
                
                f.write("\n### 热号来源说明：\n")
                f.write("1. **最近30期热号**：")
                if '最近30期' in self.analysis_results:
                    hot_30 = list(self.analysis_results['最近30期']['热门号码']['前区热门'].keys())[:3]
                    f.write(f"{hot_30}\n")
                
                f.write("2. **最近60期热号**：")
                if '最近60期' in self.analysis_results:
                    hot_60 = list(self.analysis_results['最近60期']['热门号码']['前区热门'].keys())[:3]
                    f.write(f"{hot_60}\n")
                
                f.write("3. **最近100期热号**：")
                if '最近100期' in self.analysis_results:
                    hot_100 = list(self.analysis_results['最近100期']['热门号码']['前区热门'].keys())[:3]
                    f.write(f"{hot_100}\n")
            
            # 分析方法和策略说明
            f.write("\n## 五、分析方法和策略说明\n")
            f.write("""
### 1. 分析维度
**基础分析**：
- 频率分析：统计各号码出现次数
- 遗漏分析：分析号码未出现的期数
- 模式分析：奇偶比、大小比、和值、跨度

**高级趋势分析**：
- **斜连号分析**：
  - 正斜连：号码在连续几期中递增或递减（相差1）
  - 反斜连：号码先增后减或先减后增形成V型
  - 权重设置：2连斜(1.2)、3连斜(1.5)、4连斜(2.0)、5连斜以上(2.5)

- **连续出现分析**：
  - 跟踪号码连续出现的次数
  - 权重设置：连续2期(1.2)、连续3期(1.5)、连续4期以上(2.0)

### 2. 推荐策略
**综合推荐策略**：
- **多周期加权融合**：
  - 最近30期：权重3.0（最重视近期趋势）
  - 最近60期：权重2.0（考虑中期趋势）
  - 最近100期：权重1.0（参考长期趋势）

**热号精选策略**：
- **热号筛选**：从各周期热号中精选前3名
- **热号合并**：合并各周期热号，统计出现次数
- **策略选择**：
  1. 热号最多出现次数：选择出现次数最多的号码
  2. 热号加权随机：基于出现次数的加权随机选择

### 3. 组合验证标准
**前区验证**：
- 奇偶比：2:3或3:2（平衡）
- 大小比：2:3或3:2（平衡）
- 和值范围：80-130（常见范围）
- 跨度范围：15-30（合理分布）
- 尾数分布：至少3个不同尾数

**后区验证**：
- 避免相邻号码
- 奇偶搭配
- 大小搭配

### 4. 使用建议
1. **策略选择**：
   - 综合推荐：适合追求平衡和趋势跟踪
   - 热号精选：适合追热号和短期趋势
2. **分散投注**：可以选择不同策略的组合进行分散投注
3. **长期跟踪**：建议跟踪3-5期的开奖结果，验证分析效果
4. **理性购彩**：彩票是概率游戏，请保持理性，小额投注
""")
            
            # 风险提示
            f.write("\n## 六、风险提示\n")
            f.write("""
1. **彩票的随机性**：彩票开奖结果是完全随机的，历史数据不能预测未来结果
2. **分析仅供参考**：本分析基于历史数据统计，仅供娱乐参考
3. **理性投注**：建议每次投注金额不超过个人承受能力
4. **禁止沉迷**：彩票不是投资工具，请勿沉迷
5. **概率极低**：大乐透一等奖概率约为2142万分之一，中奖机会极低

**祝您好运，理性购彩！**
""")
        
        print(f"✓ 分析文档已生成: {doc_filename}")
        return doc_filename
    
    def generate_excel_recommendations(self):
        """生成Excel推荐文件"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"大乐透推荐号码_{timestamp}.xlsx"
        
        print(f"\n正在生成Excel文件: {excel_filename}")
        
        # 创建Excel工作簿
        wb = Workbook()
        
        # 创建样式
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        center_alignment = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        
        # 1. 综合推荐工作表
        ws_comprehensive = wb.active
        ws_comprehensive.title = "综合推荐"
        
        # 设置列宽
        for col in range(1, 15):
            ws_comprehensive.column_dimensions[get_column_letter(col)].width = 12
        
        # 添加标题
        ws_comprehensive.merge_cells('A1:N1')
        title_cell = ws_comprehensive['A1']
        title_cell.value = "大乐透综合推荐号码（基于最近30、60、100期）"
        title_cell.font = Font(bold=True, size=14, color="366092")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 添加副标题
        ws_comprehensive.merge_cells('A2:N2')
        subtitle_cell = ws_comprehensive['A2']
        subtitle_cell.value = f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 综合5组最优推荐"
        subtitle_cell.font = Font(size=10, italic=True)
        subtitle_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 添加表头
        headers = [
            '组号', '前区1', '前区2', '前区3', '前区4', '前区5', 
            '后区1', '后区2', '综合评分',
            '奇偶比', '大小比', '和值', '跨度', '选号要点'
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws_comprehensive.cell(row=3, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # 添加综合推荐数据
        if self.comprehensive_recommendations:
            for row_idx, rec in enumerate(self.comprehensive_recommendations, 4):
                # 组号
                ws_comprehensive.cell(row=row_idx, column=1, value=rec['组号']).alignment = center_alignment
                
                # 前区号码
                for i, num in enumerate(rec['前区'], 2):
                    cell = ws_comprehensive.cell(row=row_idx, column=i, value=num)
                    cell.alignment = center_alignment
                    cell.border = thin_border
                
                # 后区号码
                for i, num in enumerate(rec['后区'], 7):
                    cell = ws_comprehensive.cell(row=row_idx, column=i, value=num)
                    cell.alignment = center_alignment
                    cell.border = thin_border
                
                # 综合评分
                ws_comprehensive.cell(row=row_idx, column=9, value=round(rec['总分'], 1)).alignment = center_alignment
                
                # 计算其他指标
                front_nums = rec['前区']
                
                # 奇偶比
                odd_count = sum(1 for num in front_nums if num % 2 == 1)
                ws_comprehensive.cell(row=row_idx, column=10, value=f"{odd_count}:{5-odd_count}").alignment = center_alignment
                
                # 大小比
                big_count = sum(1 for num in front_nums if num > 18)
                ws_comprehensive.cell(row=row_idx, column=11, value=f"{big_count}:{5-big_count}").alignment = center_alignment
                
                # 和值
                total_sum = sum(front_nums)
                ws_comprehensive.cell(row=row_idx, column=12, value=total_sum).alignment = center_alignment
                
                # 跨度
                span = max(front_nums) - min(front_nums)
                ws_comprehensive.cell(row=row_idx, column=13, value=span).alignment = center_alignment
                
                # 选号要点
                key_points = []
                if rec['特点']:
                    # 取前两个特点
                    for feature in rec['特点'][:2]:
                        if "奇偶比" in feature:
                            key_points.append(feature.replace("奇偶比: ", "奇偶"))
                        elif "大小比" in feature:
                            key_points.append(feature.replace("大小比: ", "大小"))
                        elif "热门号码" in feature:
                            key_points.append("含热门")
                
                ws_comprehensive.cell(row=row_idx, column=14, value=" | ".join(key_points) if key_points else "平衡组合").alignment = center_alignment
                
                # 添加边框
                for col in range(1, 15):
                    ws_comprehensive.cell(row=row_idx, column=col).border = thin_border
        
        # 2. 热号精选推荐工作表
        ws_hot_selected = wb.create_sheet(title="热号精选推荐")
        
        # 设置列宽
        for col in range(1, 15):
            ws_hot_selected.column_dimensions[get_column_letter(col)].width = 12
        
        # 添加标题
        ws_hot_selected.merge_cells('A1:N1')
        title_cell = ws_hot_selected['A1']
        title_cell.value = "大乐透热号精选推荐（基于各周期热号）"
        title_cell.font = Font(bold=True, size=14, color="366092")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 添加副标题
        ws_hot_selected.merge_cells('A2:N2')
        subtitle_cell = ws_hot_selected['A2']
        subtitle_cell.value = f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 从各周期热号中精选生成"
        subtitle_cell.font = Font(size=10, italic=True)
        subtitle_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 添加表头
        hot_headers = [
            '组号', '前区1', '前区2', '前区3', '前区4', '前区5', 
            '后区1', '后区2', '选择策略',
            '奇偶比', '大小比', '和值', '跨度', '热号特点'
        ]
        
        for col, header in enumerate(hot_headers, 1):
            cell = ws_hot_selected.cell(row=3, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # 添加热号精选推荐数据
        if self.hot_selected_recommendations:
            for row_idx, rec in enumerate(self.hot_selected_recommendations, 4):
                # 组号
                ws_hot_selected.cell(row=row_idx, column=1, value=rec['组号']).alignment = center_alignment
                
                # 前区号码
                for i, num in enumerate(rec['前区'], 2):
                    cell = ws_hot_selected.cell(row=row_idx, column=i, value=num)
                    cell.alignment = center_alignment
                    cell.border = thin_border
                    # 如果是热号，标记颜色
                    cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                
                # 后区号码
                for i, num in enumerate(rec['后区'], 7):
                    cell = ws_hot_selected.cell(row=row_idx, column=i, value=num)
                    cell.alignment = center_alignment
                    cell.border = thin_border
                
                # 选择策略
                ws_hot_selected.cell(row=row_idx, column=9, value=rec['策略']).alignment = center_alignment
                
                # 计算其他指标
                front_nums = rec['前区']
                
                # 奇偶比
                odd_count = sum(1 for num in front_nums if num % 2 == 1)
                ws_hot_selected.cell(row=row_idx, column=10, value=f"{odd_count}:{5-odd_count}").alignment = center_alignment
                
                # 大小比
                big_count = sum(1 for num in front_nums if num > 18)
                ws_hot_selected.cell(row=row_idx, column=11, value=f"{big_count}:{5-big_count}").alignment = center_alignment
                
                # 和值
                total_sum = sum(front_nums)
                ws_hot_selected.cell(row=row_idx, column=12, value=total_sum).alignment = center_alignment
                
                # 跨度
                span = max(front_nums) - min(front_nums)
                ws_hot_selected.cell(row=row_idx, column=13, value=span).alignment = center_alignment
                
                # 热号特点
                hot_features = []
                if rec['特点']:
                    for feature in rec['特点'][:2]:
                        if "热号" in feature:
                            hot_features.append(feature)
                
                ws_hot_selected.cell(row=row_idx, column=14, value=" | ".join(hot_features[:2]) if hot_features else "热号精选").alignment = center_alignment
                
                # 添加边框
                for col in range(1, 15):
                    ws_hot_selected.cell(row=row_idx, column=col).border = thin_border
        
        # 3. 各周期推荐工作表
        for period_name in ['最近30期', '最近60期', '最近100期']:
            if period_name not in self.recommendations:
                continue
            
            # 创建工作表
            ws = wb.create_sheet(title=period_name)
            
            # 设置列宽
            for col in range(1, 15):
                ws.column_dimensions[get_column_letter(col)].width = 12
            
            # 添加标题
            ws.merge_cells('A1:N1')
            title_cell = ws['A1']
            title_cell.value = f"大乐透{period_name}推荐"
            title_cell.font = Font(bold=True, size=14, color="366092")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # 添加表头
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_alignment
                cell.border = thin_border
            
            # 添加数据
            recommendations = self.recommendations[period_name]
            
            for row_idx, rec in enumerate(recommendations, 4):
                # 组号
                ws.cell(row=row_idx, column=1, value=rec['组号']).alignment = center_alignment
                
                # 前区号码
                for i, num in enumerate(rec['前区'], 2):
                    cell = ws.cell(row=row_idx, column=i, value=num)
                    cell.alignment = center_alignment
                    cell.border = thin_border
                
                # 后区号码
                for i, num in enumerate(rec['后区'], 7):
                    cell = ws.cell(row=row_idx, column=i, value=num)
                    cell.alignment = center_alignment
                    cell.border = thin_border
                
                # 评分
                ws.cell(row=row_idx, column=9, value=round(rec['总分'], 1)).alignment = center_alignment
                
                # 计算其他指标
                front_nums = rec['前区']
                
                # 奇偶比
                odd_count = sum(1 for num in front_nums if num % 2 == 1)
                ws.cell(row=row_idx, column=10, value=f"{odd_count}:{5-odd_count}").alignment = center_alignment
                
                # 大小比
                big_count = sum(1 for num in front_nums if num > 18)
                ws.cell(row=row_idx, column=11, value=f"{big_count}:{5-big_count}").alignment = center_alignment
                
                # 和值
                total_sum = sum(front_nums)
                ws.cell(row=row_idx, column=12, value=total_sum).alignment = center_alignment
                
                # 跨度
                span = max(front_nums) - min(front_nums)
                ws.cell(row=row_idx, column=13, value=span).alignment = center_alignment
                
                # 选号要点
                key_points = []
                if odd_count == 2 or odd_count == 3:
                    key_points.append("奇偶平衡")
                if big_count == 2 or big_count == 3:
                    key_points.append("大小平衡")
                if 90 <= total_sum <= 120:
                    key_points.append("和值佳")
                
                ws.cell(row=row_idx, column=14, value=" | ".join(key_points) if key_points else "标准组合").alignment = center_alignment
                
                # 添加边框
                for col in range(1, 15):
                    ws.cell(row=row_idx, column=col).border = thin_border
        
        # 4. 分析摘要工作表
        summary_ws = wb.create_sheet(title="分析摘要")
        
        # 设置列宽
        for col in range(1, 10):
            summary_ws.column_dimensions[get_column_letter(col)].width = 20
        
        # 添加标题
        summary_ws.merge_cells('A1:I1')
        title_cell = summary_ws['A1']
        title_cell.value = "大乐透分析摘要"
        title_cell.font = Font(bold=True, size=14, color="366092")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 添加摘要表头
        summary_headers = ['项目', '数值', '说明']
        
        for col, header in enumerate(summary_headers, 1):
            cell = summary_ws.cell(row=2, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # 添加摘要数据
        summary_data = [
            ("总分析期数", self.analysis_results['基本信息']['总期数'], "历史开奖总期数"),
            ("最晚开奖期", self.analysis_results['基本信息']['最晚期号'], "最新一期开奖期号"),
            ("综合推荐组数", len(self.comprehensive_recommendations), "基于多周期分析的推荐"),
            ("热号精选组数", len(self.hot_selected_recommendations), "基于各周期热号的精选"),
            ("分析周期", "30/60/100期", "分析的三个时间窗口"),
            ("权重分配", "3.0/2.0/1.0", "30期>60期>100期"),
            ("斜连号权重", "2连斜1.2,3连斜1.5,4连斜2.0", "斜连号分析权重设置"),
            ("连续出现权重", "2期1.2,3期1.5,4期以上2.0", "连续出现分析权重设置"),
            ("前区范围", "1-35", "前区号码范围"),
            ("后区范围", "1-12", "后区号码范围"),
            ("奇偶标准", "2:3或3:2", "推荐的奇偶比例"),
            ("大小标准", "2:3或3:2", "推荐的大小比例"),
            ("和值范围", "80-130", "推荐的和值范围")
        ]
        
        for i, (item, value, desc) in enumerate(summary_data, 3):
            summary_ws.cell(row=i, column=1, value=item).border = thin_border
            summary_ws.cell(row=i, column=2, value=value).border = thin_border
            summary_ws.cell(row=i, column=3, value=desc).border = thin_border
        
        # 5. 使用说明工作表
        instruction_ws = wb.create_sheet(title="使用说明")
        
        # 设置列宽
        instruction_ws.column_dimensions['A'].width = 50
        instruction_ws.column_dimensions['B'].width = 60
        
        # 添加标题
        instruction_ws.merge_cells('A1:B1')
        title_cell = instruction_ws['A1']
        title_cell.value = "大乐透分析系统使用说明"
        title_cell.font = Font(bold=True, size=14, color="366092")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        instructions = [
            ("工作表说明", "1. 综合推荐：基于最近30、60、100期数据的5组最优推荐\n2. 热号精选推荐：基于各周期热号精选生成的推荐\n3. 最近30期：基于最近30期数据的推荐\n4. 最近60期：基于最近60期数据的推荐\n5. 最近100期：基于最近100期数据的推荐\n6. 分析摘要：分析的基本信息和参数\n7. 使用说明：本文件的使用说明"),
            ("推荐说明", "1. 综合推荐：多维度综合评分，适合追求平衡\n2. 热号精选推荐：专注热号趋势，适合追热号策略\n3. 每组号码都经过合理性验证\n4. 评分越高表示综合表现越好"),
            ("分析维度", "1. 基础分析：频率、遗漏、模式分析\n2. 斜连号分析：正斜连、反斜连模式识别\n3. 连续出现分析：跟踪号码连续出现情况\n4. 多周期综合：整合30/60/100期分析结果"),
            ("选号建议", "1. 建议选择2-3组进行投注\n2. 可以结合个人偏好调整\n3. 长期跟踪效果更佳\n4. 可以尝试不同策略的组合"),
            ("风险提示", "1. 彩票具有随机性，分析仅供参考\n2. 理性购彩，量力而行\n3. 禁止沉迷，保持健康心态")
        ]
        
        row_idx = 3
        for title, content in instructions:
            title_cell = instruction_ws.cell(row=row_idx, column=1, value=title)
            title_cell.font = Font(bold=True, color="366092")
            row_idx += 1
            
            # 处理多行内容
            lines = content.split('\n')
            for line in lines:
                instruction_ws.cell(row=row_idx, column=2, value=line)
                row_idx += 1
            
            row_idx += 1  # 空一行
        
        # 保存Excel文件
        wb.save(excel_filename)
        print(f"✓ Excel文件已生成: {excel_filename}")
        
        return excel_filename
    
    def run_full_analysis(self):
        """运行完整分析流程"""
        print("=" * 80)
        print("大乐透智能分析系统 - 开始运行")
        print("=" * 80)
        
        try:
            # 1. 执行完整分析
            recommendations = self.perform_complete_analysis()
            
            # 2. 生成分析文档
            doc_file = self.generate_analysis_document()
            
            # 3. 生成Excel推荐文件
            excel_file = self.generate_excel_recommendations()
            
            print("\n" + "=" * 80)
            print("分析完成！")
            print("=" * 80)
            print(f"✓ 分析文档: {doc_file}")
            print(f"✓ 推荐文件: {excel_file}")
            print("\n文件说明:")
            print("1. 分析文档(.md): 包含详细的分析过程、选号依据和策略说明")
            print("2. 推荐文件(.xlsx): 包含综合推荐、热号精选推荐、各周期推荐和分析摘要")
            print("\n推荐策略说明:")
            print("• 综合推荐：多维度综合评分，基于30/60/100期数据加权融合")
            print("• 热号精选推荐：从各周期热号中精选生成，专注热号趋势")
            print("• 各周期推荐：分别基于30期、60期、100期数据的推荐")
            print("\n温馨提示:")
            print("• 彩票具有随机性，分析结果仅供参考")
            print("• 建议理性购彩，小额投注")
            print("• 可以尝试不同策略的组合进行分散投注")
            print("• 祝您好运！")
            print("=" * 80)
            
            return doc_file, excel_file
            
        except Exception as e:
            print(f"分析过程出错: {e}")
            import traceback
            traceback.print_exc()
            return None, None


# 主程序
if __name__ == "__main__":
    # 设置文件路径
    file_path = "/home/acuto/python_stu/data_analysis/大乐透开奖数据.xlsx"
    
    print("大乐透智能分析系统")
    print("=" * 80)
    
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"✗ 错误: 找不到文件 {file_path}")
        print("请确保文件路径正确！")
        exit(1)
    
    try:
        # 创建分析器
        print("正在初始化分析器...")
        analyzer = LotteryAnalyzer(file_path)
        
        # 运行完整分析
        analyzer.run_full_analysis()
        
    except Exception as e:
        print(f"程序运行出错: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n按Enter键退出...")
    input()