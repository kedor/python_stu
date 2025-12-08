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
import warnings
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
        
        # 时间窗口
        self.recommend_30 = []  # 30期推荐
        self.recommend_60 = []  # 60期推荐
        self.recommend_100 = []  # 100期推荐
        self.comprehensive_recommendations = []  # 综合推荐结果
        self.cold_hot_transition = {}  # 冷热转换分析
        self.hot_analysis_30 = {}  # 30期热号分析
        self.hot_analysis_60 = {}  # 60期热号分析
        self.hot_analysis_100 = {}  # 100期热号分析
        
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
            self.recent_30 = self.data.tail(30).copy() if len(self.data) >= 30 else self.data.copy()
            self.recent_60 = self.data.tail(60).copy() if len(self.data) >= 60 else self.data.copy()
            self.recent_100 = self.data.tail(100).copy() if len(self.data) >= 100 else self.data.copy()
            self.all_data = self.data.copy()
            
            print(f"✓ 数据划分: 最近30期({len(self.recent_30)})，最近60期({len(self.recent_60)})，最近100期({len(self.recent_100)})")
            
        except Exception as e:
            print(f"✗ 数据加载失败: {e}")
            import traceback
            traceback.print_exc()
            raise
    
    def analyze_hot_numbers(self, period_data, period_name, top_n=15):
        """分析热号"""
        print(f"正在分析{period_name}热号情况...")
        
        if len(period_data) < 10:
            print(f"  ✗ {period_name}数据不足，跳过热号分析")
            return {}
        
        sorted_data = period_data.sort_values('期号')
        
        hot_analysis = {
            '前区热号': {},
            '后区热号': {},
            '分析期数': len(sorted_data),
            '数据范围': f"{sorted_data['期号'].min()}-{sorted_data['期号'].max()}",
            '分析时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        # 收集所有号码
        all_front = []
        all_back = []
        
        for _, row in sorted_data.iterrows():
            all_front.extend(row['前区'])
            all_back.extend(row['后区'])
        
        # 频率统计
        front_freq = Counter(all_front)
        back_freq = Counter(all_back)
        
        # 计算理论频率
        total_periods = len(sorted_data)
        
        # 前区热号（取前top_n个）
        hot_front = front_freq.most_common(top_n)
        for num, count in hot_front:
            percentage = count / total_periods * 100
            
            recent_appearances = 0
            # 检查最近10期出现次数（如果数据足够）
            if total_periods >= 10:
                recent_data = sorted_data.tail(10)
                for _, row in recent_data.iterrows():
                    if num in row['前区']:
                        recent_appearances += 1
            
            hot_analysis['前区热号'][num] = {
                '总次数': count,
                '总频率': f"{percentage:.2f}%",
                '近期次数': recent_appearances,
                '近期频率': f"{recent_appearances/10*100:.2f}%" if total_periods >= 10 else "不足10期",
                '热力等级': self.get_heat_level(count, total_periods, recent_appearances),
                '遗漏期数': self.calculate_missing(num, sorted_data)
            }
        
        # 后区热号（取前8个）
        hot_back = back_freq.most_common(8)
        for num, count in hot_back:
            percentage = count / total_periods * 100
            hot_analysis['后区热号'][num] = {
                '总次数': count,
                '总频率': f"{percentage:.2f}%",
                '遗漏期数': self.calculate_missing_back(num, sorted_data)
            }
        
        return hot_analysis
    
    def calculate_missing(self, num, sorted_data):
        """计算前区遗漏期数"""
        missing_count = 0
        for i in range(len(sorted_data)-1, -1, -1):
            if num in sorted_data.iloc[i]['前区']:
                break
            missing_count += 1
        return missing_count
    
    def calculate_missing_back(self, num, sorted_data):
        """计算后区遗漏期数"""
        missing_count = 0
        for i in range(len(sorted_data)-1, -1, -1):
            if num in sorted_data.iloc[i]['后区']:
                break
            missing_count += 1
        return missing_count
    
    def get_heat_level(self, total_count, total_periods, recent_count):
        """判断热力等级"""
        avg_frequency = total_count / total_periods if total_periods > 0 else 0
        recent_frequency = recent_count / 10 if recent_count > 0 else 0
        
        # 基于理论频率的倍数判断（前区理论频率5/35=0.1429）
        theoretical_freq = 0.1429
        multiple = avg_frequency / theoretical_freq
        
        if multiple > 1.5 and recent_frequency > 0.4:
            return "🔥🔥🔥 超热"
        elif multiple > 1.3 and recent_frequency > 0.3:
            return "🔥🔥 高热"
        elif multiple > 1.1 or recent_frequency > 0.2:
            return "🔥 温热"
        elif multiple > 0.9:
            return "↔️ 正常"
        elif multiple > 0.7:
            return "❄️ 偏冷"
        else:
            return "❄️❄️ 冷门"
    
    def analyze_cold_hot_transition(self):
        """分析冷热转换趋势"""
        print("\n" + "=" * 80)
        print("正在分析冷热转换趋势...")
        print("=" * 80)
        
        if len(self.data) < 60:
            print("✗ 数据不足60期，无法进行冷热转换分析")
            return {}
        
        # 划分时间窗口
        sorted_data = self.data.sort_values('期号')
        
        # 窗口1: 早期30期
        early_window = sorted_data.iloc[-60:-30] if len(sorted_data) >= 60 else sorted_data.iloc[:30]
        
        # 窗口2: 近期30期
        late_window = sorted_data.tail(30)
        
        # 窗口3: 最近10期
        recent_window = sorted_data.tail(10)
        
        print(f"窗口划分: 早期{len(early_window)}期，近期{len(late_window)}期，最新{len(recent_window)}期")
        
        # 计算各个窗口的号码频率
        early_counts = Counter()
        late_counts = Counter()
        recent_counts = Counter()
        
        for _, row in early_window.iterrows():
            early_counts.update(row['前区'])
        
        for _, row in late_window.iterrows():
            late_counts.update(row['前区'])
        
        for _, row in recent_window.iterrows():
            recent_counts.update(row['前区'])
        
        # 分析冷热转换
        transition_analysis = {
            '冷转热': [],  # 由冷变热
            '热转冷': [],  # 由热变冷
            '持续热': [],  # 持续热门
            '持续冷': [],  # 持续冷门
            '反弹信号': [],  # 有反弹信号的冷号
            '预警信号': [],   # 可能转冷的预警
            '分析参数': {
                '早期窗口': f"{len(early_window)}期",
                '近期窗口': f"{len(late_window)}期",
                '最新窗口': f"{len(recent_window)}期",
                '分析时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
        }
        
        for num in range(1, 36):
            early_freq = early_counts.get(num, 0) / len(early_window) if len(early_window) > 0 else 0
            late_freq = late_counts.get(num, 0) / len(late_window) if len(late_window) > 0 else 0
            recent_freq = recent_counts.get(num, 0) / len(recent_window) if len(recent_window) > 0 else 0
            
            # 计算趋势得分
            trend_score = late_freq - early_freq
            
            # 计算当前遗漏
            current_missing = 0
            for i in range(len(sorted_data)-1, -1, -1):
                if num in sorted_data.iloc[i]['前区']:
                    break
                current_missing += 1
            
            # 判断冷热状态
            if early_freq < 0.12 and late_freq > 0.2:
                transition_analysis['冷转热'].append({
                    '号码': num,
                    '早期频率': f"{early_freq*100:.2f}%",
                    '近期频率': f"{late_freq*100:.2f}%",
                    '趋势强度': trend_score,
                    '趋势评级': self.get_trend_rating(trend_score),
                    '最新出现': '是' if recent_counts.get(num, 0) > 0 else '否',
                    '遗漏期数': current_missing
                })
            
            elif early_freq > 0.2 and late_freq < 0.12:
                transition_analysis['热转冷'].append({
                    '号码': num,
                    '早期频率': f"{early_freq*100:.2f}%",
                    '近期频率': f"{late_freq*100:.2f}%",
                    '趋势强度': trend_score,
                    '趋势评级': self.get_trend_rating(trend_score),
                    '最新出现': '是' if recent_counts.get(num, 0) > 0 else '否',
                    '遗漏期数': current_missing
                })
            
            elif early_freq > 0.18 and late_freq > 0.18:
                transition_analysis['持续热'].append({
                    '号码': num,
                    '早期频率': f"{early_freq*100:.2f}%",
                    '近期频率': f"{late_freq*100:.2f}%",
                    '稳定性': '高',
                    '最新出现': '是' if recent_counts.get(num, 0) > 0 else '否',
                    '遗漏期数': current_missing
                })
            
            elif early_freq < 0.1 and late_freq < 0.1 and current_missing > 10:
                transition_analysis['持续冷'].append({
                    '号码': num,
                    '早期频率': f"{early_freq*100:.2f}%",
                    '近期频率': f"{late_freq*100:.2f}%",
                    '稳定性': '高',
                    '最新出现': '是' if recent_counts.get(num, 0) > 0 else '否',
                    '遗漏期数': current_missing
                })
            
            # 检测反弹信号（长期冷门但有近期出现）
            if early_freq < 0.1 and recent_counts.get(num, 0) > 0 and current_missing <= 8:
                rebound_score = self.calculate_rebound_score(num, sorted_data)
                transition_analysis['反弹信号'].append({
                    '号码': num,
                    '早期频率': f"{early_freq*100:.2f}%",
                    '近期出现': '是',
                    '遗漏期数': current_missing,
                    '反弹概率': f"{rebound_score:.1f}%",
                    '热度变化': '由冷转温'
                })
            
            # 检测预警信号（热门但近期降温）
            if early_freq > 0.22 and recent_counts.get(num, 0) == 0 and current_missing >= 6:
                transition_analysis['预警信号'].append({
                    '号码': num,
                    '早期频率': f"{early_freq*100:.2f}%",
                    '近期出现': '否',
                    '遗漏期数': current_missing,
                    '预警级别': '高' if current_missing >= 12 else '中',
                    '建议': '谨慎选择'
                })
        
        # 对各分类排序
        transition_analysis['冷转热'].sort(key=lambda x: x['趋势强度'], reverse=True)
        transition_analysis['热转冷'].sort(key=lambda x: x['趋势强度'])
        transition_analysis['反弹信号'].sort(key=lambda x: float(x['反弹概率'].replace('%', '')), reverse=True)
        
        self.cold_hot_transition = transition_analysis
        
        # 打印分析结果
        print(f"冷转热号码: {len(transition_analysis['冷转热'])}个")
        print(f"热转冷号码: {len(transition_analysis['热转冷'])}个")
        print(f"持续热号码: {len(transition_analysis['持续热'])}个")
        print(f"持续冷号码: {len(transition_analysis['持续冷'])}个")
        print(f"反弹信号: {len(transition_analysis['反弹信号'])}个")
        print(f"预警信号: {len(transition_analysis['预警信号'])}个")
        
        return transition_analysis
    
    def get_trend_rating(self, trend_score):
        """获取趋势评级"""
        if trend_score > 0.15:
            return "🔥🔥🔥 强势转热"
        elif trend_score > 0.1:
            return "🔥🔥 明显转热"
        elif trend_score > 0.05:
            return "🔥 温和转热"
        elif trend_score < -0.15:
            return "❄️❄️❄️ 强势转冷"
        elif trend_score < -0.1:
            return "❄️❄️ 明显转冷"
        elif trend_score < -0.05:
            return "❄️ 温和转冷"
        else:
            return "↔️ 稳定"
    
    def calculate_rebound_score(self, num, sorted_data):
        """计算反弹概率"""
        # 获取号码的历史出现模式
        appearances = []
        for i in range(len(sorted_data)):
            if num in sorted_data.iloc[i]['前区']:
                appearances.append(i)
        
        if len(appearances) < 3:
            return 35.0
        
        # 计算平均间隔
        intervals = []
        for i in range(1, len(appearances)):
            intervals.append(appearances[i] - appearances[i-1])
        
        avg_interval = np.mean(intervals) if intervals else 18
        
        # 计算当前遗漏
        current_missing = 0
        for i in range(len(sorted_data)-1, -1, -1):
            if num in sorted_data.iloc[i]['前区']:
                break
            current_missing += 1
        
        # 计算反弹概率
        if avg_interval > 0:
            rebound_prob = min(95, (current_missing / avg_interval) * 60 + 35)
        else:
            rebound_prob = 50.0
        
        return rebound_prob
    
    def generate_recommendations_for_period(self, period_data, period_name, group_count=5):
        """为指定周期生成推荐"""
        print(f"\n正在生成{period_name}推荐...")
        
        # 分析该周期的热号
        if period_name == "30期":
            if not self.hot_analysis_30:
                self.hot_analysis_30 = self.analyze_hot_numbers(period_data, period_name)
            hot_analysis = self.hot_analysis_30
        elif period_name == "60期":
            if not self.hot_analysis_60:
                self.hot_analysis_60 = self.analyze_hot_numbers(period_data, period_name)
            hot_analysis = self.hot_analysis_60
        elif period_name == "100期":
            if not self.hot_analysis_100:
                self.hot_analysis_100 = self.analyze_hot_numbers(period_data, period_name)
            hot_analysis = self.hot_analysis_100
        else:
            hot_analysis = self.analyze_hot_numbers(period_data, period_name)
        
        if not hot_analysis.get('前区热号'):
            print(f"✗ {period_name}没有找到足够的热号数据")
            return []
        
        # 使用冷热转换分析
        if not self.cold_hot_transition:
            self.analyze_cold_hot_transition()
        transition_analysis = self.cold_hot_transition
        
        # 提取前区热号
        front_hot_numbers = list(hot_analysis['前区热号'].keys())
        front_hot_scores = {}
        
        for num in front_hot_numbers:
            info = hot_analysis['前区热号'][num]
            # 基础热号评分
            total_score = float(info['总频率'].replace('%', ''))
            recent_score = float(info['近期频率'].replace('%', '')) * 1.5 if info['近期频率'] != "不足10期" else total_score
            heat_bonus = 0
            
            if "超热" in info['热力等级']:
                heat_bonus = 30
            elif "高热" in info['热力等级']:
                heat_bonus = 20
            elif "温热" in info['热力等级']:
                heat_bonus = 10
            
            base_score = total_score + recent_score + heat_bonus
            
            # 冷热转换调整
            transition_bonus = 0
            
            # 检查是否是冷转热号码
            cold_to_hot_items = [item for item in transition_analysis.get('冷转热', []) if item['号码'] == num]
            if cold_to_hot_items:
                transition_info = cold_to_hot_items[0]
                if "强势转热" in transition_info['趋势评级']:
                    transition_bonus = 50
                elif "明显转热" in transition_info['趋势评级']:
                    transition_bonus = 35
                elif "温和转热" in transition_info['趋势评级']:
                    transition_bonus = 20
            
            # 检查是否有反弹信号
            rebound_items = [item for item in transition_analysis.get('反弹信号', []) if item['号码'] == num]
            if rebound_items:
                rebound_prob = float(rebound_items[0]['反弹概率'].replace('%', ''))
                transition_bonus += rebound_prob * 0.4
            
            # 检查是否是持续热号
            continuous_hot = [item for item in transition_analysis.get('持续热', []) if item['号码'] == num]
            if continuous_hot:
                transition_bonus += 15
            
            # 检查是否是热转冷号码（减分）
            hot_to_cold = [item for item in transition_analysis.get('热转冷', []) if item['号码'] == num]
            if hot_to_cold:
                transition_info = hot_to_cold[0]
                if "强势转冷" in transition_info['趋势评级']:
                    transition_bonus -= 40
                elif "明显转冷" in transition_info['趋势评级']:
                    transition_bonus -= 25
                elif "温和转冷" in transition_info['趋势评级']:
                    transition_bonus -= 10
            
            # 检查是否有预警信号
            warning_items = [item for item in transition_analysis.get('预警信号', []) if item['号码'] == num]
            if warning_items:
                if warning_items[0]['预警级别'] == '高':
                    transition_bonus -= 30
                else:
                    transition_bonus -= 15
            
            front_hot_scores[num] = base_score + transition_bonus
        
        # 提取后区热号
        back_hot_numbers = list(hot_analysis.get('后区热号', {}).keys())
        back_hot_scores = {}
        
        for num in back_hot_numbers:
            info = hot_analysis['后区热号'][num]
            back_hot_scores[num] = float(info['总频率'].replace('%', ''))
        
        # 生成推荐组
        recommendations = []
        
        for group_num in range(group_count):
            attempts = 0
            max_attempts = 300
            
            while attempts < max_attempts:
                attempts += 1
                
                # 从热号中选择前区号码
                front_candidates = list(front_hot_scores.items())
                front_weights = [max(score, 1) for _, score in front_candidates]
                total_weight = sum(front_weights)
                
                front_nums = []
                while len(front_nums) < 5:
                    probabilities = [w/total_weight for w in front_weights]
                    chosen_idx = np.random.choice(len(front_candidates), p=probabilities)
                    chosen_num = front_candidates[chosen_idx][0]
                    
                    if chosen_num not in front_nums:
                        front_nums.append(chosen_num)
                
                front_nums.sort()
                
                # 从热号中选择后区号码
                back_candidates = list(back_hot_scores.items())
                back_weights = [score for _, score in back_candidates]
                back_total_weight = sum(back_weights)
                
                back_nums = []
                while len(back_nums) < 2:
                    if back_total_weight == 0:
                        back_probabilities = None
                    else:
                        back_probabilities = [w/back_total_weight for w in back_weights]
                    
                    if back_probabilities:
                        chosen_idx = np.random.choice(len(back_candidates), p=back_probabilities)
                    else:
                        chosen_idx = np.random.choice(len(back_candidates))
                    
                    chosen_num = back_candidates[chosen_idx][0]
                    
                    if chosen_num not in back_nums:
                        back_nums.append(chosen_num)
                
                back_nums.sort()
                
                # 验证组合合理性
                if self.validate_combination(front_nums, back_nums, hot_analysis, transition_analysis):
                    # 计算组合评分
                    front_score = sum(front_hot_scores.get(num, 0) for num in front_nums)
                    back_score = sum(back_hot_scores.get(num, 0) for num in back_nums)
                    total_score = front_score + back_score
                    
                    # 分析特点
                    heat_levels = []
                    transition_features = []
                    
                    for num in front_nums:
                        if num in hot_analysis['前区热号']:
                            heat_info = hot_analysis['前区热号'][num]
                            heat_levels.append(f"{num}({heat_info['热力等级']})")
                        
                        cold_to_hot = [item for item in transition_analysis.get('冷转热', []) if item['号码'] == num]
                        if cold_to_hot:
                            trans_info = cold_to_hot[0]
                            transition_features.append(f"{num}({trans_info['趋势评级']})")
                    
                    # 统计冷热转换特征
                    cold_to_hot_count = sum(1 for num in front_nums 
                                          if any(item['号码'] == num for item in transition_analysis.get('冷转热', [])))
                    
                    recommendations.append({
                        '组号': group_num + 1,
                        '前区': front_nums,
                        '后区': back_nums,
                        '总分': round(total_score, 2),
                        '冷转热数量': cold_to_hot_count,
                        '奇偶比': f"{sum(1 for num in front_nums if num % 2 == 1)}:{5-sum(1 for num in front_nums if num % 2 == 1)}",
                        '大小比': f"{sum(1 for num in front_nums if num > 18)}:{5-sum(1 for num in front_nums if num > 18)}",
                        '和值': sum(front_nums),
                        '热号特点': heat_levels[:3] if heat_levels else [],
                        '转换特征': transition_features[:2] if transition_features else ["均衡组合"]
                    })
                    break
        
        # 按总分排序
        recommendations.sort(key=lambda x: x['总分'], reverse=True)
        
        print(f"✓ {period_name}推荐生成完成: {len(recommendations)}组")
        return recommendations[:group_count]
    
    def validate_combination(self, front_nums, back_nums, hot_analysis, transition_analysis):
        """验证组合是否合理"""
        if len(front_nums) != 5 or len(back_nums) != 2:
            return False
        
        # 检查重复
        if len(set(front_nums)) != 5 or len(set(back_nums)) != 2:
            return False
        
        # 检查至少包含2个热号
        hot_count = sum(1 for num in front_nums if num in hot_analysis.get('前区热号', {}))
        if hot_count < 2:
            return False
        
        # 奇偶比（2:3或3:2）
        odd_count = sum(1 for num in front_nums if num % 2 == 1)
        if not (2 <= odd_count <= 3):
            return False
        
        # 大小比（以18为界）
        big_count = sum(1 for num in front_nums if num > 18)
        if not (2 <= big_count <= 3):
            return False
        
        # 和值范围（70-140）
        total = sum(front_nums)
        if total < 70 or total > 140:
            return False
        
        # 连号检查（不超过2个连续号码）
        front_sorted = sorted(front_nums)
        consecutive_count = 0
        for i in range(1, len(front_sorted)):
            if front_sorted[i] - front_sorted[i-1] == 1:
                consecutive_count += 1
        if consecutive_count > 2:
            return False
        
        return True
    
    def generate_comprehensive_recommendations(self):
        """生成综合推荐（基于三个时间段的推荐）"""
        print("\n" + "=" * 80)
        print("正在生成综合推荐（整合三个时间段）...")
        print("=" * 80)
        
        # 确保有三个时间段的推荐
        if not self.recommend_30:
            self.recommend_30 = self.generate_recommendations_for_period(self.recent_30, "30期")
        if not self.recommend_60:
            self.recommend_60 = self.generate_recommendations_for_period(self.recent_60, "60期")
        if not self.recommend_100:
            self.recommend_100 = self.generate_recommendations_for_period(self.recent_100, "100期")
        
        # 收集所有推荐中的号码
        all_front_numbers = []
        all_back_numbers = []
        
        for rec in self.recommend_30 + self.recommend_60 + self.recommend_100:
            all_front_numbers.extend(rec['前区'])
            all_back_numbers.extend(rec['后区'])
        
        # 计算号码出现频率
        front_freq = Counter(all_front_numbers)
        back_freq = Counter(all_back_numbers)
        
        # 生成5组综合推荐
        comprehensive_recommendations = []
        
        for group_num in range(5):
            attempts = 0
            max_attempts = 200
            
            while attempts < max_attempts:
                attempts += 1
                
                # 选择前区号码（基于频率加权）
                front_candidates = list(front_freq.items())
                front_weights = [count for _, count in front_candidates]
                total_weight = sum(front_weights)
                
                front_nums = []
                while len(front_nums) < 5:
                    probabilities = [w/total_weight for w in front_weights]
                    chosen_idx = np.random.choice(len(front_candidates), p=probabilities)
                    chosen_num = front_candidates[chosen_idx][0]
                    
                    if chosen_num not in front_nums:
                        front_nums.append(chosen_num)
                
                front_nums.sort()
                
                # 选择后区号码
                back_candidates = list(back_freq.items())
                back_weights = [count for _, count in back_candidates]
                back_total_weight = sum(back_weights)
                
                back_nums = []
                while len(back_nums) < 2:
                    probabilities = [w/back_total_weight for w in back_weights] if back_total_weight > 0 else None
                    if probabilities:
                        chosen_idx = np.random.choice(len(back_candidates), p=probabilities)
                    else:
                        chosen_idx = np.random.choice(len(back_candidates))
                    
                    chosen_num = back_candidates[chosen_idx][0]
                    
                    if chosen_num not in back_nums:
                        back_nums.append(chosen_num)
                
                back_nums.sort()
                
                # 验证组合
                if self.validate_combination(front_nums, back_nums, {}, {}):
                    # 计算综合评分（基于在三个推荐中的出现次数）
                    front_score = sum(front_freq.get(num, 0) for num in front_nums)
                    back_score = sum(back_freq.get(num, 0) for num in back_nums)
                    total_score = front_score + back_score
                    
                    # 分析来源
                    sources_30 = sum(1 for rec in self.recommend_30 if any(num in rec['前区'] for num in front_nums))
                    sources_60 = sum(1 for rec in self.recommend_60 if any(num in rec['前区'] for num in front_nums))
                    sources_100 = sum(1 for rec in self.recommend_100 if any(num in rec['前区'] for num in front_nums))
                    
                    comprehensive_recommendations.append({
                        '组号': group_num + 1,
                        '前区': front_nums,
                        '后区': back_nums,
                        '综合评分': total_score,
                        '来源统计': f"30期:{sources_30} 60期:{sources_60} 100期:{sources_100}",
                        '奇偶比': f"{sum(1 for num in front_nums if num % 2 == 1)}:{5-sum(1 for num in front_nums if num % 2 == 1)}",
                        '大小比': f"{sum(1 for num in front_nums if num > 18)}:{5-sum(1 for num in front_nums if num > 18)}",
                        '和值': sum(front_nums)
                    })
                    break
        
        # 按评分排序
        comprehensive_recommendations.sort(key=lambda x: x['综合评分'], reverse=True)
        self.comprehensive_recommendations = comprehensive_recommendations
        
        print("✓ 综合推荐生成完成")
        return comprehensive_recommendations

    def generate_excel_recommendations(self):
        """生成Excel推荐文件（包含四个时间段的推荐和详细分析）"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"大乐透推荐_{timestamp}.xlsx"
        
        print(f"\n正在生成Excel文件: {excel_filename}")
        
        # 创建Excel工作簿
        wb = Workbook()
        
        # 创建样式
        header_font = Font(bold=True, color="FFFFFF")
        header_fill_30 = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")  # 蓝色
        header_fill_60 = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")  # 橙色
        header_fill_100 = PatternFill(start_color="A5A5A5", end_color="A5A5A5", fill_type="solid")  # 灰色
        header_fill_comprehensive = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")  # 绿色
        header_fill_cold_hot = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # 红色
        header_fill_hot_analysis = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")  # 紫色
        center_alignment = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        
        # 1. 30期推荐工作表
        ws_30 = wb.active
        ws_30.title = "30期推荐"
        
        # 设置列宽
        column_widths = [8, 25, 15, 15, 15, 15, 15, 20, 25]
        for col, width in enumerate(column_widths, 1):
            ws_30.column_dimensions[get_column_letter(col)].width = width
        
        # 添加标题
        ws_30.merge_cells('A1:I1')
        title_cell = ws_30['A1']
        title_cell.value = "大乐透30期推荐（短期趋势）"
        title_cell.font = Font(bold=True, size=14, color="4472C4")
        title_cell.alignment = center_alignment
        
        # 添加副标题
        ws_30.merge_cells('A2:I2')
        subtitle_cell = ws_30['A2']
        subtitle_cell.value = f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 基于最近30期热号分析+冷热转换"
        subtitle_cell.font = Font(size=10, italic=True)
        subtitle_cell.alignment = center_alignment
        
        # 添加表头
        headers = ['组号', '前区号码', '后区号码', '综合评分', '冷转热数', '奇偶比', '大小比', '和值', '热号特点']
        
        for col, header in enumerate(headers, 1):
            cell = ws_30.cell(row=3, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill_30
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # 添加30期推荐数据
        if self.recommend_30:
            for row_idx, rec in enumerate(self.recommend_30, 4):
                # 组号
                ws_30.cell(row=row_idx, column=1, value=rec['组号']).alignment = center_alignment
                
                # 前区号码
                front_nums_str = ' '.join(f"{num:02d}" for num in rec['前区'])
                ws_30.cell(row=row_idx, column=2, value=front_nums_str).alignment = center_alignment
                
                # 后区号码
                back_nums_str = ' '.join(f"{num:02d}" for num in rec['后区'])
                ws_30.cell(row=row_idx, column=3, value=back_nums_str).alignment = center_alignment
                
                # 综合评分
                ws_30.cell(row=row_idx, column=4, value=round(rec['总分'], 1)).alignment = center_alignment
                
                # 冷转热数量
                ws_30.cell(row=row_idx, column=5, value=rec['冷转热数量']).alignment = center_alignment
                
                # 奇偶比
                ws_30.cell(row=row_idx, column=6, value=rec['奇偶比']).alignment = center_alignment
                
                # 大小比
                ws_30.cell(row=row_idx, column=7, value=rec['大小比']).alignment = center_alignment
                
                # 和值
                ws_30.cell(row=row_idx, column=8, value=rec['和值']).alignment = center_alignment
                
                # 热号特点
                heat_chars = ', '.join(rec['热号特点']) if rec['热号特点'] else "均衡组合"
                ws_30.cell(row=row_idx, column=9, value=heat_chars).alignment = center_alignment
                
                # 添加边框
                for col in range(1, 10):
                    ws_30.cell(row=row_idx, column=col).border = thin_border
        
        # 2. 60期推荐工作表
        ws_60 = wb.create_sheet(title="60期推荐")
        
        # 设置列宽
        for col, width in enumerate(column_widths, 1):
            ws_60.column_dimensions[get_column_letter(col)].width = width
        
        # 添加标题
        ws_60.merge_cells('A1:I1')
        title_cell = ws_60['A1']
        title_cell.value = "大乐透60期推荐（中期趋势）"
        title_cell.font = Font(bold=True, size=14, color="ED7D31")
        title_cell.alignment = center_alignment
        
        # 添加副标题
        ws_60.merge_cells('A2:I2')
        subtitle_cell = ws_60['A2']
        subtitle_cell.value = f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 基于最近60期热号分析+冷热转换"
        subtitle_cell.font = Font(size=10, italic=True)
        subtitle_cell.alignment = center_alignment
        
        # 添加表头
        for col, header in enumerate(headers, 1):
            cell = ws_60.cell(row=3, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill_60
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # 添加60期推荐数据
        if self.recommend_60:
            for row_idx, rec in enumerate(self.recommend_60, 4):
                ws_60.cell(row=row_idx, column=1, value=rec['组号']).alignment = center_alignment
                front_nums_str = ' '.join(f"{num:02d}" for num in rec['前区'])
                ws_60.cell(row=row_idx, column=2, value=front_nums_str).alignment = center_alignment
                back_nums_str = ' '.join(f"{num:02d}" for num in rec['后区'])
                ws_60.cell(row=row_idx, column=3, value=back_nums_str).alignment = center_alignment
                ws_60.cell(row=row_idx, column=4, value=round(rec['总分'], 1)).alignment = center_alignment
                ws_60.cell(row=row_idx, column=5, value=rec['冷转热数量']).alignment = center_alignment
                ws_60.cell(row=row_idx, column=6, value=rec['奇偶比']).alignment = center_alignment
                ws_60.cell(row=row_idx, column=7, value=rec['大小比']).alignment = center_alignment
                ws_60.cell(row=row_idx, column=8, value=rec['和值']).alignment = center_alignment
                heat_chars = ', '.join(rec['热号特点']) if rec['热号特点'] else "均衡组合"
                ws_60.cell(row=row_idx, column=9, value=heat_chars).alignment = center_alignment
                
                for col in range(1, 10):
                    ws_60.cell(row=row_idx, column=col).border = thin_border
        
        # 3. 100期推荐工作表
        ws_100 = wb.create_sheet(title="100期推荐")
        
        # 设置列宽
        for col, width in enumerate(column_widths, 1):
            ws_100.column_dimensions[get_column_letter(col)].width = width
        
        # 添加标题
        ws_100.merge_cells('A1:I1')
        title_cell = ws_100['A1']
        title_cell.value = "大乐透100期推荐（长期趋势）"
        title_cell.font = Font(bold=True, size=14, color="A5A5A5")
        title_cell.alignment = center_alignment
        
        # 添加副标题
        ws_100.merge_cells('A2:I2')
        subtitle_cell = ws_100['A2']
        subtitle_cell.value = f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 基于最近100期热号分析+冷热转换"
        subtitle_cell.font = Font(size=10, italic=True)
        subtitle_cell.alignment = center_alignment
        
        # 添加表头
        for col, header in enumerate(headers, 1):
            cell = ws_100.cell(row=3, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill_100
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # 添加100期推荐数据
        if self.recommend_100:
            for row_idx, rec in enumerate(self.recommend_100, 4):
                ws_100.cell(row=row_idx, column=1, value=rec['组号']).alignment = center_alignment
                front_nums_str = ' '.join(f"{num:02d}" for num in rec['前区'])
                ws_100.cell(row=row_idx, column=2, value=front_nums_str).alignment = center_alignment
                back_nums_str = ' '.join(f"{num:02d}" for num in rec['后区'])
                ws_100.cell(row=row_idx, column=3, value=back_nums_str).alignment = center_alignment
                ws_100.cell(row=row_idx, column=4, value=round(rec['总分'], 1)).alignment = center_alignment
                ws_100.cell(row=row_idx, column=5, value=rec['冷转热数量']).alignment = center_alignment
                ws_100.cell(row=row_idx, column=6, value=rec['奇偶比']).alignment = center_alignment
                ws_100.cell(row=row_idx, column=7, value=rec['大小比']).alignment = center_alignment
                ws_100.cell(row=row_idx, column=8, value=rec['和值']).alignment = center_alignment
                heat_chars = ', '.join(rec['热号特点']) if rec['热号特点'] else "均衡组合"
                ws_100.cell(row=row_idx, column=9, value=heat_chars).alignment = center_alignment
                
                for col in range(1, 10):
                    ws_100.cell(row=row_idx, column=col).border = thin_border
        
        # 4. 综合推荐工作表
        ws_comp = wb.create_sheet(title="综合推荐")
        
        # 设置列宽
        comp_widths = [8, 25, 15, 15, 20, 15, 15, 15, 25]
        for col, width in enumerate(comp_widths, 1):
            ws_comp.column_dimensions[get_column_letter(col)].width = width
        
        # 添加标题
        ws_comp.merge_cells('A1:I1')
        title_cell = ws_comp['A1']
        title_cell.value = "大乐透综合推荐（整合30/60/100期）"
        title_cell.font = Font(bold=True, size=14, color="70AD47")
        title_cell.alignment = center_alignment
        
        # 添加副标题
        ws_comp.merge_cells('A2:I2')
        subtitle_cell = ws_comp['A2']
        subtitle_cell.value = f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 整合三个时间段的推荐结果"
        subtitle_cell.font = Font(size=10, italic=True)
        subtitle_cell.alignment = center_alignment
        
        # 添加表头
        comp_headers = ['组号', '前区号码', '后区号码', '综合评分', '来源统计', '奇偶比', '大小比', '和值', '推荐特点']
        
        for col, header in enumerate(comp_headers, 1):
            cell = ws_comp.cell(row=3, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill_comprehensive
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # 添加综合推荐数据
        if self.comprehensive_recommendations:
            for row_idx, rec in enumerate(self.comprehensive_recommendations, 4):
                ws_comp.cell(row=row_idx, column=1, value=rec['组号']).alignment = center_alignment
                front_nums_str = ' '.join(f"{num:02d}" for num in rec['前区'])
                ws_comp.cell(row=row_idx, column=2, value=front_nums_str).alignment = center_alignment
                back_nums_str = ' '.join(f"{num:02d}" for num in rec['后区'])
                ws_comp.cell(row=row_idx, column=3, value=back_nums_str).alignment = center_alignment
                ws_comp.cell(row=row_idx, column=4, value=rec['综合评分']).alignment = center_alignment
                ws_comp.cell(row=row_idx, column=5, value=rec['来源统计']).alignment = center_alignment
                ws_comp.cell(row=row_idx, column=6, value=rec['奇偶比']).alignment = center_alignment
                ws_comp.cell(row=row_idx, column=7, value=rec['大小比']).alignment = center_alignment
                ws_comp.cell(row=row_idx, column=8, value=rec['和值']).alignment = center_alignment
                ws_comp.cell(row=row_idx, column=9, value="整合三个时间段推荐").alignment = center_alignment
                
                for col in range(1, 10):
                    ws_comp.cell(row=row_idx, column=col).border = thin_border
        
        # 5. 冷热转换分析工作表
        ws_transition = wb.create_sheet(title="冷热转换分析")
        
        # 设置列宽
        transition_widths = [10, 15, 15, 15, 15, 15, 15]
        for col, width in enumerate(transition_widths, 1):
            ws_transition.column_dimensions[get_column_letter(col)].width = width
        
        # 添加标题
        ws_transition.merge_cells('A1:G1')
        title_cell = ws_transition['A1']
        title_cell.value = "冷热转换分析详情"
        title_cell.font = Font(bold=True, size=14, color="FF0000")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 添加副标题
        ws_transition.merge_cells('A2:G2')
        subtitle_cell = ws_transition['A2']
        subtitle_cell.value = f"分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 基于早期30期vs近期30期对比"
        subtitle_cell.font = Font(size=10, italic=True)
        subtitle_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 冷转热分析
        row_idx = 3
        ws_transition.cell(row=row_idx, column=1, value="🔥 由冷转热的号码（重点推荐）").font = Font(bold=True, color="00B050")
        ws_transition.merge_cells(f'A{row_idx}:G{row_idx}')
        row_idx += 1
        
        # 冷转热表头
        cold_hot_headers = ['号码', '早期频率', '近期频率', '趋势强度', '趋势评级', '最新出现', '遗漏期数']
        
        for col, header in enumerate(cold_hot_headers, 1):
            cell = ws_transition.cell(row=row_idx, column=col, value=header)
            cell.font = header_font
            cell.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
            cell.alignment = center_alignment
            cell.border = thin_border
        
        row_idx += 1
        
        # 添加冷转热数据
        if self.cold_hot_transition.get('冷转热'):
            for item in self.cold_hot_transition['冷转热']:
                ws_transition.cell(row=row_idx, column=1, value=item['号码']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=2, value=item['早期频率']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=3, value=item['近期频率']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=4, value=f"{item['趋势强度']:.3f}").alignment = center_alignment
                ws_transition.cell(row=row_idx, column=5, value=item['趋势评级']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=6, value=item['最新出现']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=7, value=item['遗漏期数']).alignment = center_alignment
                
                for col in range(1, 8):
                    ws_transition.cell(row=row_idx, column=col).border = thin_border
                row_idx += 1
        
        # 热转冷分析
        row_idx += 1
        ws_transition.cell(row=row_idx, column=1, value="❄️ 由热转冷的号码（谨慎选择）").font = Font(bold=True, color="FF0000")
        ws_transition.merge_cells(f'A{row_idx}:G{row_idx}')
        row_idx += 1
        
        for col, header in enumerate(cold_hot_headers, 1):
            cell = ws_transition.cell(row=row_idx, column=col, value=header)
            cell.font = header_font
            cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            cell.alignment = center_alignment
            cell.border = thin_border
        
        row_idx += 1
        
        # 添加热转冷数据
        if self.cold_hot_transition.get('热转冷'):
            for item in self.cold_hot_transition['热转冷']:
                ws_transition.cell(row=row_idx, column=1, value=item['号码']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=2, value=item['早期频率']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=3, value=item['近期频率']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=4, value=f"{item['趋势强度']:.3f}").alignment = center_alignment
                ws_transition.cell(row=row_idx, column=5, value=item['趋势评级']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=6, value=item['最新出现']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=7, value=item['遗漏期数']).alignment = center_alignment
                
                for col in range(1, 8):
                    ws_transition.cell(row=row_idx, column=col).border = thin_border
                row_idx += 1
        
        # 反弹信号分析
        row_idx += 1
        ws_transition.cell(row=row_idx, column=1, value="📈 有反弹信号的冷号（值得关注）").font = Font(bold=True, color="FFC000")
        ws_transition.merge_cells(f'A{row_idx}:G{row_idx}')
        row_idx += 1
        
        bounce_headers = ['号码', '早期频率', '近期出现', '遗漏期数', '反弹概率', '热度变化', '关注度']
        
        for col, header in enumerate(bounce_headers, 1):
            cell = ws_transition.cell(row=row_idx, column=col, value=header)
            cell.font = header_font
            cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
            cell.alignment = center_alignment
            cell.border = thin_border
        
        row_idx += 1
        
        # 添加反弹信号数据
        if self.cold_hot_transition.get('反弹信号'):
            for item in self.cold_hot_transition['反弹信号']:
                ws_transition.cell(row=row_idx, column=1, value=item['号码']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=2, value=item['早期频率']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=3, value=item['近期出现']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=4, value=item['遗漏期数']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=5, value=item['反弹概率']).alignment = center_alignment
                ws_transition.cell(row=row_idx, column=6, value=item['热度变化']).alignment = center_alignment
                
                prob = float(item['反弹概率'].replace('%', ''))
                if prob >= 70:
                    attention = "高关注"
                elif prob >= 50:
                    attention = "中关注"
                else:
                    attention = "关注"
                
                ws_transition.cell(row=row_idx, column=7, value=attention).alignment = center_alignment
                
                for col in range(1, 8):
                    ws_transition.cell(row=row_idx, column=col).border = thin_border
                row_idx += 1
        
        # 6. 热号分析工作表（30期）
        ws_hot_analysis = wb.create_sheet(title="热号分析")
        
        # 设置列宽
        hot_widths = [10, 10, 15, 15, 15, 20, 15]
        for col, width in enumerate(hot_widths, 1):
            ws_hot_analysis.column_dimensions[get_column_letter(col)].width = width
        
        # 添加标题
        ws_hot_analysis.merge_cells('A1:G1')
        title_cell = ws_hot_analysis['A1']
        title_cell.value = "热号分析详情（基于最近30期）"
        title_cell.font = Font(bold=True, size=14, color="7030A0")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 添加副标题
        ws_hot_analysis.merge_cells('A2:G2')
        subtitle_cell = ws_hot_analysis['A2']
        subtitle_cell.value = f"分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 数据范围: {self.recent_30['期号'].min()}-{self.recent_30['期号'].max()}"
        subtitle_cell.font = Font(size=10, italic=True)
        subtitle_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 前区热号表头
        row_idx = 3
        ws_hot_analysis.cell(row=row_idx, column=1, value="前区热号（按频率排序）").font = Font(bold=True, color="7030A0")
        ws_hot_analysis.merge_cells(f'A{row_idx}:G{row_idx}')
        row_idx += 1
        
        hot_headers = ['排名', '号码', '总出现次数', '总频率', '近期频率', '热力等级', '遗漏期数']
        
        for col, header in enumerate(hot_headers, 1):
            cell = ws_hot_analysis.cell(row=row_idx, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill_hot_analysis
            cell.alignment = center_alignment
            cell.border = thin_border
        
        row_idx += 1
        
        # 添加前区热号数据
        if self.hot_analysis_30.get('前区热号'):
            for i, (num, info) in enumerate(list(self.hot_analysis_30['前区热号'].items())[:15], 1):
                ws_hot_analysis.cell(row=row_idx, column=1, value=i).alignment = center_alignment
                ws_hot_analysis.cell(row=row_idx, column=2, value=num).alignment = center_alignment
                ws_hot_analysis.cell(row=row_idx, column=3, value=info.get('总次数', 0)).alignment = center_alignment
                ws_hot_analysis.cell(row=row_idx, column=4, value=info.get('总频率', '0%')).alignment = center_alignment
                ws_hot_analysis.cell(row=row_idx, column=5, value=info.get('近期频率', '0%')).alignment = center_alignment
                ws_hot_analysis.cell(row=row_idx, column=6, value=info.get('热力等级', '')).alignment = center_alignment
                ws_hot_analysis.cell(row=row_idx, column=7, value=info.get('遗漏期数', 0)).alignment = center_alignment
                
                for col in range(1, 8):
                    ws_hot_analysis.cell(row=row_idx, column=col).border = thin_border
                row_idx += 1
        
        # 后区热号
        row_idx += 1
        ws_hot_analysis.cell(row=row_idx, column=1, value="后区热号（按频率排序）").font = Font(bold=True, color="7030A0")
        ws_hot_analysis.merge_cells(f'A{row_idx}:G{row_idx}')
        row_idx += 1
        
        back_hot_headers = ['排名', '号码', '总出现次数', '总频率', '遗漏期数']
        
        for col, header in enumerate(back_hot_headers, 1):
            cell = ws_hot_analysis.cell(row=row_idx, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill_hot_analysis
            cell.alignment = center_alignment
            cell.border = thin_border
        
        row_idx += 1
        
        # 添加后区热号数据
        if self.hot_analysis_30.get('后区热号'):
            for i, (num, info) in enumerate(list(self.hot_analysis_30['后区热号'].items())[:8], 1):
                ws_hot_analysis.cell(row=row_idx, column=1, value=i).alignment = center_alignment
                ws_hot_analysis.cell(row=row_idx, column=2, value=num).alignment = center_alignment
                ws_hot_analysis.cell(row=row_idx, column=3, value=info.get('总次数', 0)).alignment = center_alignment
                ws_hot_analysis.cell(row=row_idx, column=4, value=info.get('总频率', '0%')).alignment = center_alignment
                ws_hot_analysis.cell(row=row_idx, column=5, value=info.get('遗漏期数', 0)).alignment = center_alignment
                
                for col in range(1, 6):
                    ws_hot_analysis.cell(row=row_idx, column=col).border = thin_border
                row_idx += 1
        
        # 7. 分析过程记录工作表
        ws_process = wb.create_sheet(title="分析过程记录")
        
        # 设置列宽
        process_widths = [25, 50, 30, 20]
        for col, width in enumerate(process_widths, 1):
            ws_process.column_dimensions[get_column_letter(col)].width = width
        
        # 添加标题
        ws_process.merge_cells('A1:D1')
        title_cell = ws_process['A1']
        title_cell.value = "分析过程记录"
        title_cell.font = Font(bold=True, size=14, color="000000")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 添加表头
        process_headers = ['分析步骤', '分析内容', '分析参数', '分析结果']
        
        for col, header in enumerate(process_headers, 1):
            cell = ws_process.cell(row=2, column=col, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # 添加分析过程数据
        process_data = [
            ("数据加载", f"成功加载{len(self.data)}期开奖数据", f"数据范围: {self.data['期号'].min()}-{self.data['期号'].max()}", "成功"),
            ("数据划分", "划分三个时间段数据", "30期/60期/100期", f"{len(self.recent_30)}/{len(self.recent_60)}/{len(self.recent_100)}期"),
            ("热号分析(30期)", "分析最近30期热号", "取前15个热号", f"识别{len(self.hot_analysis_30.get('前区热号', {}))}个前区热号"),
            ("热号分析(60期)", "分析最近60期热号", "取前15个热号", f"识别{len(self.hot_analysis_60.get('前区热号', {}))}个前区热号"),
            ("热号分析(100期)", "分析最近100期热号", "取前15个热号", f"识别{len(self.hot_analysis_100.get('前区热号', {}))}个前区热号"),
            ("冷热转换分析", "分析号码冷热转换趋势", "早期30期vs近期30期对比", f"识别{len(self.cold_hot_transition.get('冷转热', []))}个冷转热号码"),
            ("30期推荐生成", "基于30期热号+冷热转换", "生成5组推荐", f"生成{len(self.recommend_30)}组推荐"),
            ("60期推荐生成", "基于60期热号+冷热转换", "生成5组推荐", f"生成{len(self.recommend_60)}组推荐"),
            ("100期推荐生成", "基于100期热号+冷热转换", "生成5组推荐", f"生成{len(self.recommend_100)}组推荐"),
            ("综合推荐生成", "整合三个时间段推荐", "基于频率加权", f"生成{len(self.comprehensive_recommendations)}组综合推荐"),
            ("Excel输出", "生成完整分析报告", "7个工作表", "完成")
        ]
        
        for i, (step, content, params, result) in enumerate(process_data, 3):
            ws_process.cell(row=i, column=1, value=step).border = thin_border
            ws_process.cell(row=i, column=2, value=content).border = thin_border
            ws_process.cell(row=i, column=3, value=params).border = thin_border
            ws_process.cell(row=i, column=4, value=result).border = thin_border
        
        # 8. 分析摘要工作表
        summary_ws = wb.create_sheet(title="分析摘要")
        
        # 设置列宽
        summary_widths = [20, 20, 25, 20, 20]
        for col, width in enumerate(summary_widths, 1):
            summary_ws.column_dimensions[get_column_letter(col)].width = width
        
        # 添加标题
        summary_ws.merge_cells('A1:E1')
        title_cell = summary_ws['A1']
        title_cell.value = "大乐透分析摘要"
        title_cell.font = Font(bold=True, size=14, color="366092")
        title_cell.alignment = center_alignment
        
        # 添加摘要表头
        summary_headers = ['分析项目', '分析结果', '说明', '参考意义', '建议']
        
        for col, header in enumerate(summary_headers, 1):
            cell = summary_ws.cell(row=2, column=col, value=header)
            cell.font = header_font
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # 添加摘要数据
        summary_data = [
            ("总分析期数", self.analysis_results['基本信息']['总期数'], "历史开奖总期数", "数据量充足，分析可靠", "可信任度高"),
            ("30期推荐数量", len(self.recommend_30), "基于最近30期热号分析", "短期趋势明显", "适合关注近期热号的用户"),
            ("60期推荐数量", len(self.recommend_60), "基于最近60期热号分析", "中期趋势稳定", "适合稳健型用户"),
            ("100期推荐数量", len(self.recommend_100), "基于最近100期热号分析", "长期趋势可靠", "适合保守型用户"),
            ("综合推荐数量", len(self.comprehensive_recommendations), "整合三个时间段", "综合优势明显", "适合初次使用或全面考虑的用户"),
            ("冷转热号码", len(self.cold_hot_transition.get('冷转热', [])), "趋势向上号码", "反弹概率较高", "重点选择1-2个"),
            ("热转冷号码", len(self.cold_hot_transition.get('热转冷', [])), "趋势向下号码", "近期可能降温", "谨慎选择或避免"),
            ("反弹信号", len(self.cold_hot_transition.get('反弹信号', [])), "有反弹潜力的冷号", "可能迎来反弹", "适当关注"),
            ("前区热号数量", len(self.hot_analysis_30.get('前区热号', {})), "基于30期的热号", "短期热门号码", "建议选择2-3个"),
            ("后区热号数量", len(self.hot_analysis_30.get('后区热号', {})), "基于30期的热号", "短期热门号码", "建议选择1-2个"),
            ("生成时间", datetime.now().strftime('%Y-%m-%d %H:%M:%S'), "分析完成时间", "时效性参考", "建议每周更新"),
            ("数据更新时间", str(self.data['开奖日期'].max()), "最新开奖日期", "数据新鲜度", "保持数据更新")
        ]
        
        for i, (item, result, desc, meaning, suggestion) in enumerate(summary_data, 3):
            summary_ws.cell(row=i, column=1, value=item).border = thin_border
            summary_ws.cell(row=i, column=2, value=result).border = thin_border
            summary_ws.cell(row=i, column=3, value=desc).border = thin_border
            summary_ws.cell(row=i, column=4, value=meaning).border = thin_border
            summary_ws.cell(row=i, column=5, value=suggestion).border = thin_border
        
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
            # 1. 分析30期热号
            print("\n正在执行30期热号分析...")
            self.hot_analysis_30 = self.analyze_hot_numbers(self.recent_30, "30期")
            
            # 2. 分析60期热号
            print("\n正在执行60期热号分析...")
            self.hot_analysis_60 = self.analyze_hot_numbers(self.recent_60, "60期")
            
            # 3. 分析100期热号
            print("\n正在执行100期热号分析...")
            self.hot_analysis_100 = self.analyze_hot_numbers(self.recent_100, "100期")
            
            # 4. 进行冷热转换分析
            print("\n正在执行冷热转换分析...")
            self.analyze_cold_hot_transition()
            
            # 5. 生成30期推荐
            print("\n正在生成30期推荐...")
            self.recommend_30 = self.generate_recommendations_for_period(self.recent_30, "30期")
            
            # 6. 生成60期推荐
            print("\n正在生成60期推荐...")
            self.recommend_60 = self.generate_recommendations_for_period(self.recent_60, "60期")
            
            # 7. 生成100期推荐
            print("\n正在生成100期推荐...")
            self.recommend_100 = self.generate_recommendations_for_period(self.recent_100, "100期")
            
            # 8. 生成综合推荐
            print("\n正在生成综合推荐...")
            self.generate_comprehensive_recommendations()
            
            # 9. 生成Excel推荐文件
            print("\n正在生成Excel输出文件...")
            excel_file = self.generate_excel_recommendations()
            
            print("\n" + "=" * 80)
            print("分析完成！")
            print("=" * 80)
            print(f"✓ 推荐文件: {excel_file}")
            print("\n文件包含的工作表:")
            print("1. 30期推荐 - 基于最近30期热号分析（短期趋势）")
            print("2. 60期推荐 - 基于最近60期热号分析（中期趋势）")
            print("3. 100期推荐 - 基于最近100期热号分析（长期趋势）")
            print("4. 综合推荐 - 整合三个时间段的推荐结果")
            print("5. 冷热转换分析 - 详细的冷热转换趋势分析")
            print("6. 热号分析 - 详细的热号统计和分析")
            print("7. 分析过程记录 - 完整的分析过程记录")
            print("8. 分析摘要 - 分析的基本信息和参数")
            print("\n选号策略建议:")
            print("1. 综合推荐：整合三个时间段的优势，适合初次使用")
            print("2. 30期推荐：关注近期趋势，适合激进型选号")
            print("3. 60期推荐：平衡近期与历史，适合稳健型选号")
            print("4. 100期推荐：重视历史规律，适合保守型选号")
            print("5. 重点关注冷转热号码（冷热转换分析表中）")
            print("6. 适当关注反弹信号号码")
            print("7. 谨慎选择热转冷号码")
            print("\n温馨提示:")
            print("• 彩票具有随机性，分析结果仅供参考")
            print("• 理性购彩，量力而行")
            print("• 祝您好运！")
            print("=" * 80)
            
            return excel_file
            
        except Exception as e:
            print(f"分析过程出错: {e}")
            import traceback
            traceback.print_exc()
            return None


# 主程序
if __name__ == "__main__":
    # 设置文件路径
    file_path = "/home/acuto/python_stu/data_analysis/大乐透开奖数据.xlsx"
    
    print("大乐透智能分析系统（完整版）")
    print("=" * 80)
    print("时间窗口：30期（短期）、60期（中期）、100期（长期）")
    print("输出内容：四个推荐表 + 冷热分析 + 分析过程")
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