import re
from collections import Counter
import os

def clean_text(text):
    """
    清洗文本：转换为小写，去除标点符号
    """
    # 将文本转换为小写
    text = text.lower()
    
    # 使用正则表达式去除标点符号，只保留字母、数字和空格
    # \w 匹配字母、数字、下划线，[^\w\s] 匹配非字母数字和下划线非空格的字符
    text = re.sub(r'[^\w\s]', '', text)
    
    return text

def get_word_frequency(word_list):
    """
    统计单词频率
    """
    # 方法1：使用字典手动统计
    word_count = {}
    for word in word_list:
        if word:  # 确保单词不为空
            if word in word_count:
                word_count[word] += 1
            else:
                word_count[word] = 1
    
    return word_count

def get_word_frequency_counter(word_list):
    """
    使用Collections.Counter统计单词频率（更简洁的方法）
    """
    # 过滤空字符串
    filtered_words = [word for word in word_list if word]
    return Counter(filtered_words)

def sort_word_frequency(word_count):
    """
    按频率从高到低排序
    """
    # 将字典转换为列表进行排序
    sorted_words = sorted(word_count.items(), key=lambda x: x[1], reverse=True)
    return sorted_words

def read_file(filename):
    """
    读取文件内容
    """
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            content = file.read()
        return content
    except FileNotFoundError:
        print(f"错误：文件 '{filename}' 未找到！")
        return None
    except Exception as e:
        print(f"读取文件时发生错误：{e}")
        return None

def display_results(sorted_words, top_n=None):
    """
    显示结果
    """
    if not sorted_words:
        print("没有找到任何单词！")
        return
    
    if top_n:
        sorted_words = sorted_words[:top_n]
    
    print("\n" + "="*50)
    print("单词频率统计结果")
    print("="*50)
    print(f"{'排名':<6} {'单词':<20} {'频率':<10} {'百分比':<10}")
    print("-"*50)
    
    total_words = sum(count for _, count in sorted_words)
    
    for rank, (word, count) in enumerate(sorted_words, 1):
        percentage = (count / total_words) * 100
        print(f"{rank:<6} {word:<20} {count:<10} {percentage:.2f}%")
    
    print("-"*50)
    print(f"总单词数（去重）：{len(sorted_words)}")
    print(f"总单词数（所有）：{total_words}")

def analyze_text_detailed(text):
    """
    详细分析文本
    """
    print("正在分析文本...")
    
    # 原始文本统计
    original_chars = len(text)
    original_words = len(text.split())
    
    print(f"原始文本字符数：{original_chars}")
    print(f"原始文本单词数：{original_words}")
    
    # 清洗文本
    cleaned_text = clean_text(text)
    print(f"清洗后字符数：{len(cleaned_text)}")
    
    # 分割单词
    words = cleaned_text.split()
    print(f"有效单词数：{len(words)}")
    
    return words

def save_results_to_file(sorted_words, filename="word_frequency_results.txt"):
    """
    将结果保存到文件
    """
    try:
        with open(filename, 'w', encoding='utf-8') as file:
            file.write("单词频率统计结果\n")
            file.write("="*50 + "\n")
            file.write(f"{'排名':<6} {'单词':<20} {'频率':<10} {'百分比':<10}\n")
            file.write("-"*50 + "\n")
            
            total_words = sum(count for _, count in sorted_words)
            
            for rank, (word, count) in enumerate(sorted_words, 1):
                percentage = (count / total_words) * 100
                file.write(f"{rank:<6} {word:<20} {count:<10} {percentage:.2f}%\n")
            
            file.write("-"*50 + "\n")
            file.write(f"总单词数（去重）：{len(sorted_words)}\n")
            file.write(f"总单词数（所有）：{total_words}\n")
        
        print(f"\n结果已保存到文件：{filename}")
    except Exception as e:
        print(f"保存文件时发生错误：{e}")

def main():
    """
    主函数
    """
    print("单词频率分析器")
    print("=" * 30)
    
    # 获取文件名
    filename = input("请输入要分析的文本文件名（或直接回车使用默认文件）：").strip()
    
    # 如果没有输入文件名，使用默认文件
    if not filename:
        # 创建一个示例文件
        filename = "sample_text.txt"
        create_sample_file(filename)
        print(f"使用示例文件：{filename}")
    
    # 读取文件
    print(f"\n正在读取文件：{filename}")
    text = read_file(filename)
    
    if text is None:
        return
    
    # 详细分析文本
    words = analyze_text_detailed(text)
    
    if not words:
        print("文本中没有找到有效的单词！")
        return
    
    # 统计词频（使用两种方法，这里选择Counter方法）
    print("\n正在统计单词频率...")
    word_count = get_word_frequency_counter(words)
    
    # 排序
    print("正在排序...")
    sorted_words = sort_word_frequency(word_count)
    
    # 显示结果
    top_n = input("\n显示前多少名？（直接回车显示全部）：").strip()
    if top_n and top_n.isdigit():
        display_results(sorted_words, int(top_n))
    else:
        display_results(sorted_words)
    
    # 保存结果
    save_choice = input("\n是否保存结果到文件？(y/n)：").strip().lower()
    if save_choice == 'y':
        save_filename = input("请输入保存文件名（直接回车使用默认文件名）：").strip()
        if not save_filename:
            save_filename = "word_frequency_results.txt"
        save_results_to_file(sorted_words, save_filename)

def create_sample_file(filename):
    """
    创建示例文本文件
    """
    sample_text = """Hello world! This is a sample text for word frequency analysis.
The quick brown fox jumps over the lazy dog. Hello again, world!
Python is a great programming language. I love Python programming.
This text contains repeated words to test our frequency counter.
Word frequency analysis helps in text mining and natural language processing.
Let's see how many times each word appears in this sample text."""

    try:
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(sample_text)
    except Exception as e:
        print(f"创建示例文件时发生错误：{e}")

def advanced_analysis():
    """
    高级分析功能
    """
    print("\n高级分析功能")
    print("=" * 30)
    
    filename = input("请输入文本文件名：").strip()
    if not filename:
        print("文件名不能为空！")
        return
    
    text = read_file(filename)
    if text is None:
        return
    
    words = analyze_text_detailed(text)
    
    if not words:
        return
    
    # 使用Counter进行更详细的分析
    word_counter = Counter(words)
    
    # 基本统计
    total_unique_words = len(word_counter)
    total_words = sum(word_counter.values())
    most_common = word_counter.most_common(5)
    
    print(f"\n详细统计信息：")
    print(f"总单词数：{total_words}")
    print(f"唯一单词数：{total_unique_words}")
    print(f"词汇丰富度：{total_unique_words/total_words:.2%}")
    print(f"最常用的5个单词：")
    
    for word, count in most_common:
        percentage = (count / total_words) * 100
        print(f"  '{word}': {count}次 ({percentage:.2f}%)")
    
    # 长度分析
    word_lengths = [len(word) for word in words]
    avg_length = sum(word_lengths) / len(word_lengths)
    print(f"\n单词平均长度：{avg_length:.2f} 个字母")
    
    # 按单词长度分组
    length_groups = {}
    for word in words:
        length = len(word)
        if length in length_groups:
            length_groups[length] += 1
        else:
            length_groups[length] = 1
    
    print("\n单词长度分布：")
    for length in sorted(length_groups.keys()):
        count = length_groups[length]
        percentage = (count / total_words) * 100
        print(f"  {length}个字母：{count}个单词 ({percentage:.2f}%)")

if __name__ == "__main__":
    while True:
        print("\n单词频率分析器")
        print("=" * 30)
        print("1. 基础分析")
        print("2. 高级分析")
        print("3. 退出")
        
        choice = input("请选择功能（1-3）：").strip()
        
        if choice == '1':
            main()
        elif choice == '2':
            advanced_analysis()
        elif choice == '3':
            print("谢谢使用！")
            break
        else:
            print("无效选择，请重新输入！")
        
        input("\n按回车键继续...")