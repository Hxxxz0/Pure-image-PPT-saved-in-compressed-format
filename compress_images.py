#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片压缩脚本 - 大幅压缩图片大小但保持视觉清晰度
针对PPT制作优化，确保最终文件小于20MB
"""

import os
import glob
from PIL import Image
import pillow_heif
from pathlib import Path

# 注册HEIF格式支持
pillow_heif.register_heif_opener()

def get_user_input():
    """获取用户输入的配置"""
    print("=" * 60)
    print("PPT图片压缩工具 - 配置向导")
    print("=" * 60)
    
    # 获取文件名模式
    print("\n1. 请输入图片文件名模式:")
    print("   例如: 'slide_*.png' 或 'presentation_*.jpg' 或 '*.png'")
    print("   提示: 使用 * 代表通配符，确保文件名中包含数字用于排序")
    
    pattern = input("   文件名模式: ").strip()
    if not pattern:
        pattern = "*.png"
        print(f"   使用默认模式: {pattern}")
    
    # 获取目标PPT大小
    print("\n2. 请输入目标PPT文件大小限制 (MB):")
    print("   例如: 18 (推荐为20MB以下)")
    
    try:
        size_input = input("   目标大小 (MB): ").strip()
        target_size = float(size_input) if size_input else 18.0
    except ValueError:
        target_size = 18.0
        print(f"   使用默认大小: {target_size} MB")
    
    # 获取输出PPT文件名
    print("\n3. 请输入输出PPT文件名前缀:")
    print("   例如: '我的演示文稿' 或 '项目汇报'")
    
    output_name = input("   文件名前缀: ").strip()
    if not output_name:
        output_name = "图片集"
        print(f"   使用默认前缀: {output_name}")
    
    return pattern, target_size, output_name

def optimize_image_for_ppt(image_path, output_path=None, target_size_mb=0.5):
    """
    为PPT优化图片，使用JPEG格式获得更好的压缩比
    target_size_mb: 每张图片的目标大小(MB)
    """
    if output_path is None:
        name, ext = os.path.splitext(image_path)
        output_path = f"{name}_optimized.jpg"
    
    try:
        with Image.open(image_path) as img:
            # 保存原始信息
            original_size = os.path.getsize(image_path)
            original_width, original_height = img.size
            
            print(f"处理: {os.path.basename(image_path)}")
            print(f"  原始大小: {original_size / 1024 / 1024:.2f} MB")
            print(f"  原始尺寸: {original_width}x{original_height}")
            
            # 转换为RGB模式（JPEG不支持透明度）
            if img.mode in ('RGBA', 'LA', 'P'):
                # 创建白色背景
                rgb_img = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                rgb_img.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
                img = rgb_img
                print(f"  转换为RGB模式")
            
            # 计算合适的尺寸和质量
            # 对于PPT，1920x1080是很好的分辨率
            max_width = 1920
            max_height = 1080
            
            # 保持纵横比缩放
            if original_width > max_width or original_height > max_height:
                img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
                print(f"  调整尺寸为: {img.size[0]}x{img.size[1]}")
            
            # 使用二分法找到合适的质量设置
            def get_file_size(quality):
                temp_path = output_path + ".temp"
                img.save(temp_path, 'JPEG', quality=quality, optimize=True)
                size = os.path.getsize(temp_path)
                os.remove(temp_path)
                return size
            
            # 目标文件大小（字节）
            target_size = target_size_mb * 1024 * 1024
            
            # 二分查找最佳质量
            low_quality, high_quality = 30, 95
            best_quality = 85
            
            for _ in range(10):  # 最多尝试10次
                mid_quality = (low_quality + high_quality) // 2
                file_size = get_file_size(mid_quality)
                
                if file_size <= target_size:
                    best_quality = mid_quality
                    low_quality = mid_quality + 1
                else:
                    high_quality = mid_quality - 1
                
                if low_quality > high_quality:
                    break
            
            # 使用找到的最佳质量保存
            img.save(output_path, 'JPEG', quality=best_quality, optimize=True)
            
            # 计算压缩效果
            compressed_size = os.path.getsize(output_path)
            compression_ratio = (1 - compressed_size / original_size) * 100
            
            print(f"  使用质量: {best_quality}")
            print(f"  压缩后大小: {compressed_size / 1024 / 1024:.2f} MB")
            print(f"  压缩率: {compression_ratio:.1f}%")
            print(f"  保存为: {os.path.basename(output_path)}")
            print()
            
            return {
                'original_size': original_size,
                'compressed_size': compressed_size,
                'compression_ratio': compression_ratio,
                'input_file': image_path,
                'output_file': output_path,
                'quality': best_quality
            }
            
    except Exception as e:
        print(f"处理 {image_path} 时出错: {str(e)}")
        return None

def batch_compress_for_ppt(input_pattern="*.png", target_ppt_size_mb=18, output_prefix="图片集"):
    """
    批量压缩图片用于PPT制作
    input_pattern: 输入图片文件名模式
    target_ppt_size_mb: 目标PPT总大小(MB)
    output_prefix: 输出文件名前缀
    """
    image_files = glob.glob(input_pattern)
    
    # 尝试按文件名中的数字排序
    import re
    def extract_number(filename):
        numbers = re.findall(r'\d+', os.path.basename(filename))
        return int(numbers[-1]) if numbers else 0
    
    image_files.sort(key=extract_number)
    
    if not image_files:
        print(f"未找到匹配模式 '{input_pattern}' 的图片文件")
        print("请检查文件名模式是否正确，或确保图片文件存在于当前目录")
        return None
    
    # 计算每张图片的目标大小
    # 为PPT结构预留2MB空间
    available_space = (target_ppt_size_mb - 2) * 1024 * 1024
    target_size_per_image = available_space / len(image_files) / 1024 / 1024  # MB
    
    print(f"找到 {len(image_files)} 个图片文件")
    print(f"目标PPT大小: {target_ppt_size_mb} MB")
    print(f"每张图片目标大小: {target_size_per_image:.2f} MB")
    print("=" * 50)
    
    results = []
    total_original_size = 0
    total_compressed_size = 0
    
    # 创建压缩文件夹
    compressed_dir = "compressed_for_ppt"
    os.makedirs(compressed_dir, exist_ok=True)
    
    for image_file in image_files:
        # 生成输出文件名
        name = os.path.splitext(os.path.basename(image_file))[0]
        output_file = os.path.join(compressed_dir, f"{name}_optimized.jpg")
        
        result = optimize_image_for_ppt(image_file, output_file, target_size_per_image)
        if result:
            results.append(result)
            total_original_size += result['original_size']
            total_compressed_size += result['compressed_size']
    
    # 显示总体统计
    print("=" * 50)
    print("压缩完成！总体统计:")
    print(f"处理文件数: {len(results)}")
    print(f"原始总大小: {total_original_size / 1024 / 1024:.2f} MB")
    print(f"压缩后总大小: {total_compressed_size / 1024 / 1024:.2f} MB")
    print(f"总体压缩率: {(1 - total_compressed_size / total_original_size) * 100:.1f}%")
    print(f"节省空间: {(total_original_size - total_compressed_size) / 1024 / 1024:.2f} MB")
    print(f"预计PPT大小: {(total_compressed_size / 1024 / 1024) + 2:.2f} MB")
    print(f"压缩后的文件保存在: {compressed_dir}/ 目录中")
    
    # 保存配置信息
    config = {
        'input_pattern': input_pattern,
        'target_ppt_size_mb': target_ppt_size_mb,
        'output_prefix': output_prefix,
        'compressed_dir': compressed_dir,
        'file_count': len(results)
    }
    
    return config

if __name__ == "__main__":
    print("欢迎使用PPT图片压缩工具！")
    print("本工具可以将图片批量压缩，确保生成的PPT文件小于指定大小。")
    print()
    
    # 获取用户配置
    pattern, target_size, output_prefix = get_user_input()
    
    print("\n" + "=" * 60)
    print("开始压缩处理...")
    print("=" * 60)
    
    # 批量压缩图片
    config = batch_compress_for_ppt(pattern, target_size, output_prefix)
    
    if config:
        print("\n" + "=" * 60)
        print("压缩完成！")
        print("下一步：运行 'python create_ppt.py' 来创建PPT文件")
        print("=" * 60) 