import pandas as pd
import os
import glob
import logging
from datetime import datetime

# 读取配置
config = pd.read_excel('config.xlsx', sheet_name=None)
basic = config['基本设置'].iloc[0]
source_folder = basic['源文件夹路径']
output_file = basic['输出文件路径']
include_sub = basic['包含子文件夹'] == '是'

# 获取文件列表
pattern = os.path.join(source_folder, '**', '*.xlsx') if include_sub else os.path.join(source_folder, '*.xlsx')
files = glob.glob(pattern, recursive=include_sub)

# 合并
all_data = []
for f in files:
    try:
        df = pd.read_excel(f, sheet_name=0)
        df['来源文件'] = os.path.basename(f)
        all_data.append(df)
    except Exception as e:
        logging.error(f'读取失败 {f}: {e}')

if all_data:
    merged = pd.concat(all_data, ignore_index=True)
    # 清洗示例：去重
    dedup_cols = config['清洗规则']['去重列'].dropna().tolist()
    if dedup_cols:
        merged.drop_duplicates(subset=dedup_cols, inplace=True)
    # 日期格式化
    date_cols = config['清洗规则']['日期列'].dropna().tolist()
    for col in date_cols:
        merged[col] = pd.to_datetime(merged[col], errors='coerce').dt.strftime('%Y-%m-%d')
    merged.to_excel(output_file, index=False)
    logging.info(f'合并完成，共 {len(merged)} 行，输出至 {output_file}')
else:
    logging.warning('未找到可合并的文件')
