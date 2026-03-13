import pandas as pd
import json

def json_to_excel(json_path, excel_path):
    """将 JSON 文件转换为 Excel 表格"""
    try:
        # 读取 JSON 文件
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 转换为 DataFrame 并导出为 Excel
        df = pd.DataFrame(data)
        df.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"✅ 成功！已将 {json_path} 转换为 {excel_path}")
        
    except Exception as e:
        print(f"❌ JSON 转 Excel 失败: {e}")

def excel_to_json(excel_path, json_path):
    """将 Excel 表格转换为 JSON 文件"""
    try:
        # 读取 Excel 文件，将所有内容当做字符串处理
        df = pd.read_excel(excel_path, engine='openpyxl', dtype=str)
        
        # 将表格里的 NaN（空值）替换为空字符串
        df = df.fillna("")
        
        # 🌟 新增：专门处理 date 列，去掉 Excel 自动加上的 " 00:00:00"
        if 'date' in df.columns:
            df['date'] = df['date'].str.replace(' 00:00:00', '', regex=False)
            
        # 转换为字典列表格式
        data = df.to_dict(orient='records')
        
        # 写入 JSON 文件
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        print(f"✅ 成功！已将 {excel_path} 转换为 {json_path}")
        
    except Exception as e:
        print(f"❌ Excel 转 JSON 失败: {e}")

if __name__ == "__main__":
    # 假设你的原始数据叫 data.json
    json_file = "schedule.json"
    excel_file = "seminar_schedule.xlsx"

    # --- 使用方法 ---
    
    # 1. 第一次使用：先把现有的 JSON 转换成 Excel
    # json_to_excel(json_file, excel_file)
    
    # 2. 以后日常维护：你只需要打开 seminar_schedule.xlsx 编辑数据，
    #    编辑保存后，运行下面这行代码，就能生成最新的 JSON
    excel_to_json(excel_file, json_file)