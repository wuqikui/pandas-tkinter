import sys
import pandas as pd

def validate_row(row):
    errors = []
    # 检查地号
    if pd.isna(row.get("地号")) or str(row["地号"]).strip() == "":
        errors.append("地号为空")
    # 检查土地性质
    land_type = row.get("土地性质")
    if pd.isna(land_type) or land_type not in ["国有", "集体"]:
        errors.append("土地性质值不合法")
    return errors

def main():
    try:
        # 读取测试数据
        df = pd.read_excel("Python测试练习数据.xls", engine="xlrd")
        
        # 获取区域参数
        target_area = sys.argv[1] if len(sys.argv) > 1 else None
        
        # 筛选数据
        if target_area:
            filtered_df = df[df["区域"] == f"{target_area}区"]
        else:
            filtered_df = df

        # 执行校验
        error_records = []
        for index, row in filtered_df.iterrows():
            if validate_row(row):
                errors = validate_row(row)
                error_msg = f"第 {index + 2} 行数据，{', '.join(errors)}"  # +2因Excel标题行
                error_records.append(error_msg)

        # 写入结果文件
        #timestamp = pd.Timestamp.now().strftime("%Y%m%d%H%M")
        output_file = f"checkResult.txt"
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(";\n".join(error_records))
            f.write(';')
        
        print(f"发现 {len(error_records)} 条问题记录，已保存到 {output_file}")

    except IndexError:
        print("参数错误：最多只能接收1个区域参数（如'A'代表A区）")
    except FileNotFoundError:
        print("错误：未找到Python测试练习数据.xls文件")
    except Exception as e:
        print(f"运行错误：{str(e)}")

if __name__ == "__main__":
    main()