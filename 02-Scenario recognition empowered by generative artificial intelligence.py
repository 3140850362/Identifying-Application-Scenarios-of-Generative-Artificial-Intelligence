"""
逐行读取 Excel 中的 USE 字段，调用 DeepSeek API 进行行业分类（中文提示词）
支持指定行范围、自动保存中间结果、断点续跑
"""

import pandas as pd
import json
import re
import time
import os
from openai import OpenAI
import numpy as np



# ==================== 配置参数 ====================
# DeepSeek API 配置
API_KEY = ""          # 替换为你的 API Key
BASE_URL = "https://api.deepseek.com"
MODEL_NAME = "deepseek-v4-flash"      

# 文件路径
INPUT_EXCEL = "专利USE用法61365.xlsx"       # 输入文件
OUTPUT_EXCEL = "01结果.xlsx"

# Excel 列名（请根据实际修改）
COL_PATENT_ID = "专利号"                # 专利号/ID 列
COL_USE = "用法USE"                       # USE 字段列

# 行范围控制（行号从 1 开始，包含起始行，不包含结束行）
# 例如：start_line = 1, end_line = 101 表示处理第 1 行到第 100 行
start_line = 1
end_line = 61366          # 如果设置为 None 或大于总行数，则处理到文件末尾

# 中间结果保存间隔（每处理多少行保存一次）
SAVE_INTERVAL = 10

# ==================== 行业分类列表（98类，含 00.未指定）====================
INDUSTRY_LIST = [
    "00.未指定",
    "01.农业",
    "02.林业",
    "03.畜牧业",
    "04.渔业",
    "05.农、林、牧、渔专业及辅助性活动",
    "06.煤炭开采和洗选业",
    "07.石油和天然气开采业",
    "08.黑色金属矿采选业",
    "09.有色金属矿采选业",
    "10.非金属矿采选业",
    "11.开采专业及辅助性活动",
    "12.其他采矿业",
    "13.农副食品加工业",
    "14.食品制造业",
    "15.酒、饮料和精制茶制造业",
    "16.烟草制品业",
    "17.纺织业",
    "18.纺织服装、服饰业",
    "19.皮革、毛皮、羽毛及其制品和制鞋业",
    "20.木材加工和木、竹、藤、棕、草制品业",
    "21.家具制造业",
    "22.造纸和纸制品业",
    "23.印刷和记录媒介复制业",
    "24.文教、工美、体育和娱乐用品制造业",
    "25.石油、煤炭及其他燃料加工业",
    "26.化学原料和化学制品制造业",
    "27.医药制造业",
    "28.化学纤维制造业",
    "29.橡胶和塑料制品业",
    "30.非金属矿物制品业",
    "31.黑色金属冶炼和压延加工业",
    "32.有色金属冶炼和压延加工业",
    "33.金属制品业",
    "34.通用设备制造业",
    "35.专用设备制造业",
    "36.汽车制造业",
    "37.铁路、船舶、航空航天和其他运输设备制造业",
    "38.电气机械和器材制造业",
    "39.计算机、通信和其他电子设备制造业",
    "40.仪器仪表制造业",
    "41.其他制造业",
    "42.废弃资源综合利用业",
    "43.金属制品、机械和设备修理业",
    "44.电力、热力生产和供应业",
    "45.燃气生产和供应业",
    "46.水的生产和供应业",
    "47.房屋建筑业",
    "48.土木工程建筑业",
    "49.建筑安装业",
    "50.建筑装饰、装修和其他建筑业",
    "51.批发业",
    "52.零售业",
    "53.铁路运输业",
    "54.道路运输业",
    "55.水上运输业",
    "56.航空运输业",
    "57.管道运输业",
    "58.多式联运和运输代理业",
    "59.装卸搬运和仓储业",
    "60.邮政业",
    "61.住宿业",
    "62.餐饮业",
    "63.电信、广播电视和卫星传输服务",
    "64.互联网和相关服务",
    "65.软件和信息技术服务业",
    "66.货币金融服务",
    "67.资本市场服务",
    "68.保险业",
    "69.其他金融业",
    "70.房地产业",
    "71.租赁业",
    "72.商务服务业",
    "73.研究和试验发展",
    "74.专业技术服务业",
    "75.科技推广和应用服务业",
    "76.水利管理业",
    "77.生态保护和环境治理业",
    "78.公共设施管理业",
    "79.土地管理业",
    "80.居民服务业",
    "81.机动车、电子产品和日用产品修理业",
    "82.其他服务业",
    "83.教育",
    "84.卫生",
    "85.社会工作",
    "86.新闻和出版业",
    "87.广播、电视、电影和录音制作业",
    "88.文化艺术业",
    "89.体育",
    "90.娱乐业",
    "91.中国共产党机关",
    "92.国家机构",
    "93.人民政协、民主党派",
    "94.社会保障",
    "95.群众团体、社会团体和其他成员组织",
    "96.基层群众自治组织",
    "97.国际组织"
]

# ==================== 辅助函数 ====================
def build_prompt(use_text: str) -> tuple:
    """构造 system 和 user 提示词（中文）"""
    system_msg = (
        "你是一名专利技术分类专家，同时精通计算机科学和人工智能技术。你也熟悉《国民经济行业分类》（GB/T 4754-2017）标准。"
        "请根据专利的 USE 字段描述，列出该技术可应用于哪些具体的国民经济行业。"
        "输出一个 JSON 数组，每个元素必须包含：\"code\"（两位数字代码）、\"name\"（行业名称）、\"reasoning\"（判断依据）。"
        "按适用性从高到低排序。"
        "【重要】如果 USE 字段没有说明任何具体的应用场景（例如仅描述技术原理、算法、通用框架，未提及任何行业或用途），则只输出一个元素：{\"code\": \"00\", \"name\": \"未指定\", \"reasoning\": \"USE 未说明具体应用场景\"}。"
    )

    industry_text = "\n".join(INDUSTRY_LIST)
    user_msg = f"""
                # 分类体系（共98类，含00.未指定）
                {industry_text}

                # 分类规则
                1. 利用你的计算机科学和AI知识，分析该技术能解决哪些现实问题。
                2. 优先选择最直接、最典型的应用行业。如果USE描述中出现行业关键词（如“银行”、“医疗”、“自动驾驶”），优先采用。
                3. 如果一项AI技术可用于多个行业（例如图像识别可用于农业、安防、医疗），请列出 USE 中明确提到的所有行业，按相关程度或出现频率排序。
                4. 如果 USE 中完全没有行业或用途描述（只讲算法结构、训练方法、模型改进），请输出 [{{"code":"00","name":"未指定","reasoning":"无具体应用场景"}}]。


                # 示例
                示例1-1（AI + 具体行业）：
                USE: "method for positioning fault of a power grid based on generating adversarial network."
                输出：[{{"code": "44", "name": "电力、热力生产和供应业", "reasoning": "直接用于电网故障定位，属于电力系统运行和维护领域。"}}]

                示例1-2（AI + 具体行业）：
                USE: "hybrid deep learning and gan-supported system for facilitating scalable crop disease detection in an agricultural technology and precision farming field."
                输出：[{{"code": "01", "name": "农业", "reasoning": "直接用于作物病害检测，属于农业生产中的精准农业技术。"}}]


                示例2-1（多个行业，USE 中已列出）：
                USE: "blockchain based large language model question and answer realizing method for use in fields such as finance, supply chain or electronic government."
                输出：[
                  {{"code": "66", "name": "货币金融服务", "reasoning": "可直接用于银行、证券等金融机构的客户服务、风险评估、合规查询等场景。"}},
                  {{"code": "72", "name": "商务服务业", "reasoning": "可用于供应链管理中的信息追溯、供应商查询、合同审核等商务服务环节。"}},
                  {{"code": "92", "name": "国家机构", "reasoning": "可用于电子政务系统中的政策咨询、公文处理、公众服务等政府机构场景。"}}
                ]

                示例2-2（多个行业，USE 中已列出）：
                USE: "method for assisting artificial intelligent medicine administration based on large language model by an artificial intelligent medicine administration auxiliary device(claimed) for medical assistant work in a medical field such as medicine consultation, health consultation and medical research and education."
                输出：[
                  {{"code": "84", "name": "卫生", "reasoning": "直接用于医疗咨询、健康咨询等医疗辅助工作"}},
                  {{"code": "83", "name": "教育", "reasoning": "用于医学研究和教育中的辅助教学与培训"}},
                  {{"code": "73", "name": "研究和试验发展", "reasoning": "支持医学研究中的数据分析与知识辅助"}}
                ]


                示例3-1（仅技术描述，无特定行业）：
                USE: "method for converting natural language query to standard query using sequence-to-sequence neural network."
                输出：[{{"code": "00", "name": "未指定", "reasoning": "USE 未说明具体应用场景"}}]

                示例3-2（仅技术描述，无特定行业）：
                USE: "method for predicting token given context text and vocabulary text, including circumstances in which token is in vocabulary text and not in context text."
                输出：[{{"code": "00", "name": "未指定", "reasoning": "USE 未说明具体应用场景"}}]


                # 输入
                专利 USE 字段内容：
                \"\"\"
                {use_text}
                \"\"\"

                请输出 JSON 数组：
                """
    return system_msg, user_msg

def call_deepseek_api(use_text: str, client: OpenAI) -> list:
    """调用 DeepSeek API 并解析返回的 JSON 数组，失败时返回 [{"code":"00",...}]"""
    system_msg, user_msg = build_prompt(use_text)
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg}
            ],
            temperature=0.1,
            max_tokens=2000,
            top_p=0.9,
            seed=42,
            stream=False
        )
        content = response.choices[0].message.content
        # 提取 JSON 数组
        cleaned = re.sub(r"```json\s*|\s*```", "", content.strip())
        start = cleaned.find('[')
        end = cleaned.rfind(']')
        if start == -1 or end == -1:
            raise ValueError("未找到 JSON 数组")
        json_str = cleaned[start:end+1]
        data = json.loads(json_str)
        if isinstance(data, list):
            return data
        else:
            raise ValueError("输出不是列表")
    except Exception as e:
        print(f"API 调用失败: {e}")
    # 失败后返回未指定
    return [{"code": "98", "name": "调用失败", "reasoning": "API 调用失败"}]

def save_checkpoint(results, next_idx, output_file):
    """保存当前结果到 Excel，并保存检查点文件（下一个要处理的相对索引）"""
    out_df = pd.DataFrame(results)
    out_df.to_excel(output_file, index=False)
    checkpoint_file = output_file + ".checkpoint"
    with open(checkpoint_file, "w") as f:
        f.write(str(next_idx))
    print(f"已保存检查点：已完成 {next_idx} 行，结果已保存至 {output_file}")

def load_checkpoint(output_file):
    """如果存在检查点，返回下一个要处理的相对索引；否则返回0"""
    checkpoint_file = output_file + ".checkpoint"
    if os.path.exists(checkpoint_file):
        with open(checkpoint_file, "r") as f:
            return int(f.read().strip())
    return 0

def load_existing_results(output_file):
    """如果输出文件存在，加载已有的结果列表；否则返回空列表"""
    if os.path.exists(output_file):
        try:
            df_existing = pd.read_excel(output_file)
            return df_existing.to_dict(orient="records")
        except Exception as e:
            print(f"加载已有结果失败: {e}，将重新开始。")
    return []

# ==================== 主程序 ====================
def main():
    # 1. 读取 Excel
    print(f"正在读取 Excel 文件：{INPUT_EXCEL}")
    df = pd.read_excel(INPUT_EXCEL)
    total_rows = len(df)
    print(f"文件总行数：{total_rows}")

    # 2. 确定行范围（转换为 0-index）
    start_idx = start_line - 1 if start_line is not None else 0
    if end_line is None or end_line > total_rows:
        end_idx = total_rows
    else:
        end_idx = end_line - 1
    start_idx = max(0, start_idx)
    end_idx = min(total_rows, end_idx)
    if start_idx >= end_idx:
        print(f"无效范围：start_line={start_line}, end_line={end_line}，没有需要处理的行。")
        return

    print(f"将处理第 {start_idx+1} 行到第 {end_idx} 行（索引 {start_idx} 到 {end_idx-1}）")
    # 切片选取子 DataFrame
    df_sub = df.iloc[start_idx:end_idx].copy()
    total_to_process = len(df_sub)
    print(f"待处理行数：{total_to_process}")

    # 3. 断点续跑：加载已有的结果和检查点
    all_results = []
    next_idx = 0  # 已经处理的行数（相对于 df_sub 的开始索引）
    
    checkpoint = load_checkpoint(OUTPUT_EXCEL)
    if checkpoint > 0:
        if checkpoint < total_to_process:
            print(f"从检查点恢复：已完成 {checkpoint} 行")
            all_results = load_existing_results(OUTPUT_EXCEL)
            if not all_results:
                all_results = []
            # 跳过已处理的行
            df_sub = df_sub.iloc[checkpoint:].copy()
            total_remaining = len(df_sub)
            next_idx = checkpoint
            print(f"剩余待处理行数：{total_remaining}")
        else:
            print("范围内的所有行已处理完毕。退出。")
            return
    else:
        # 无检查点，检查是否有旧结果文件
        if os.path.exists(OUTPUT_EXCEL):
            print("发现已有结果文件但无检查点，将从头开始（旧文件将被覆盖）。")
        all_results = []
        df_sub = df_sub.copy()
        total_remaining = total_to_process
        next_idx = 0
        print(f"全新开始，待处理行数：{total_remaining}")

    if total_remaining == 0:
        print("没有需要处理的行。退出。")
        return

    # 4. 初始化 DeepSeek 客户端
    client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

    # 5. 逐行处理
    processed_in_this_run = 0
    for rel_idx, (original_idx, row) in enumerate(df_sub.iterrows()):
        current_row_num = original_idx + 1
        remaining = total_remaining - rel_idx - 1
        print(f"正在处理第 {current_row_num} 行，剩余 {remaining} 行")

        patent_id = row[COL_PATENT_ID]
        use_text = str(row[COL_USE]) if pd.notna(row[COL_USE]) else ""

        if not use_text.strip():
            # USE 字段为空
            all_results.append({
                "patent_id": patent_id,
                "industry_code": "00",
                "industry_name": "未指定",
                "reasoning": "USE 字段为空",
                "order_rank": 1
            })
        else:
            industries = call_deepseek_api(use_text, client)
            for rank, ind in enumerate(industries[:5], start=1):
                all_results.append({
                    "patent_id": patent_id,
                    "industry_code": ind.get("code", ""),
                    "industry_name": ind.get("name", ""),
                    "reasoning": ind.get("reasoning", ""),
                    "order_rank": rank
                })

        processed_in_this_run += 1
        total_processed = next_idx + processed_in_this_run

        # 每 SAVE_INTERVAL 行保存一次中间结果
        if processed_in_this_run % SAVE_INTERVAL == 0:
            save_checkpoint(all_results, total_processed, OUTPUT_EXCEL)

        time.sleep(0.5)  # 避免限流

    # 6. 最终保存
    save_checkpoint(all_results, total_processed, OUTPUT_EXCEL)
    print(f"处理完成！共处理 {total_processed} 件专利，生成 {len(all_results)} 条行业映射记录。")
    print(f"最终结果保存至：{OUTPUT_EXCEL}")

    # 清理检查点文件
    checkpoint_file = OUTPUT_EXCEL + ".checkpoint"
    if os.path.exists(checkpoint_file):
        os.remove(checkpoint_file)
        print("已删除检查点文件。")

if __name__ == "__main__":
    main()