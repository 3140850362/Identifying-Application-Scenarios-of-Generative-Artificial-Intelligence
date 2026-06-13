import json
import time
import pandas as pd
from openai import OpenAI
from tqdm import tqdm

# ==================== 配置 ====================
API_KEY = ""         # 替换为你的 API Key
MODEL = "deepseek-v4-flash" 
BASE_URL = "https://api.deepseek.com"


INPUT_FILE = "标题+摘要-82326.xlsx"
OUTPUT_FILE = "patents_classified.xlsx"
TITLE_COL = "标题"
ABSTRACT_COL = "摘要"

df = pd.read_excel(INPUT_FILE)

m = 1
n = 82326
df = df.iloc[m:n,:]

# ==================== 提示词 ====================
SYSTEM_PROMPT = """
    # 角色
    你是一位精通计算机科学、人工智能技术及专利分析的知识产权专家。你的任务是基于专利的标题与摘要，精确区分真正的生成式人工智能专利与虚假的噪声专利。

    # 判定标准
    若专利满足以下主要特征之一，则判定为"是"：
       - 将生成式AI、大语言模型和相关工具如：ChatGPT、DeepSeek、Gemini、Claude、Bard、Grok、Copilot、Qwen、Ernie、Doubao、Kimi、GLM、Llama、Perplexity、DALL-E、Midjourney、Sora等应用于特定领域。
       - 涉及生成模型和大语言模型的底层架构。
       - 涉及生成式AI和大语言模型的训练方法。
       - 涉及生成式AI或大语言模型的微调、对齐或适配等方法。
       - 涉及生成式AI或大语言模型的推理优化、部署或压缩技术。
       - 涉及生成式AI或大语言模型的多模态融合或跨模态生成。
       - 涉及生成式AI或大语言模型的评估、测试或基准方法。
       - 涉及生成式AI或大语言模型的特定任务智能体（Agent）或工作流编排。
       - 涉及生成式AI或大语言模型的提示工程（Prompt Engineering）或上下文学习方法。
       - 涉及生成式AI或大语言模型的数据构建、预处理或增强方法。
       - 涉及生成式AI或大语言模型的输出解码、控制或引导方法。
       - 涉及生成式AI或大语言模型的记忆、检索或外部知识融合机制。
       - 涉及生成式AI或大语言模型的安全防护、内容风控、隐私保护与合规技术。
       - 涉及生成式AI或大语言模型的人机交互、对话交互与交互体验优化技术。
       - 涉及生成式AI或大语言模型的迁移学习与模型复用技术。
       - 涉及生成式AI或大语言模型的错误修正、幻觉抑制与内容保真技术。
       - 涉及生成式AI或大语言模型的插件拓展、功能拓展与接口适配技术。
       - 涉及生成式AI或大语言模型的算力调度、资源分配与集群运行技术。
       - 涉及生成式AI或大语言模型的语音、文本、图像、视频等单一模态生成优化技术。
       - 涉及生成式AI或大语言模型的知识图谱结合、逻辑推理与因果分析技术。
       - 涉及生成式AI或大语言模型的多模型协同、模型组合与联动调用技术。
       - 涉及生成式AI或大语言模型的权限管理、访问控制与使用监管技术。
       - 涉及生成式AI或大语言模型的版本迭代、模型更新与增量学习技术。
       - 涉及生成式AI或大语言模型的模型可解释性与透明度技术。
       - 涉及生成式AI或大语言模型的模型水印、指纹识别、溯源取证与知识产权保护技术。
       - 涉及生成式AI或大语言模型的对抗鲁棒性增强、后门防御、输入净化与异常检测技术。
       - 涉及生成式AI或大语言模型的模型量化、剪枝、知识蒸馏等轻量化与推理加速技术。
       - 涉及生成式AI或大语言模型的长文本建模、上下文窗口外推与高效注意力机制。
       - 涉及生成式AI或大语言模型的低资源场景适配、小样本 / 零样本学习技术。
       - 涉及生成式AI或大语言模型的实时流式生成、低延迟响应技术。
       - 涉及生成式AI或大语言模型的终端侧部署、端边云协同运行技术。
       - 涉及生成式AI或大语言模型的其他新兴或前沿技术。

       
    # 分析流程
    1. 通读标题与摘要，识别专利核心内容。
    2. 检查其是否包含上述"主要特征"所列举的某一方面。
    3. 给出最终判断，并附上一句关键的技术性理由。

    # 输出格式
    请严格按照以下JSON格式输出，不要添加任何额外的解释或Markdown代码块标记：
    {
      "is_generative_ai": "是/否",
      "reason": "用中文撰写的一句话核心理由，精确指出专利中的生成技术或说明为何不是生成式AI。"
    }"""

# ==================== 初始化 ====================
client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

df["is_generative_ai"] = None
df["judge_reason"] = None
print(f"共 {len(df)} 条数据")

# ==================== 逐条处理 ====================
# for idx in tqdm(df.index, desc="处理进度"):
for idx in range(m, n):
    print('第'+str(idx+1)+'件专利，还剩：'+str(n-idx)+'件专利---------------------------------------'+str(idx+1))

    try:
	    title = df.at[idx, TITLE_COL]
	    abstract = df.at[idx, ABSTRACT_COL]

	    user_msg = f"# 待分析专利信息\n标题：{title}\n摘要：{abstract}"


	    response = client.chat.completions.create(
	        model=MODEL,
	        messages=[
	            {"role": "system", "content": SYSTEM_PROMPT},
	            {"role": "user", "content": user_msg}
	        ],
	        temperature=0.1,
	        max_tokens=2000,
	        top_p=0.9,
	        seed = 42,
	        stream=False
	    )

	    result = json.loads(response.choices[0].message.content)
	    df.at[idx, "is_generative_ai"] = result["is_generative_ai"]
	    df.at[idx, "judge_reason"] = result["reason"]

	    time.sleep(0.5)

    except:
	    df.at[idx, "is_generative_ai"] = '出现错误'
	    df.at[idx, "judge_reason"] = '出现错误'
	    print('第'+str(idx+1)+'件专利出现错误++++++++++++++++++++，还剩：'+str(n-idx)+'件专利---------------------------------------'+str(idx+1))



# ==================== 保存 ====================
df.to_excel(OUTPUT_FILE, index=False)

yes = (df["is_generative_ai"] == "是").sum()
no = (df["is_generative_ai"] == "否").sum()
print(f"生成式AI: {yes}")
print(f"非生成式: {no}")
print(f"结果已保存至: {OUTPUT_FILE}")