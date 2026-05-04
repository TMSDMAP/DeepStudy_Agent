from typing import TypedDict, Annotated
from langgraph.graph.message import add_messages

class LessonState(TypedDict):
    topic: str              # 用户输入的教学主题
    messages: Annotated[list, add_messages] # 支持工具调用的消息历史
    include_code: bool      # 是否包含示例代码章节
    include_quiz: bool      # 是否包含课后题与答案章节
    teaching_traits: str    # 教师自定义教学风格与约束
    external_knowledge: str # 教师上传资料提取出的外部知识文本
    knowledge_sources: list[str] # 外部资料文件名列表
    outline: str            # 学科专家生成的大纲
    pedagogy_plan: str      # 教学设计 Agent 产出的课堂策略
    activity_plan: str      # 互动活动 Agent 产出的活动方案
    quiz_bank: str          # 题库 Agent 产出的题目与答案草案
    draft_content: str      # 内容创作者生成的草稿
    feedback: str           # 评估专家的反馈意见
    repair_target: str      # 质检指定返工对象：subject/pedagogy/activity/quiz/content
    feedback_category: str  # 质检问题类型，供前端展示与路由判断
    is_approved: bool       # 是否通过评估
    final_content: str      # 最终版 Markdown 讲义
    iteration_count: int    # 记录重试次数
