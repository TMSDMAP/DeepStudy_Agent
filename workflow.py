from langgraph.graph import END, StateGraph
from langgraph.prebuilt import ToolNode, tools_condition

from agents import (
    activity_designer_node,
    content_creator_node,
    evaluation_expert_node,
    evaluator_tools,
    pedagogy_planner_node,
    quiz_designer_node,
    subject_expert_node,
    tools,
)
from state import LessonState


MAX_REWRITE_ROUNDS = 5


def should_continue(state: LessonState):
    if state.get("is_approved") or state.get("iteration_count", 0) >= MAX_REWRITE_ROUNDS:
        return "end"
    target = state.get("repair_target") or "content"
    if target == "quiz" and not state.get("include_quiz", True):
        return "content"
    if target in {"subject", "pedagogy", "activity", "quiz", "content"}:
        return target
    return "content"


def eval_should_continue(state: LessonState):
    """判断质检人员是想调用代码执行工具，还是给出了审核结论。"""
    messages = state.get("messages") or []
    if messages:
        last_message = messages[-1]
        if hasattr(last_message, "tool_calls") and last_message.tool_calls:
            return "tools"
    return should_continue(state)


workflow = StateGraph(LessonState)

# 实例化工具节点
search_tool_node = ToolNode(tools)
eval_tool_node = ToolNode(evaluator_tools)

# 添加节点
workflow.add_node("SubjectExpert", subject_expert_node)
workflow.add_node("Tools", search_tool_node)
workflow.add_node("PedagogyPlanner", pedagogy_planner_node)
workflow.add_node("ActivityDesigner", activity_designer_node)
workflow.add_node("QuizDesigner", quiz_designer_node)
workflow.add_node("ContentCreator", content_creator_node)
workflow.add_node("EvaluationExpert", evaluation_expert_node)
workflow.add_node("EvalTools", eval_tool_node)

# 设置起点
workflow.set_entry_point("SubjectExpert")

# 教研专家可调用一次搜索工具
workflow.add_conditional_edges(
    "SubjectExpert",
    tools_condition,
    {
        "tools": "Tools",
        END: "PedagogyPlanner",
    },
)

workflow.add_edge("Tools", "SubjectExpert")
workflow.add_edge("PedagogyPlanner", "ActivityDesigner")
workflow.add_edge("ActivityDesigner", "QuizDesigner")
workflow.add_edge("QuizDesigner", "ContentCreator")

# 常规流程
workflow.add_edge("ContentCreator", "EvaluationExpert")

# 质检专家可调用代码执行工具
workflow.add_conditional_edges(
    "EvaluationExpert",
    eval_should_continue,
    {
        "tools": "EvalTools",
        "end": END,
        "subject": "SubjectExpert",
        "pedagogy": "PedagogyPlanner",
        "activity": "ActivityDesigner",
        "quiz": "QuizDesigner",
        "content": "ContentCreator",
    },
)

workflow.add_edge("EvalTools", "EvaluationExpert")

app_workflow = workflow.compile()
