import contextlib
import io
import os
import re
import sys
from pathlib import Path
from typing import Any, Mapping

from dotenv import load_dotenv
from langchain_core.messages import AIMessage, BaseMessage
from langchain_core.outputs import ChatResult
from langchain_community.tools import DuckDuckGoSearchRun
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.tools import tool
from langchain_openai import ChatOpenAI
from langchain_openai.chat_models.base import (
    _convert_dict_to_message,
    _convert_from_v1_to_chat_completions,
    _convert_message_to_dict,
    _create_usage_metadata,
)

from state import LessonState

def _load_runtime_env() -> None:
    """Load .env from common packaged/runtime locations."""
    candidates = [
        Path.cwd() / ".env",
        Path(__file__).resolve().parent / ".env",
    ]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / ".env")

    seen = set()
    for path in candidates:
        resolved = str(path.resolve()) if path.exists() else ""
        if not resolved or resolved in seen:
            continue
        seen.add(resolved)
        load_dotenv(dotenv_path=path, override=False)


# 加载 .env 文件中的 API Key
_load_runtime_env()
if os.getenv("LLM_API_KEY") and not os.getenv("OPENAI_API_KEY"):
    os.environ["OPENAI_API_KEY"] = os.getenv("LLM_API_KEY", "")
if os.getenv("LLM_API_BASE") and not os.getenv("OPENAI_API_BASE"):
    os.environ["OPENAI_API_BASE"] = os.getenv("LLM_API_BASE", "")


def _env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on", "enabled"}


def _env_int(name: str) -> int | None:
    value = (os.getenv(name) or "").strip()
    if not value:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _env_float(name: str) -> float | None:
    value = (os.getenv(name) or "").strip()
    if not value:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _thinking_extra_body() -> dict[str, Any] | None:
    if not _env_bool("LLM_THINKING_ENABLED", default=False):
        return None
    thinking_type = (os.getenv("LLM_THINKING_TYPE") or "enabled").strip() or "enabled"
    return {"thinking": {"type": thinking_type}}


def _reasoning_effort() -> str | None:
    value = (os.getenv("LLM_REASONING_EFFORT") or "").strip()
    return value or None


def _assistant_message_to_deepseek_dict(message: BaseMessage) -> dict[str, Any]:
    if isinstance(message, AIMessage):
        msg = _convert_message_to_dict(_convert_from_v1_to_chat_completions(message))
        reasoning_content = message.additional_kwargs.get("reasoning_content")
        if reasoning_content:
            msg["reasoning_content"] = reasoning_content
        return msg
    return _convert_message_to_dict(message)


class ChatDeepSeekCompat(ChatOpenAI):
    """Preserve DeepSeek thinking-mode fields across tool-call rounds."""

    def _get_request_payload(
        self,
        input_: Any,
        *,
        stop: list[str] | None = None,
        **kwargs: Any,
    ) -> dict[str, Any]:
        messages = self._convert_input(input_).to_messages()
        if stop is not None:
            kwargs["stop"] = stop

        payload = {**self._default_params, **kwargs}
        payload["messages"] = [_assistant_message_to_deepseek_dict(m) for m in messages]
        return payload

    def _create_chat_result(
        self,
        response: dict | Any,
        generation_info: dict | None = None,
    ) -> ChatResult:
        chat_result = super()._create_chat_result(response, generation_info)
        response_dict = (
            response
            if isinstance(response, dict)
            else response.model_dump(exclude={"choices": {"__all__": {"message": {"parsed"}}}})
        )
        choices = response_dict.get("choices") or []
        for generation, res in zip(chat_result.generations, choices, strict=False):
            message = generation.message
            if not isinstance(message, AIMessage):
                continue
            raw_message: Mapping[str, Any] = res.get("message") or {}
            reasoning_content = raw_message.get("reasoning_content")
            if reasoning_content:
                message.additional_kwargs["reasoning_content"] = reasoning_content
            token_usage = response_dict.get("usage")
            service_tier = response_dict.get("service_tier")
            if token_usage:
                message.usage_metadata = _create_usage_metadata(token_usage, service_tier)
        return chat_result


def _strip_internal_tool_markup(text: str) -> str:
    """Remove internal function-call markers accidentally leaked by the model."""
    if not text:
        return ""
    text = re.sub(r"</?[^>\n]*DSML[^>\n]*>", "", text, flags=re.I)
    cleaned_lines = []
    for line in text.splitlines():
        if "DSML" in line.upper():
            continue
        cleaned_lines.append(line)
    return "\n".join(cleaned_lines).strip()


def _count_tool_calls(messages: list, tool_name: str) -> int:
    count = 0
    for msg in messages:
        tool_calls = getattr(msg, "tool_calls", None) or []
        for call in tool_calls:
            if call.get("name") == tool_name:
                count += 1
    return count


def _tail_tool_history(messages: list) -> list:
    history = []
    for msg in reversed(messages):
        msg_type = getattr(msg, "type", "")
        has_calls = bool(getattr(msg, "tool_calls", None))
        if msg_type == "tool" or (msg_type == "ai" and has_calls):
            history.insert(0, msg)
        else:
            break
    return history


def _extract_topic_scope(topic: str) -> str:
    if not topic:
        return ""
    text = topic.strip()
    if re.search(r"(最后一回|末回|终回|第一百二十回|第120回)", text):
        return "最后一回（第一百二十回）"
    match = re.search(r"第\s*([一二三四五六七八九十百零两\d]+)\s*回", text)
    if match:
        return f"第{match.group(1)}回"
    return ""


def _hard_scope_violation(scope: str, draft: str) -> str:
    if not scope or not draft:
        return ""
    text = draft or ""

    if scope == "最后一回（第一百二十回）":
        first_hits = len(re.findall(r"第一回", text))
        last_hits = len(re.findall(r"(最后一回|第一百二十回|第120回|一百二十回)", text))
        title_line = re.search(r"(?m)^#\s*(.+)$", text)
        title_mismatch = bool(title_line and "第一回" in title_line.group(1) and last_hits == 0)
        if title_mismatch or (first_hits >= 2 and last_hits == 0):
            return (
                "范围硬约束未满足：主题要求“最后一回（第一百二十回）”，"
                "但草稿主要内容仍落在“第一回”。请整体重写并严格聚焦最后一回。"
            )
        return ""

    if scope not in text:
        wrong_chapter_in_title = re.search(
            r"(?m)^#\s*.*第[一二三四五六七八九十百零两\d]+回",
            text,
        )
        if wrong_chapter_in_title:
            return (
                f"范围硬约束未满足：主题要求“{scope}”，但标题章节不一致。"
                "请按目标回目重写内容。"
            )
    return ""


def _sanitize_feedback_for_regen(feedback: str) -> str:
    cleaned = _strip_internal_tool_markup(feedback or "")
    if not cleaned:
        return ""
    # 关键硬校验错误需要尽量保留细节，避免下一轮修复丢失定位信息。
    if "示例代码第" in cleaned and "语法错误" in cleaned:
        return cleaned[:2800].strip()
    if "范围硬约束未满足" in cleaned:
        return cleaned[:2200].strip()
    # 反馈过长会污染下一轮提示词，常规错误做截断。
    return cleaned[:1600].strip()


def _targeted_feedback(state: LessonState, target: str) -> str:
    """Return feedback only when the evaluator routed repair to this agent."""
    if state.get("repair_target") != target:
        return ""
    return _sanitize_feedback_for_regen(state.get("feedback", ""))


def _extract_python_code_blocks(text: str) -> list[str]:
    blocks = re.findall(r"```(?:python|py)\s*(.*?)```", text or "", flags=re.I | re.S)
    return [block.strip() for block in blocks if block and block.strip()]


def _clean_code_for_compile(code: str) -> str:
    cleaned = (code or "").replace("\ufeff", "")
    for ch in ("\u200b", "\u200c", "\u200d", "\u2060"):
        cleaned = cleaned.replace(ch, "")
    return cleaned


def _remove_invisible_chars(text: str) -> str:
    if not text:
        return ""
    cleaned = text
    for ch in ("\ufeff", "\u200b", "\u200c", "\u200d", "\u2060"):
        cleaned = cleaned.replace(ch, "")
    return cleaned


def _strip_quiz_inline_markup(text: str) -> str:
    if not text:
        return ""
    return text.replace("**", "").replace("__", "").replace("`", "").strip()


def _normalize_quiz_sections_for_validation(text: str) -> str:
    """Normalize quiz/answer headings and line formats before validation."""
    lines = _remove_invisible_chars(text).splitlines()
    normalized: list[str] = []
    in_quiz_section = False
    in_answer_section = False
    answer_index = 1
    question_index = 1
    pending_answer_letter: str | None = None

    for raw in lines:
        line = _remove_invisible_chars(raw.rstrip())
        stripped = line.strip()
        match_text = _strip_quiz_inline_markup(stripped)
        heading_text = re.sub(r"^#{1,6}\s*", "", stripped).strip()

        if (
            stripped.startswith("##")
            or re.match(r"^[一二三四五六七八九十七八\d]+[、.．]", stripped)
        ) and re.search(r"(课后选择题|练习题目)", heading_text):
            in_quiz_section = True
            in_answer_section = False
            normalized.append("## 七、课后选择题（至少4题）")
            continue

        if (
            stripped.startswith("##")
            or re.match(r"^[一二三四五六七八九十七八\d]+[、.．]", stripped)
        ) and re.search(r"答案", heading_text):
            in_quiz_section = False
            in_answer_section = True
            answer_index = 1
            pending_answer_letter = None
            normalized.append("## 八、答案与解析")
            continue

        if stripped.startswith("## ") and not re.search(r"(课后选择题|练习题目|答案)", stripped):
            in_quiz_section = False
            in_answer_section = False

        if in_quiz_section and re.fullmatch(r"\d+", match_text):
            continue

        if in_quiz_section and re.match(r"^\d+\.\s+", match_text):
            q_num = re.match(r"^(\d+)\.\s+", match_text)
            if q_num:
                question_index = int(q_num.group(1)) + 1
            if all(re.search(rf"{opt}[\.．:：]\s*", match_text) for opt in ["A", "B", "C", "D"]):
                first_opt = re.search(r"[ABCD][\.．:：]\s*", match_text)
                stem = match_text[:first_opt.start()].strip() if first_opt else match_text
                normalized.append(stem)
                option_chunks = re.findall(
                    r"([ABCD])[\.．:：]\s*(.*?)(?=\s+[ABCD][\.．:：]\s*|$)",
                    match_text[first_opt.start():] if first_opt else "",
                )
                for opt, body in option_chunks:
                    normalized.append(f"{opt}. {body.strip()}")
                continue
            normalized.append(match_text)
            continue

        if in_quiz_section and all(re.search(rf"{opt}[\.．:：]\s*", match_text) for opt in ["A", "B", "C", "D"]):
            first_opt = re.search(r"[ABCD][\.．:：]\s*", match_text)
            stem = match_text[:first_opt.start()].strip() if first_opt else match_text
            if stem:
                normalized.append(f"{question_index}. {stem}")
                question_index += 1
            option_chunks = re.findall(
                r"([ABCD])[\.．:：]\s*(.*?)(?=\s+[ABCD][\.．:：]\s*|$)",
                match_text[first_opt.start():] if first_opt else "",
            )
            for opt, body in option_chunks:
                normalized.append(f"{opt}. {body.strip()}")
            continue

        if in_quiz_section and re.match(r"^[\-*]\s*[ABCD][\.．:：]\s+", match_text):
            opt_line = re.sub(r"^[\-*]\s*", "", match_text)
            opt_match = re.match(r"^([ABCD])[\.．:：]\s*(.*)$", opt_line)
            if opt_match:
                normalized.append(f"{opt_match.group(1)}. {opt_match.group(2).strip()}")
            else:
                normalized.append(opt_line)
            continue

        if in_quiz_section and re.match(r"^[ABCD][\.．:：]\s+", match_text):
            opt_match = re.match(r"^([ABCD])[\.．:：]\s*(.*)$", match_text)
            if opt_match:
                normalized.append(f"{opt_match.group(1)}. {opt_match.group(2).strip()}")
            else:
                normalized.append(match_text)
            continue

        if in_answer_section:
            single_option = re.fullmatch(r"[A-Da-d]", match_text)
            numbered_option = re.fullmatch(r"(\d+)\.\s*([A-Da-d])", match_text)
            answer_with_explain = re.fullmatch(r"([A-Da-d])[\.．:：]\s*(.+)", match_text)
            numbered_with_explain = re.fullmatch(r"(\d+)\.\s*([A-Da-d])[\.．:：]\s*(.+)", match_text)
            correct_with_explain = re.fullmatch(
                r"(?:正确答案|答案)\s*[：:]\s*([A-Da-d])\s*(?:[，,；;。]?\s*)?(?:解析)\s*[：:]\s*(.+)",
                match_text,
            )
            answer_only = re.fullmatch(r"(?:正确答案|答案)\s*[：:]\s*([A-Da-d])\s*", match_text)
            explain_only = re.fullmatch(r"(?:解析)\s*[：:]\s*(.+)", match_text)
            if single_option:
                normalized.append(f"{answer_index}. {match_text.upper()}")
                answer_index += 1
                pending_answer_letter = None
                continue
            if correct_with_explain:
                letter = correct_with_explain.group(1).upper()
                explain = correct_with_explain.group(2).strip()
                normalized.append(f"{answer_index}. {letter}：{explain}")
                answer_index += 1
                pending_answer_letter = None
                continue
            if numbered_option:
                normalized.append(f"{numbered_option.group(1)}. {numbered_option.group(2).upper()}")
                answer_index = int(numbered_option.group(1)) + 1
                pending_answer_letter = None
                continue
            if answer_with_explain:
                letter = answer_with_explain.group(1).upper()
                explain = answer_with_explain.group(2).strip()
                normalized.append(f"{answer_index}. {letter}：{explain}")
                answer_index += 1
                pending_answer_letter = None
                continue
            if numbered_with_explain:
                num = int(numbered_with_explain.group(1))
                letter = numbered_with_explain.group(2).upper()
                explain = numbered_with_explain.group(3).strip()
                normalized.append(f"{num}. {letter}：{explain}")
                answer_index = num + 1
                pending_answer_letter = None
                continue
            if answer_only:
                pending_answer_letter = answer_only.group(1).upper()
                continue
            if explain_only and pending_answer_letter:
                normalized.append(f"{answer_index}. {pending_answer_letter}：{explain_only.group(1).strip()}")
                answer_index += 1
                pending_answer_letter = None
                continue

        normalized.append(match_text if in_answer_section else line)

    return "\n".join(normalized).strip()


def _draft_quality_feedback(draft: str, include_code: bool, include_quiz: bool) -> str:
    text = _normalize_quiz_sections_for_validation(_strip_internal_tool_markup(draft or ""))
    if not text:
        return "本地硬校验未通过：讲义内容为空。请重新生成完整讲义。"

    required = [
        r"##\s*.*(学习目标|教学目标)",
        r"##\s*.*(先修知识|学情|学习者分析)",
        r"##\s*.*(课堂流程|时间分配|教学过程|教学环节)",
        r"##\s*.*(核心知识|知识点|重点内容|教学内容)",
        r"##\s*.*(互动活动|课堂活动|师生互动|分组讨论)",
    ]
    missing = [pat for pat in required if not re.search(pat, text)]
    if len(missing) >= 3:
        return "本地硬校验未通过：主体章节缺失过多，至少需要学习目标、学情/先修、课堂流程、核心知识、互动活动中的大部分内容。请完整重写。"

    if include_quiz:
        quiz_or_answer_sections = re.findall(
            r"(?ms)^##\s*(?:七、课后选择题（至少4题）|八、答案与解析)\s*.*?(?=^##\s|\Z)",
            text,
        )
        quiz_or_answer_text = "\n\n".join(quiz_or_answer_sections)
        if quiz_or_answer_text and re.search(r"(?i)(待补充答案|请补充|TODO|# TODO)", quiz_or_answer_text):
            return "本地硬校验未通过：课后题或答案解析仍包含待补充/TODO/占位内容。请题库 Agent 重新生成完整题目、选项、答案与解析。"

    if re.search(r"(?i)(待补充答案|请补充|TODO|# TODO)", text):
        return "本地硬校验未通过：讲义仍包含待补充/TODO/占位内容。请替换为真实教学内容。"

    code_blocks = _extract_python_code_blocks(text)
    if include_code and code_blocks:
        for idx, code in enumerate(code_blocks, start=1):
            try:
                compile(_clean_code_for_compile(code), f"<lesson_code_{idx}>", "exec")
            except SyntaxError as exc:
                line_text = (exc.text or "").strip()
                line_detail = f"（第{exc.lineno}行：{exc.msg}）"
                if line_text:
                    line_detail += f" -> {line_text}"
                return f"本地硬校验未通过：示例代码第 {idx} 段存在语法错误{line_detail}。请修复后重新输出完整讲义。"

    if include_quiz:
        quiz_match = re.search(
            r"(?ms)^##\s*七、课后选择题（至少4题）\s*(.*?)(?=^##\s*八、答案与解析|\Z)",
            text,
        )
        answer_match = re.search(
            r"(?ms)^##\s*八、答案与解析\s*(.*?)(?=^##\s|\Z)",
            text,
        )
        quiz_text = quiz_match.group(1).strip() if quiz_match else ""
        answer_text = answer_match.group(1).strip() if answer_match else ""
        question_count = len(re.findall(r"(?m)^\d+\.\s+", quiz_text))
        if question_count == 0:
            question_count = len(
                re.findall(
                    r"(?m)^.*A[\.．:：].*B[\.．:：].*C[\.．:：].*D[\.．:：].*$",
                    quiz_text,
                )
            )
        answer_count = len(
            re.findall(
                r"(?m)^\s*(?:\d+\.\s*)?(?:[A-D][：:]\s*|(?:正确答案|答案)\s*[：:]\s*[A-D])",
                answer_text,
            )
        )
        if (quiz_match or answer_match) and (question_count < 4 or answer_count < 4):
            return (
                "本地硬校验未通过：已输出课后题或答案章节，但题目与答案解析不足 4 条。"
                f"当前识别到题目 {question_count} 条、答案解析 {answer_count} 条。"
                "请题库 Agent 按固定格式重新生成完整题库："
                "每题用“1.”编号；A、B、C、D 四个选项各占一行；"
                "答案区使用“1. B：解析文字”格式，且与题目一一对应。"
            )

    return ""


def _classify_repair_target(feedback: str, include_quiz: bool) -> tuple[str, str]:
    """Map evaluator feedback to the agent that should repair the source artifact."""
    text = feedback or ""
    if "范围硬约束" in text or re.search(r"(主题|范围|回目|章节).*(不一致|跑偏|不合格)", text):
        return "subject", "主题范围"
    if include_quiz and re.search(r"(题目|题干|选项|答案|解析|课后题|选择题|题库)", text):
        return "quiz", "题库与答案"
    if re.search(r"(讲义内容为空|主体章节缺失|示例代码|待补充|TODO|占位)", text, flags=re.I):
        return "content", "总编整合"
    if re.search(r"(互动活动|课堂活动|分组|形成性评价|学生产出|评价方式)", text):
        return "activity", "互动活动"
    if re.search(r"(学情|先修|教学目标|课堂流程|时间分配|教学过程|分层教学)", text):
        return "pedagogy", "教学设计"
    return "content", "总编整合"


@tool
def search_web(query: str) -> str:
    """当你不确定最新知识点或缺乏相关教学案例时，使用此工具搜索网络。"""
    try:
        search = DuckDuckGoSearchRun()
        return search.run(query)
    except Exception as e:
        return f"网络搜索工具当前不可用或连接超时，请你基于自己的知识库完成编写（错误信息: {str(e)}）。"


@tool
def run_python_code(code: str) -> str:
    """运行草稿中的 Python 代码并返回输出或报错，用于教学质检。"""
    output = io.StringIO()
    try:
        with contextlib.redirect_stdout(output), contextlib.redirect_stderr(output):
            exec(code, {})
        result = output.getvalue().strip()
        return result if result else "代码执行成功，无任何输出。"
    except Exception:
        import traceback

        return f"代码抛出异常:\n{traceback.format_exc()}"


tools = [search_web]
evaluator_tools = [run_python_code]

_thinking_enabled = _env_bool("LLM_THINKING_ENABLED", default=False)
_temperature = None if _thinking_enabled else (_env_float("LLM_TEMPERATURE") or 0.5)

llm = ChatDeepSeekCompat(
    model=os.getenv("LLM_MODEL", "deepseek-chat"),
    temperature=_temperature,
    max_tokens=_env_int("LLM_MAX_TOKENS"),
    base_url=os.getenv("LLM_API_BASE") or os.getenv("OPENAI_API_BASE"),
    api_key=os.getenv("LLM_API_KEY") or os.getenv("OPENAI_API_KEY"),
    extra_body=_thinking_extra_body(),
    reasoning_effort=_reasoning_effort(),
    use_responses_api=False,
)

llm_with_tools = llm.bind_tools(tools)
llm_with_eval_tools = llm.bind_tools(evaluator_tools)


def subject_expert_node(state: LessonState):
    messages = state.get("messages", [])
    scope = _extract_topic_scope(state.get("topic", ""))
    teaching_traits = (state.get("teaching_traits") or "").strip()
    external_knowledge = (state.get("external_knowledge") or "").strip()
    repair_feedback = _targeted_feedback(state, "subject")
    sources = state.get("knowledge_sources") or []
    source_line = f"外部资料来源：{', '.join(sources)}\n" if sources else ""
    knowledge_block = (
        f"\n\n请优先参考以下教师上传资料（若与通用知识冲突，以教师资料为准）：\n{external_knowledge}"
        if external_knowledge
        else ""
    )
    scope_block = (
        f"\n范围硬约束：本次内容必须严格围绕“{scope}”，禁止跨回目泛化到其他章节。\n"
        if scope
        else ""
    )
    traits_block = f"\n教师教学风格要求：\n{teaching_traits}\n" if teaching_traits else ""

    if repair_feedback:
        prompt = ChatPromptTemplate.from_messages(
            [
                (
                    "system",
                    "你是教研 Agent，正在处理质检返工。请重新输出完整学科大纲，"
                    "必须修正主题范围、知识边界和资料依据。不要输出修改说明。",
                ),
                (
                    "user",
                    "教学主题：{topic}\n"
                    "章节范围约束：{scope}\n"
                    "教师教学风格要求：\n{teaching_traits}\n"
                    "教师上传资料摘要：\n{external_knowledge}\n\n"
                    "上一版大纲：\n{outline}\n\n"
                    "质检反馈：\n{feedback}\n\n"
                    "请重新输出结构化大纲，至少包含：学习目标、核心知识点、常见误区、课堂互动建议。",
                ),
            ]
        )
        response = (prompt | llm).invoke(
            {
                "topic": state["topic"],
                "scope": scope or "无（按主题常规覆盖）",
                "teaching_traits": teaching_traits,
                "external_knowledge": external_knowledge,
                "outline": state.get("outline", ""),
                "feedback": repair_feedback,
            }
        )
        return {
            "messages": [response],
            "outline": _strip_internal_tool_markup(response.content),
            "iteration_count": state.get("iteration_count", 0),
        }

    if not messages:
        messages = [
            (
                "system",
                "你是一位资深学科教研专家。你最多只能调用 1 次 search_web 搜索工具。"
                "搜索后必须立即输出纯文字大纲，禁止反复搜索。",
            ),
            (
                "user",
                "请围绕教学主题给出结构化大纲，至少包含："
                "学习目标、核心知识点、常见误区、课堂互动建议。\n"
                f"教学主题：{state['topic']}\n"
                f"{source_line}"
                f"{scope_block}"
                f"{traits_block}"
                f"{knowledge_block}"
            ),
        ]

    search_call_count = _count_tool_calls(messages, "search_web")
    current_llm = llm_with_tools if search_call_count < 1 else llm
    response = current_llm.invoke(messages)

    outline = state.get("outline", "")
    if not getattr(response, "tool_calls", None):
        outline = _strip_internal_tool_markup(response.content)

    return {
        "messages": [response],
        "outline": outline,
        "iteration_count": state.get("iteration_count", 0),
    }


def pedagogy_planner_node(state: LessonState):
    scope = _extract_topic_scope(state.get("topic", ""))
    teaching_traits = (state.get("teaching_traits") or "").strip()
    external_knowledge = (state.get("external_knowledge") or "").strip()
    repair_feedback = _targeted_feedback(state, "pedagogy")
    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                "你是教学设计 Agent。请给出可执行课堂设计，输出 markdown，必须包含：\n"
                "## 学情假设\n"
                "## 教学目标对齐（知识/能力/思政或素养）\n"
                "## 45-90分钟课堂流程（分阶段、含时间）\n"
                "## 分层教学策略（基础/进阶）\n"
                "## 常见误区与纠偏提示\n"
                "如果收到质检返工反馈，必须重新输出完整教学设计，不要只补片段。",
            ),
            (
                "user",
                "教学主题：{topic}\n"
                "章节范围约束：{scope}\n"
                "已有大纲：\n{outline}\n"
                "教师教学风格要求：\n{teaching_traits}\n"
                "教师上传资料摘要：\n{external_knowledge}"
                "{repair_block}",
            ),
        ]
    )
    response = (prompt | llm).invoke(
        {
            "topic": state["topic"],
            "scope": scope or "无（按主题常规覆盖）",
            "outline": state.get("outline", ""),
            "teaching_traits": teaching_traits,
            "external_knowledge": external_knowledge,
            "repair_block": (
                "\n\n上一版教学设计：\n{plan}\n\n质检反馈：\n{feedback}\n".format(
                    plan=state.get("pedagogy_plan", ""),
                    feedback=repair_feedback,
                )
                if repair_feedback
                else ""
            ),
        }
    )

    return {
        "pedagogy_plan": _strip_internal_tool_markup(response.content),
        "messages": [response],
    }


def activity_designer_node(state: LessonState):
    scope = _extract_topic_scope(state.get("topic", ""))
    teaching_traits = (state.get("teaching_traits") or "").strip()
    external_knowledge = (state.get("external_knowledge") or "").strip()
    repair_feedback = _targeted_feedback(state, "activity")
    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                "你是课堂互动 Agent。请输出 markdown，必须包含：\n"
                "## 课堂互动活动\n"
                "- 活动1（目标、步骤、时间、分组方式、产出物）\n"
                "- 活动2（目标、步骤、时间、分组方式、产出物）\n"
                "## 形成性评价\n"
                "- 至少3条可观察评价指标\n"
                "如果收到质检返工反馈，必须重新输出完整互动活动方案，不要只补片段。",
            ),
            (
                "user",
                "教学主题：{topic}\n"
                "章节范围约束：{scope}\n"
                "教学设计草案：\n{pedagogy_plan}\n"
                "课程大纲：\n{outline}\n"
                "教师教学风格要求：\n{teaching_traits}\n"
                "教师上传资料摘要：\n{external_knowledge}"
                "{repair_block}",
            ),
        ]
    )
    response = (prompt | llm).invoke(
        {
            "topic": state["topic"],
            "scope": scope or "无（按主题常规覆盖）",
            "pedagogy_plan": state.get("pedagogy_plan", ""),
            "outline": state.get("outline", ""),
            "teaching_traits": teaching_traits,
            "external_knowledge": external_knowledge,
            "repair_block": (
                "\n\n上一版互动活动方案：\n{plan}\n\n质检反馈：\n{feedback}\n".format(
                    plan=state.get("activity_plan", ""),
                    feedback=repair_feedback,
                )
                if repair_feedback
                else ""
            ),
        }
    )

    return {
        "activity_plan": _strip_internal_tool_markup(response.content),
        "messages": [response],
    }


def quiz_designer_node(state: LessonState):
    include_quiz = state.get("include_quiz", True)
    if not include_quiz:
        return {
            "quiz_bank": "",
            "messages": [],
        }

    scope = _extract_topic_scope(state.get("topic", ""))
    teaching_traits = (state.get("teaching_traits") or "").strip()
    external_knowledge = (state.get("external_knowledge") or "").strip()
    repair_feedback = _targeted_feedback(state, "quiz")
    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                "你是题库 Agent。输出 markdown，必须包含 4-6 道单选题。"
                "每题必须有题干、A/B/C/D、正确答案、一句解析。"
                "题目难度要覆盖基础与进阶。"
                "格式必须稳定：每道题用“1.”编号，A/B/C/D 四个选项各占一行，"
                "答案解析必须与题目一一对应。"
                "必须使用以下两级标题：## 七、课后选择题（至少4题） 和 ## 八、答案与解析。"
                "答案区必须写成“1. B：解析文字”，不要只写字母。"
                "如果收到质检返工反馈，必须重新输出完整题库，不要只补答案或只补解析。",
            ),
            (
                "user",
                "教学主题：{topic}\n"
                "章节范围约束：{scope}\n"
                "课程大纲：\n{outline}\n"
                "教学设计：\n{pedagogy_plan}\n"
                "教师教学风格要求：\n{teaching_traits}\n"
                "教师上传资料摘要：\n{external_knowledge}"
                "{repair_block}",
            ),
        ]
    )
    response = (prompt | llm).invoke(
        {
            "topic": state["topic"],
            "scope": scope or "无（按主题常规覆盖）",
            "outline": state.get("outline", ""),
            "pedagogy_plan": state.get("pedagogy_plan", ""),
            "teaching_traits": teaching_traits,
            "external_knowledge": external_knowledge,
            "repair_block": (
                "\n\n上一版题库草案：\n{quiz}\n\n质检反馈：\n{feedback}\n".format(
                    quiz=state.get("quiz_bank", ""),
                    feedback=repair_feedback,
                )
                if repair_feedback
                else ""
            ),
        }
    )

    return {
        "quiz_bank": _strip_internal_tool_markup(response.content),
        "messages": [response],
    }


def content_creator_node(state: LessonState):
    include_code = state.get("include_code", True)
    include_quiz = state.get("include_quiz", True)
    scope = _extract_topic_scope(state.get("topic", ""))
    teaching_traits = (state.get("teaching_traits") or "").strip()
    external_knowledge = (state.get("external_knowledge") or "").strip()
    sources = state.get("knowledge_sources") or []
    source_line = ", ".join(sources) if sources else "无"
    cleaned_feedback = _sanitize_feedback_for_regen(state.get("feedback", ""))
    repair_target = state.get("repair_target", "")
    if cleaned_feedback and repair_target and repair_target != "content":
        feedback_text = (
            "\n上一版讲义被质检打回，相关子 Agent 已按反馈重新产出对应模块。"
            "请基于最新的大纲、教学设计、互动活动和题库草案重新整合完整讲义，"
            "不要只输出修改片段，也不要沿用旧的错误章节：\n" + cleaned_feedback
        )
    elif cleaned_feedback:
        feedback_text = (
            "\n你上次提交的草稿被打回。请逐条修复以下问题，并重新输出一份完整讲义，不要只输出修改片段：\n" + cleaned_feedback
        )
    else:
        feedback_text = ""

    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                "你是讲义总编 Agent。把多个子 Agent 的产出整合成可投屏讲义。\n"
                "输出必须是 markdown。\n"
                "必选主章节：\n"
                "# {topic} 讲义\n"
                "## 一、学习目标\n"
                "## 二、先修知识与学情假设\n"
                "## 三、课堂流程（含时间分配）\n"
                "## 四、核心知识点\n"
                "## 五、互动活动设计\n"
                "{optional_sections}\n"
                "规则：\n"
                "1. 第一到第五部分必须写得更详细，每部分至少 6 行有效内容。\n"
                "2. 第一到第五部分采用“简短段落 + 要点列表”混合写法：\n"
                "   - 先写 1-2 段解释（每段建议 60-120 字）\n"
                "   - 再写 3-6 条要点（可含子标题）\n"
                "3. 第三部分（课堂流程）必须有清晰时间分配与教学活动说明，避免只写标题。\n"
                "4. 第四部分（核心知识点）至少覆盖概念、原理、误区、课堂案例四类信息。\n"
                "5. 第五部分（互动活动设计）每个活动必须包含：目标、步骤、教师动作、学生产出、评价方式。\n"
                "{code_quiz_rules}\n"
                "6. 若包含题库章节，必须优先复制[题库草案]中的完整题干、A/B/C/D选项、答案与解析；"
                "不要压缩成一行，不要删掉编号，不要只保留答案。\n"
                "7. 严禁输出 HTML、链接占位符和内部工具标记。\n"
                "8. 如果收到质检反馈，必须从标题开始重新输出完整讲义，不能只输出某一章节、答案解析或修改说明。"
                "{scope_rule}"
                "{feedback_text}",
            ),
            (
                "user",
                "主题：{topic}\n\n"
                "[章节范围约束]\n{scope}\n\n"
                "[学科大纲]\n{outline}\n\n"
                "[教学设计]\n{pedagogy_plan}\n\n"
                "[互动活动]\n{activity_plan}\n\n"
                "[题库草案]\n{quiz_bank}\n\n"
                "[教师教学风格要求]\n{teaching_traits}\n\n"
                "[教师外部资料来源]\n{source_line}\n\n"
                "[教师外部资料摘要]\n{external_knowledge}",
            ),
        ]
    )

    response = (prompt | llm).invoke(
        {
            "topic": state["topic"],
            "scope": scope or "无（按主题常规覆盖）",
            "outline": state.get("outline", ""),
            "pedagogy_plan": state.get("pedagogy_plan", ""),
            "activity_plan": state.get("activity_plan", ""),
            "quiz_bank": state.get("quiz_bank", ""),
            "teaching_traits": teaching_traits,
            "source_line": source_line,
            "external_knowledge": external_knowledge,
            "optional_sections": (
                ("## 六、示例代码\n" if include_code else "")
                + ("## 七、课后选择题（至少4题）\n## 八、答案与解析\n" if include_quiz else "")
            ),
            "code_quiz_rules": (
                ("- 若包含示例代码章节，代码块必须可运行且语言为 python。\n" if include_code else "- 本次配置关闭示例代码：禁止输出任何 ```python 代码块或“示例代码”章节。\n")
                + ("- 若包含课后题章节，需至少 4 道单选题并给出答案解析。\n" if include_quiz else "- 本次配置关闭课后题：禁止输出选择题与答案章节。\n")
            ),
            "scope_rule": (
                f"\n9. 范围硬约束：本次内容必须严格聚焦“{scope}”，若偏向其他回目（例如把最后一回写成第一回）则视为不合格。\n"
                if scope
                else ""
            ),
            "feedback_text": feedback_text,
        }
    )

    content = _strip_internal_tool_markup(response.content)
    if include_quiz:
        content = _normalize_quiz_sections_for_validation(content)
    return {
        "draft_content": content,
        "iteration_count": state.get("iteration_count", 0) + 1,
        "messages": [response],
    }


def evaluation_expert_node(state: LessonState):
    include_code = state.get("include_code", True)
    include_quiz = state.get("include_quiz", True)
    scope = _extract_topic_scope(state.get("topic", ""))
    hard_scope_feedback = _hard_scope_violation(scope, state.get("draft_content", ""))
    if hard_scope_feedback:
        target, category = _classify_repair_target(hard_scope_feedback, include_quiz=include_quiz)
        return {
            "is_approved": False,
            "feedback": hard_scope_feedback,
            "repair_target": target,
            "feedback_category": category,
            "messages": [],
        }
    hard_quality_feedback = _draft_quality_feedback(
        state.get("draft_content", ""),
        include_code=include_code,
        include_quiz=include_quiz,
    )
    if hard_quality_feedback:
        target, category = _classify_repair_target(hard_quality_feedback, include_quiz=include_quiz)
        return {
            "is_approved": False,
            "feedback": hard_quality_feedback,
            "repair_target": target,
            "feedback_category": category,
            "messages": [],
        }

    return {
        "is_approved": True,
        "final_content": state.get("draft_content", ""),
        "feedback": "",
        "repair_target": "",
        "feedback_category": "",
        "messages": [],
    }
