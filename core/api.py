# -*- coding: utf-8 -*-
"""系统提示词构建 + API 调用 / 重试 / 错误分类。"""

import time

from core import config


# ============================================================
# 提示词
# ============================================================

def build_system_prompt(target_lang, reader_profile="", want_terms=True,
                        domain="general", extra_prompt=""):
    lang = config.LANG_NAMES.get(target_lang, "简体中文")
    # ★ 修复：按领域选角色，不再一律"文学翻译家"
    role = config._DOMAIN_ROLES.get(domain, config._DOMAIN_ROLES["general"])
    base = (
        f"{role}"
        f"请把用户发来的内容翻译成【{lang}】，并遵循："
        f"1) 译文必须符合{lang}母语者的阅读习惯和审美，流畅、有文采；"
        f"2) 对话要自然生动，符合人物身份；3) 修辞、隐喻、双关尽量找到{lang}对应表达，"
        f"实在无法对应则意译并保留神韵；4) 不遗漏任何内容，不总结，不输出任何解释。\n\n"
        f"【格式要求】用户会给出一页原文的多个段落，每段以 [[B0]] [[B1]] [[B2]] ... 标记开头。"
        f"你必须严格保留所有标记、保持顺序，标记后紧跟该段译文。"
        f"除标记和译文外，不要输出任何其他文字、不加解释、不用代码块。"
    )
    if config.HAS_NOTES and config.NB is not None:
        base += config.NB.build_terms_instruction(reader_profile, want_terms)
    # ★ 新增：领域微调
    if domain in config._DOMAIN_PROMPTS:
        base += config._DOMAIN_PROMPTS[domain]
    # ★ 新增：用户自定义要求
    if extra_prompt and extra_prompt.strip():
        base += "\n【用户额外要求】\n" + extra_prompt.strip() + "\n"
    return base


def build_simple_system(target_lang, domain="general", extra_prompt=""):
    lang = config.LANG_NAMES.get(target_lang, "简体中文")
    # ★ 修复：按领域选角色，不再一律"文学翻译家"
    role = config._DOMAIN_ROLES.get(domain, config._DOMAIN_ROLES["general"])
    base = (
        f"{role}"
        f"请把用户发来的内容翻译成流畅、有文采的【{lang}】，符合{lang}母语者的阅读习惯。"
        f"保留段落结构，不遗漏内容，不总结，不输出解释，直接给出译文。"
    )
    # ★ 新增
    if domain in config._DOMAIN_PROMPTS:
        base += config._DOMAIN_PROMPTS[domain]
    if extra_prompt and extra_prompt.strip():
        base += "\n【用户额外要求】\n" + extra_prompt.strip() + "\n"
    return base


def cache_prefix(target_lang):
    return "" if target_lang == "zh-CN" else f"{target_lang}_"


# ============================================================
# API
# ============================================================

def _extract_status(err):
    for attr in ("status_code", "code"):
        v = getattr(err, attr, None)
        if isinstance(v, int):
            return v
    resp = getattr(err, "response", None)
    if resp is not None:
        v = getattr(resp, "status_code", None)
        if isinstance(v, int):
            return v
    return None


def _is_fatal_api_error(msg):
    if not msg:
        return False
    m = str(msg)
    low = m.lower()
    return ("余额" in m
            or "API Key 无效" in m
            or "无效或已过期" in m
            or "insufficient" in low
            or "unauthorized" in low)


def _interruptible_sleep(seconds, stop_event):
    if stop_event is None:
        time.sleep(seconds)
        return False
    end = time.time() + seconds
    while True:
        remaining = end - time.time()
        if remaining <= 0:
            return False
        if stop_event.is_set():
            return True
        time.sleep(min(0.25, remaining))


def _api_call(client, model, system, user_content, retries=4, stop_event=None):
    last_err = ""
    for attempt in range(retries):
        if stop_event is not None and stop_event.is_set():
            raise RuntimeError("用户已停止")
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user_content},
                ],
                temperature=0.3,
            )
            return resp.choices[0].message.content or ""
        except Exception as e:
            last_err = str(e)
            status = _extract_status(e)
            low = last_err.lower()
            if status == 402 or "insufficient" in low or "余额" in last_err:
                raise RuntimeError("账户余额不足，请去 DeepSeek 平台充值后再继续")
            if status == 401:
                raise RuntimeError("API Key 无效或已过期，请检查后重试")
            # ★ 修复：最后一次重试不再白睡一轮再抛错
            last_attempt = attempt >= retries - 1
            if status == 429:
                if last_attempt:
                    break
                _interruptible_sleep(10 * (attempt + 1), stop_event)
                continue
            if last_attempt:
                break
            _interruptible_sleep(6 * (attempt + 1), stop_event)

    raise RuntimeError(f"API 连续失败：{last_err}")


def call_api(client, model, text, target_lang="zh-CN", retries=4,
             stop_event=None, reader_profile="", want_terms=True,
             domain="general", extra_prompt=""):
    system = build_system_prompt(target_lang, reader_profile, want_terms,
                                  domain=domain, extra_prompt=extra_prompt)
    return _api_call(client, model, system, text, retries, stop_event=stop_event)


def call_simple_api(client, model, text, target_lang="zh-CN", retries=4,
                    stop_event=None, domain="general", extra_prompt=""):
    system = build_simple_system(target_lang, domain, extra_prompt)
    return _api_call(client, model, system, text, retries, stop_event=stop_event)
