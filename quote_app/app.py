from __future__ import annotations

import pandas as pd
import streamlit as st

from app_config import (
    DATA_DIR,
    EXPORT_PATH,
    IGNORED_TERMS_PATH,
    OUTPUT_DIR,
    PRICE_DB_PATH,
    RULES_PATH,
    SECTION_CONFIRM_COLUMNS,
    FINAL_QUOTE_COLUMNS,
)
from app_services import (
    append_default_ignored_terms,
    apply_activity_section_changes,
    apply_selected_suggestions,
    clear_result_state,
    configure_app_services,
    final_quote_export_df,
    get_file_signature,
    get_text_hash,
    initialize_session_state,
    make_final_quote_df,
    make_judgement_df,
    make_sub_activity_section_df,
    make_suggestion_editor_df,
    merge_final_quote_edits,
    merge_judgement_edits,
    normalize_feedback_df,
    recalculate_final_quote_df,
    reset_editor_widget_state,
    run_recognition,
    sync_working_quote_state,
)
from core.excel_exporter import export_quote_to_excel
from core.normalizer import QUOTE_TYPES, load_rules
from core.quote_builder import ensure_price_db, sort_by_quote_section
from core.rule_feedback import apply_feedback_rows, ensure_ignored_terms, load_ignored_terms
from core.text_reader import read_text_from_upload


st.set_page_config(page_title="活动方案报价单生成器", page_icon="📋", layout="wide")

DATA_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
ensure_price_db(PRICE_DB_PATH)
ensure_ignored_terms(IGNORED_TERMS_PATH)
append_default_ignored_terms()
rules = load_rules(RULES_PATH)
configure_app_services(rules)
initialize_session_state()

st.title("活动方案报价单生成器")
st.caption("输入方案 → 识别报价项 → 人工判断 → 编辑报价单 → 导出 Excel")
if st.session_state.get("feedback_rerun_notice"):
    st.success(st.session_state.feedback_rerun_notice)
    st.session_state.feedback_rerun_notice = ""

st.header("1. 方案输入区")
col_a, col_b, col_c = st.columns(3)
with col_a:
    activity_name = st.text_input("活动名称", placeholder="例如：农野开心乐")
with col_b:
    activity_time = st.text_input("活动时间", placeholder="例如：2026年6月")
with col_c:
    client_name = st.text_input("单位名称", placeholder="例如：某某单位")

uploaded_file = st.file_uploader("上传方案文件（支持 .txt / .docx / .pdf）", type=["txt", "docx", "pdf"])
if uploaded_file is not None:
    new_signature = get_file_signature(uploaded_file)
    if st.session_state.get("current_file_signature") != new_signature:
        try:
            clear_result_state()
            extracted_text = read_text_from_upload(uploaded_file)
            st.session_state["plan_text_input"] = extracted_text
            st.session_state["plan_text_value"] = extracted_text
            st.session_state["current_file_signature"] = new_signature
            st.session_state["last_text_hash"] = get_text_hash(extracted_text)
            st.session_state["upload_notice"] = "已读取上传文件，并填入方案文本区域。"
            st.rerun()
        except Exception as exc:
            st.error(f"文件读取失败：{exc}")

if st.session_state.get("upload_notice"):
    st.success(st.session_state["upload_notice"])
    st.session_state["upload_notice"] = ""

plan_text = st.text_area("方案文本", key="plan_text_input", height=220, placeholder="粘贴活动方案文本，或上传文件自动读取。")
st.session_state["plan_text_value"] = plan_text
col_recognize, col_rerun = st.columns([1, 1])
with col_recognize:
    recognize_clicked = st.button("识别报价项", type="primary")
with col_rerun:
    rerun_clicked = st.button("重新识别当前方案")

if recognize_clicked or rerun_clicked:
    plan_text = st.session_state.get("plan_text_value", "")
    if not plan_text.strip():
        st.warning("请先输入或上传方案文本。")
    else:
        current_text_hash = get_text_hash(plan_text)
        if current_text_hash != st.session_state.get("last_text_hash", ""):
            clear_result_state()
            st.session_state["plan_text_value"] = plan_text
            st.session_state.last_text_hash = current_text_hash
        run_recognition(plan_text)
        st.success("已完成识别。" if recognize_clicked else "已使用最新规则重新识别当前方案。")

if st.session_state.coverage is not None:
    st.header("2. 识别结果摘要")
    working_quote_df = st.session_state.get("working_quote_df", st.session_state.quote_df)
    judgement_df = make_judgement_df(working_quote_df)
    confirm_count = 0 if judgement_df.empty else int((judgement_df["确认状态"] != "已确认").sum())
    candidate_count = len(st.session_state.candidate_display_df)

    metric_a, metric_b, metric_c, metric_d = st.columns(4)
    metric_a.metric("已识别报价项", f"{len(working_quote_df)} 项")
    metric_b.metric("需人工确认", f"{confirm_count} 项")
    metric_c.metric("建议补充项", f"{len(st.session_state.suggested_df)} 项")
    metric_d.metric("未识别候选", f"{candidate_count} 项")

    if st.session_state.activity_type_df.empty:
        st.write("识别到的活动类型：暂未命中明确模板")
    else:
        activity_types = " / ".join(st.session_state.activity_type_df["活动类型"].astype(str).tolist())
        st.write(f"识别到的活动类型：{activity_types}")

    diagnostics = st.session_state.get("activity_content_diagnostics", {})
    with st.expander("活动内容区域调试信息", expanded=False):
        ranges = diagnostics.get("ranges", []) if isinstance(diagnostics, dict) else []
        if not ranges:
            st.warning("未定位到活动内容区域，已启用全文强标题、白名单、编号标题、连续短标题和报价项反推兜底。")
        else:
            debug_rows = [
                {
                    "区域": item.get("index", index + 1),
                    "起点": item.get("start", 0),
                    "终点": item.get("end", 0),
                    "截取文本长度": item.get("length", 0),
                    "疑似误用了目录页": "是" if item.get("directory_like") else "否",
                    "提示": "；".join(item.get("warnings", [])),
                }
                for index, item in enumerate(ranges)
            ]
            st.dataframe(pd.DataFrame(debug_rows), width="stretch", hide_index=True)
            if bool(diagnostics.get("has_short_range")):
                st.warning("活动内容区域可能切片错误")
            if any(bool(item.get("directory_like")) for item in ranges):
                st.warning("活动内容区域疑似误用了目录页")
    filtered_candidate_df = st.session_state.get("section_candidate_diagnostics", pd.DataFrame())
    if isinstance(filtered_candidate_df, pd.DataFrame) and not filtered_candidate_df.empty:
        with st.expander("被过滤的活动板块候选", expanded=False):
            st.dataframe(filtered_candidate_df, width="stretch", hide_index=True)

    section_df = st.session_state.activity_section_df.copy()
    sub_section_df = make_sub_activity_section_df(st.session_state.get("activity_sections_all", []))
    if not section_df.empty and "层级" not in section_df.columns:
        section_df["层级"] = section_df["置信度"].map(lambda value: "主活动板块" if str(value) == "strong" else "疑似子活动")
    strong_sections = []
    section_candidate_count = 0
    if not section_df.empty:
        strong_sections = section_df[
            (section_df["层级"] == "主活动板块") & (section_df["是否作为报价板块"].astype(bool))
        ]["标准板块名"].astype(str).tolist()
        section_candidate_count = len(sub_section_df)

    if strong_sections:
        lines = [f"{index}. {section}" for index, section in enumerate(strong_sections, start=1)]
        st.write("识别到的活动板块：")
        for line in lines:
            st.write(line)
        if section_candidate_count:
            st.caption(f"疑似活动板块：{section_candidate_count} 个，点击展开确认")
    else:
        st.info("未识别到明确活动板块，已根据项目内容反推报价板块，请人工检查“未归属板块”。")

    if not section_df.empty:
        with st.expander("活动板块确认区", expanded=False):
            if "section_editor_df" not in st.session_state:
                st.session_state["section_editor_df"] = section_df.copy()
            editor_source_df = st.session_state.get("section_editor_df", section_df.copy())
            if isinstance(editor_source_df, pd.DataFrame) and not editor_source_df.empty:
                editor_source_df = editor_source_df.copy()
                for column in SECTION_CONFIRM_COLUMNS:
                    if column not in editor_source_df.columns:
                        editor_source_df[column] = ""
            else:
                editor_source_df = section_df.copy()

            main_preview = editor_source_df[editor_source_df["层级"] == "主活动板块"]["标准板块名"].astype(str).tolist()
            if main_preview:
                st.write("主活动板块：")
                for section_name in main_preview:
                    st.write(f"- {section_name}")
            if not sub_section_df.empty:
                with st.expander("疑似子活动（默认不作为报价板块）", expanded=False):
                    st.dataframe(sub_section_df, width="stretch", hide_index=True)

            with st.form("activity_section_form"):
                edited_section_df = st.data_editor(
                    editor_source_df[SECTION_CONFIRM_COLUMNS],
                    width="stretch",
                    hide_index=True,
                    num_rows="fixed",
                    disabled=["层级", "原始识别文本", "置信度", "判断原因"],
                    column_config={
                        "是否作为报价板块": st.column_config.CheckboxColumn("是否作为报价板块", default=False),
                    },
                    key="activity_section_editor",
                )
                apply_sections = st.form_submit_button("应用活动板块修改")

            if apply_sections:
                st.session_state["section_editor_df"] = edited_section_df.copy()
                base_quote_df = st.session_state.get("working_quote_df", st.session_state.get("raw_quote_df", st.session_state.quote_df)).copy()
                updated_quote_df, updated_section_df, updated_sections, confirmed_sections = apply_activity_section_changes(
                    base_quote_df,
                    edited_section_df,
                    st.session_state.get("activity_sections_all", []),
                )
                st.session_state.activity_section_df = updated_section_df.copy()
                st.session_state.section_editor_df = updated_section_df.copy()
                st.session_state.activity_sections_all = updated_sections
                st.session_state.activity_sections = confirmed_sections
                st.session_state.confirmed_sections = [section.get("name", "") for section in confirmed_sections]
                st.session_state.suggested_df = sort_by_quote_section(st.session_state.suggested_df, confirmed_sections)
                sync_working_quote_state(updated_quote_df, confirmed_sections)
                st.session_state["quote_editor_version"] = st.session_state.get("quote_editor_version", 0) + 1
                st.session_state.excel_bytes = None
                reset_editor_widget_state()
                st.success("已应用活动板块修改，最终报价单已更新。")
                st.rerun()

    if st.session_state.coverage < 0.6:
        st.warning("当前方案存在较多疑似报价内容未进入报价单，请优先检查：1. 未识别候选清单 2. 建议补充项 3. 需确认报价项")

    st.header("3. 需要你判断的项目")
    working_quote_df = st.session_state.get("working_quote_df", st.session_state.quote_df)
    if working_quote_df.empty:
        st.info("暂未识别到报价项，可以在规则配置中补充同义词后重试。")
    else:
        st.subheader("报价项保留与确认")
        review_editor_df = st.session_state.get("review_quote_df", judgement_df.copy())
        with st.form("review_quote_form"):
            edited_judgement_df = st.data_editor(
                review_editor_df,
                width="stretch",
                hide_index=True,
                num_rows="fixed",
                disabled=["来源依据"],
                column_config={
                    "是否加入": st.column_config.CheckboxColumn("是否加入", default=True),
                    "数量": st.column_config.NumberColumn("数量", min_value=0.0, step=1.0),
                    "确认状态": st.column_config.SelectboxColumn("确认状态", options=["已确认", "需确认数量", "需确认规格", "需确认是否报价", "系统建议"]),
                },
                key="review_quote_editor",
            )
            review_submitted = st.form_submit_button("应用本表修改")
        if review_submitted:
            st.session_state["review_quote_df"] = edited_judgement_df.copy()
            updated_working_df = merge_judgement_edits(working_quote_df, edited_judgement_df)
            st.session_state["working_quote_df"] = updated_working_df.copy()
            st.session_state["quote_df"] = updated_working_df.copy()
            st.session_state["recognized_df"] = updated_working_df.copy()
            updated_final_df = make_final_quote_df(updated_working_df)
            st.session_state["final_quote_df"] = updated_final_df.copy()
            st.session_state["confirmed_quote_df"] = updated_final_df.copy()
            st.session_state["excel_bytes"] = None

    st.subheader("建议补充项")
    if st.session_state.suggested_df.empty:
        st.caption("暂无建议补充项。")
    else:
        suggestion_editor_source = st.session_state.get("suggestion_editor_df")
        if not isinstance(suggestion_editor_source, pd.DataFrame) or suggestion_editor_source.empty:
            suggestion_editor_source = make_suggestion_editor_df(st.session_state.suggested_df)
        with st.form("suggestion_quote_form"):
            edited_suggestion_df = st.data_editor(
                suggestion_editor_source,
                width="stretch",
                hide_index=True,
                num_rows="fixed",
                column_config={
                    "是否加入报价单": st.column_config.CheckboxColumn("是否加入报价单", default=False),
                    "数量": st.column_config.NumberColumn("数量", min_value=0.0, step=1.0),
                    "报价类型": st.column_config.SelectboxColumn("报价类型", options=list(QUOTE_TYPES)),
                },
                key="suggestion_quote_editor",
            )
            suggestion_submitted = st.form_submit_button("应用建议补充项")
        if suggestion_submitted:
            st.session_state.suggestion_editor_df = edited_suggestion_df.copy()
            added_count, duplicate_items = apply_selected_suggestions(edited_suggestion_df)
            if duplicate_items:
                st.warning(f"已存在 {len(duplicate_items)} 项，未重复添加：{'、'.join(dict.fromkeys(duplicate_items))}")
            if added_count:
                st.success(f"已加入 {added_count} 个建议补充项。")
                st.rerun()
            elif not duplicate_items:
                st.info("请先勾选需要加入报价单的建议补充项。")

    st.header("4. 最终报价单编辑区")
    working_quote_df = st.session_state.get("working_quote_df", st.session_state.quote_df)
    final_quote_df = st.session_state.get("final_quote_df")
    if final_quote_df is None or final_quote_df.empty and not working_quote_df.empty:
        final_quote_df = make_final_quote_df(working_quote_df)
        st.session_state.final_quote_df = final_quote_df.copy()
        st.session_state.confirmed_quote_df = final_quote_df.copy()
    elif final_quote_df is not None and not final_quote_df.empty:
        final_quote_df = recalculate_final_quote_df(final_quote_df)
        st.session_state.final_quote_df = final_quote_df.copy()
        st.session_state.confirmed_quote_df = final_quote_df.copy()
    if final_quote_df.empty:
        st.info("当前没有勾选进入报价单的项目。")
    else:
        with st.form("final_quote_form"):
            edited_final_df = st.data_editor(
                final_quote_df,
                width="stretch",
                hide_index=True,
                num_rows="fixed",
                disabled=["合计"],
                column_config={
                    "数量": st.column_config.NumberColumn("数量", min_value=0.0, step=1.0),
                    "单价": st.column_config.NumberColumn("单价", min_value=0.0, step=1.0, format="%.2f"),
                    "合计": st.column_config.NumberColumn("合计", disabled=True, format="%.2f"),
                },
                key="final_quote_editor",
            )
            final_submitted = st.form_submit_button("保存报价单编辑")
        if final_submitted:
            edited_final_df = recalculate_final_quote_df(edited_final_df)
            st.session_state["final_quote_df"] = edited_final_df.copy()
            updated_working_df = merge_final_quote_edits(working_quote_df, edited_final_df)
            st.session_state["working_quote_df"] = updated_working_df.copy()
            st.session_state["quote_df"] = updated_working_df.copy()
            st.session_state["recognized_df"] = updated_working_df.copy()
            st.session_state["confirmed_quote_df"] = edited_final_df.copy()
            st.session_state["excel_bytes"] = None
        st.caption("这里就是最终导出的报价单明细。单价为 0 时，合计保持 0。")

    st.header("5. 导出区")
    final_export_source_df = st.session_state.get("final_quote_df", pd.DataFrame(columns=FINAL_QUOTE_COLUMNS))
    can_export = not final_export_source_df.empty
    if st.button("生成 Excel 报价单", disabled=not can_export):
        final_export_source_df = recalculate_final_quote_df(st.session_state.get("final_quote_df", pd.DataFrame(columns=FINAL_QUOTE_COLUMNS)))
        st.session_state["final_quote_df"] = final_export_source_df.copy()
        export_df = final_quote_export_df(final_export_source_df)
        excel_bytes = export_quote_to_excel(
            export_df,
            activity_name=activity_name,
            activity_time=activity_time,
            client_name=client_name,
            output_path=EXPORT_PATH,
        )
        st.session_state.excel_bytes = excel_bytes
        st.success("Excel 报价单已生成。")

    if st.session_state.get("excel_bytes"):
        st.download_button(
            "下载报价单.xlsx",
            data=st.session_state.excel_bytes,
            file_name="报价单.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.header("6. 未识别候选处理区")
    with st.expander(f"未识别候选：发现 {candidate_count} 个可能漏项，点击展开处理", expanded=False):
        if st.session_state.candidate_display_df.empty:
            st.caption("暂无需要处理的未识别候选。")
        else:
            processing_options = ["暂不处理", "加入已有标准项目别名", "新建标准项目", "标记为无需报价"]
            standard_item_options = ["", *sorted(rules.keys())]
            editor_source_df = normalize_feedback_df(st.session_state.feedback_df)
            with st.form("candidate_feedback_form"):
                feedback_df = st.data_editor(
                    editor_source_df,
                    width="stretch",
                    hide_index=True,
                    num_rows="fixed",
                    disabled=["候选词", "可能归类", "上下文摘要"],
                    column_config={
                        "是否处理": st.column_config.CheckboxColumn("是否处理", default=False),
                        "处理方式": st.column_config.SelectboxColumn("处理方式", options=processing_options),
                        "选择标准项目": st.column_config.SelectboxColumn("选择标准项目", options=standard_item_options),
                    },
                    key="candidate_feedback_editor",
                )
                with st.expander("新建标准项目字段（仅选择新建标准项目时使用）", expanded=False):
                    new_col_a, new_col_b, new_col_c = st.columns(3)
                    with new_col_a:
                        new_category = st.text_input("项目分类", value="待分类")
                        new_standard_item = st.text_input("标准项目", placeholder="默认使用候选词")
                    with new_col_b:
                        new_quote_type = st.selectbox("报价类型", options=list(QUOTE_TYPES), index=2)
                        new_default_unit = st.text_input("默认单位", value="项")
                    with new_col_c:
                        new_default_desc = st.text_area("默认说明", value="待补充说明", height=100)
                feedback_submitted = st.form_submit_button("保存规则更新")

            if feedback_submitted:
                feedback_df = normalize_feedback_df(feedback_df)
                st.session_state.feedback_df = feedback_df
                result = apply_feedback_rows(
                    RULES_PATH,
                    IGNORED_TERMS_PATH,
                    feedback_df.to_dict("records"),
                    {
                        "项目分类": new_category,
                        "标准项目": new_standard_item,
                        "报价类型": new_quote_type,
                        "默认单位": new_default_unit,
                        "默认说明": new_default_desc,
                    },
                )
                messages = result.get("messages", [])
                if messages:
                    with st.expander("保存明细", expanded=False):
                        for message in messages:
                            st.write(message)
                    st.success(f"已更新 {result.get('rules_updated', 0)} 条规则，已忽略 {result.get('ignored_count', 0)} 条候选词")
                    if result.get("rules_updated", 0) or result.get("ignored_count", 0):
                        st.success("规则已更新，下次识别将自动生效。")
                else:
                    st.info("没有需要保存的规则更新。")

            if st.button("重新识别当前方案", key="rerun_after_feedback"):
                if plan_text.strip():
                    run_recognition(plan_text)
                    st.session_state.feedback_rerun_notice = "已按最新规则重新识别当前方案。"
                    st.rerun()
                else:
                    st.warning("当前没有可重新识别的方案文本。")

    st.header("7. 高级信息折叠区")
    with st.expander("高级信息：活动类型识别详情", expanded=False):
        if st.session_state.activity_type_df.empty:
            st.caption("暂无活动类型详情。")
        else:
            st.dataframe(st.session_state.activity_type_df, width="stretch", hide_index=True)

    with st.expander("高级信息：完整识别数据", expanded=False):
        st.dataframe(st.session_state.recognized_df, width="stretch", hide_index=True)

    with st.expander("高级信息：完整未识别候选", expanded=False):
        if st.session_state.candidate_df.empty:
            st.caption("暂无完整候选数据。")
        else:
            st.dataframe(st.session_state.candidate_df, width="stretch", hide_index=True)

    with st.expander("高级信息：规则库调试信息", expanded=False):
        st.write(f"规则库路径：{RULES_PATH}")
        st.write(f"忽略词路径：{IGNORED_TERMS_PATH}")
        st.write(f"标准项目数量：{len(rules)}")
        st.write(f"忽略词数量：{len(load_ignored_terms(IGNORED_TERMS_PATH))}")
else:
    st.info("请先输入或上传方案文本，然后点击“识别报价项”。")
