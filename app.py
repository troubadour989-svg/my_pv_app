# -*- coding: utf-8 -*-
import datetime
import io
import os
import re
import pandas as pd
import streamlit as st
from extractor import NationalSolarExtractor

st.set_page_config(
    page_title="全国光伏备案数据清洗平台", page_icon="⚔️", layout="wide"
)

st.title("⚔️ 全国光伏备案清洗神器")
st.caption(
    "核心能力：精准省份隔离 | 联合公司匹配 | 户用/工商业判定 | 全流程防卡顿状态保持"
)

# ==========================================
# 0. 初始化 Session State 状态（防止交互时重跑）
# ==========================================
if "cleaned_df" not in st.session_state:
    st.session_state["cleaned_df"] = None
if "excel_bytes" not in st.session_state:
    st.session_state["excel_bytes"] = None
if "display_prov" not in st.session_state:
    st.session_state["display_prov"] = "未知"
if "prov_detail_str" not in st.session_state:
    st.session_state["prov_detail_str"] = ""
if "file_prov_suffix" not in st.session_state:
    st.session_state["file_prov_suffix"] = "全国"


# 动态生成标准配置模板 Excel
@st.cache_data
def get_template_excel():
    template_data = [
        {
            "大区": "华南大区",
            "适用省份": "湖南省",
            "集团品牌": "正泰",
            "识别关键词": "湖南湘泰",
        },
        {
            "大区": "西部大区",
            "适用省份": "重庆市",
            "集团品牌": "TCL",
            "识别关键词": "重庆毅敏",
        },

    ]
    df_template = pd.DataFrame(template_data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False, sheet_name="配置规则")
    return output.getvalue()


# ==========================================
# 1. 侧边栏：映射规则管理 & 省份强制设置
# ==========================================
with st.sidebar:
    st.header("⚙️ 规则与省份配置")

    # 1. 模板下载
    st.download_button(
        label="📥 下载标准配置模板 Excel",
        data=get_template_excel(),
        file_name="集团品牌映射表_标准模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("---")

# 2. 上传自定义规则
    custom_rule_file = st.file_uploader(
        "📤 上传新增/临时规则表 (自动追加合并)", 
        type=["xlsx"],
        help="支持仅包含新增几行的小表格，系统将自动与底库智能合并，绝不丢失原有规则！"
    )
    default_mapping_file = "brand_mapping.xlsx"
    required_cols = {"大区", "适用省份", "集团品牌", "识别关键词"}

    # 第一步：先读取系统自带的基准底库
    if os.path.exists(default_mapping_file):
        try:
            base_df = pd.read_excel(default_mapping_file)
            base_df.columns = base_df.columns.astype(str).str.strip()
        except Exception:
            base_df = pd.DataFrame(columns=list(required_cols))
    else:
        base_df = pd.DataFrame(columns=list(required_cols))

    # 第二步：如果有同事上传新表，执行【智能增量合并】
    if custom_rule_file:
        try:
            uploaded_df = pd.read_excel(custom_rule_file)
            uploaded_df.columns = uploaded_df.columns.astype(str).str.strip()

            if not required_cols.issubset(set(uploaded_df.columns)):
                missing = required_cols - set(uploaded_df.columns)
                st.error(f"❌ 上传失败！表格缺少必要列：{missing}")
                brand_mapping_df = base_df
            else:
                # 核心机制：同事上传的排在前面，与底库合并后，按【适用省份 + 识别关键词】去重
                # keep='first' 保证：若有冲突，以同事最新上传的为准！
                combined_df = pd.concat([uploaded_df, base_df], ignore_index=True)
                
                # 剔除空关键词
                combined_df = combined_df.dropna(subset=["识别关键词"])
                combined_df["_clean_kw"] = combined_df["识别关键词"].astype(str).str.strip()
                combined_df = combined_df[combined_df["_clean_kw"] != ""]
                
                # 智能去重（适用省份 + 关键词）
                combined_df = combined_df.drop_duplicates(
                    subset=["适用省份", "_clean_kw"], 
                    keep="first"
                ).drop(columns=["_clean_kw"])

                brand_mapping_df = combined_df
                
                # 给同事极其温暖、确定的正反馈
                st.success(
                    f"✅ **智能增量合并成功！**\n\n"
                    f"- 系统基准底库：`{len(base_df)}` 条\n"
                    f"- 您上传的规则：`{len(uploaded_df)}` 条\n"
                    f"- 智能合并去重后，当前全库共生效：**`{len(brand_mapping_df)}`** 条！"
                )
        except Exception as e:
            st.error(f"读取规则文件失败: {e}")
            brand_mapping_df = base_df
    else:
        # 没有上传时，直接使用基准底库
        brand_mapping_df = base_df

    st.caption(f"📊 当前内存生效规则总数：**{len(brand_mapping_df)}** 条")

    st.markdown("---")
    st.header("📍 运行模式")
    st.info(
        "💡 **已启用分省自适应引擎**：系统将严格依据原始数据中的【备案省份】列进行逐行分省物理隔离清洗，支持多省混合文件直接上传！"
    )

# ==========================================
# 2. 主页面：数据上传与处理
# ==========================================
uploaded_files = st.file_uploader(
    "📥 上传原始备案 Excel 文件 (支持多选批量合并)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"📁 已就绪 **{len(uploaded_files)}** 个文件。")

    if st.button("🚀 启动全国标准化清洗", type="primary"):
        all_results = []
        has_error = False

        progress_bar = st.progress(0)
        status = st.status("⏳ 数据清洗管线启动中...", expanded=True)

        for i, file in enumerate(uploaded_files):
            status.write(f"正在读取并清洗: `{file.name}` ...")

            # 1. 读入原始文件
            try:
                raw_df = pd.read_excel(file, converters={"备案项目名称": str})
                raw_df.columns = raw_df.columns.astype(str).str.strip()
                raw_df = raw_df.dropna(subset=["备案项目名称"], how="all")
                raw_df = raw_df[raw_df["备案项目名称"].astype(str).str.strip() != ""]
            except Exception as e:
                status.update(label="❌ 读取文件失败", state="error")
                st.error(f"文件 `{file.name}` 读取错误: {e}")
                has_error = True
                break

            # 2. 运行清洗核心引擎（底层按行上的【备案省份】进行强隔离清洗）
            try:
                extractor = NationalSolarExtractor(
                    df=raw_df,
                    brand_mapping_df=brand_mapping_df,
                    district_file_path="全国区划编码.csv",
                )
                processed_df = extractor.run_pipeline()
                all_results.append(processed_df)
            except ValueError as ve:
                status.update(label="❌ 数据校验未通过", state="error")
                st.error(f"文件 `{file.name}` 校验错误: {ve}")
                has_error = True
                break
            except Exception as ex:
                status.update(label="❌ 清洗过程异常", state="error")
                st.error(f"文件 `{file.name}` 处理异常: {ex}")
                has_error = True
                break

            progress_bar.progress(int((i + 1) / len(uploaded_files) * 60))

        if not has_error and all_results:
            status.write("正在执行全量数据合并与去重...")
            final_df = pd.concat(all_results, ignore_index=True)

            if "备案项目代码" in final_df.columns:
                final_df = final_df.drop_duplicates(
                    subset=["备案项目代码"], keep="first"
                )

            progress_bar.progress(75)

            # ==========================================
            # 3. 动态提取清洗覆盖的真实省份与分布统计
            # ==========================================
            def clean_p(p):
                if pd.isna(p):
                    return ""
                return (
                    str(p)
                    .strip()
                    .replace("省", "")
                    .replace("市", "")
                    .replace("自治区", "")
                )

            raw_provs = final_df["备案省份"].dropna().astype(str).tolist()
            unique_provs = sorted(list({clean_p(p) for p in raw_provs if clean_p(p)}))

            # 统计各省的具体行数分布
            prov_series = final_df["备案省份"].apply(clean_p)
            prov_counts = prov_series[prov_series != ""].value_counts().to_dict()
            prov_detail_str = " | ".join(
                [f"**{p}**: {count:,}行" for p, count in prov_counts.items()]
            )

            # 根据识别出的省份数量自适应排版
            if len(unique_provs) == 1:
                display_prov = f"【{unique_provs[0]}】"
                file_prov_suffix = unique_provs[0]
            elif 1 < len(unique_provs) <= 3:
                display_prov = f"【{'、'.join(unique_provs)}】(共{len(unique_provs)}省)"
                file_prov_suffix = "_".join(unique_provs)
            elif len(unique_provs) > 3:
                display_prov = (
                    f"【{'、'.join(unique_provs[:3])} 等多省混合】(共{len(unique_provs)}省)"
                )
                file_prov_suffix = f"多省混合({len(unique_provs)}省)"
            else:
                display_prov = "【全国通用】"
                file_prov_suffix = "全国"

            # 4. 核心呈现列排版规范
            status.write("正在对核心指标进行顺序排版...")
            final_df.columns = final_df.columns.astype(str).str.strip()

            target_cols = [
                "备案省份",  # 强显备案省份列
                "备案项目代码",
                "备案项目名称",
                "备案项目公司",
                "项目分类",
                "类型判断关键字",
                "所属集团",
                "集团判断关键字",
                "省",
                "地市",
                "区县",
                "备案办理时间",
                "建设规模",
            ]
            final_cols = [c for c in target_cols if c in final_df.columns]
            other_cols = [c for c in final_df.columns if c not in final_cols]

            ordered_final_df = final_df[final_cols + other_cols]

            # 5. 生成 Excel 打包
            status.write("📦 正在压制打包 Excel 文件...")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                ordered_final_df.to_excel(
                    writer, index=False, sheet_name="清洗精炼结果"
                )

            excel_data = output.getvalue()
            progress_bar.progress(100)

            status.update(
                label="🎉 清洗与打包全部完成！", state="complete", expanded=False
            )

            # 写入 Session 状态持久保持
            st.session_state["cleaned_df"] = ordered_final_df
            st.session_state["excel_bytes"] = excel_data
            st.session_state["display_prov"] = display_prov
            st.session_state["prov_detail_str"] = prov_detail_str
            st.session_state["file_prov_suffix"] = file_prov_suffix

# ==========================================
# 4. 结果展示与下载区域（基于 Session 渲染，防闪烁）
# ==========================================
if st.session_state["cleaned_df"] is not None:
    df_show = st.session_state["cleaned_df"]

    st.markdown("---")

    # 【无歧义动态指示牌】：同时展示覆盖省份 + 各省明细行数
    st.info(
        f"🏷️ **本次清洗覆盖省份**：{st.session_state['display_prov']}  |  "
        f"**总行数**：**{len(df_show):,}** 行  "
        f"（各省分布 👉 {st.session_state['prov_detail_str']}）"
    )

    # 核心看板指标
    c1, c2, c3 = st.columns(3)
    c1.metric("总备案项目数", f"{len(df_show):,} 条")
    c2.metric("户用项目总数", f"{(df_show['项目分类'] == '户用').sum():,} 条")
    c3.metric(
        "工商业项目总数", f"{(df_show['项目分类'] == '工商业').sum():,} 条"
    )

    # 导出下载按钮（文件名动态包含所有涉及省份）
    tz_beijing = datetime.timezone(datetime.timedelta(hours=8))
    time_tag = datetime.datetime.now(tz_beijing).strftime("%Y%m%d_%H%M")
    
    export_filename = (
        f"全国备案清洗_{st.session_state['file_prov_suffix']}_{time_tag}.xlsx"
    )

    st.download_button(
        label=f"📥 点击下载最终精炼 Excel 结果表 ({export_filename})",
        data=st.session_state["excel_bytes"],
        file_name=export_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    # 集团排名统计
    st.write("##### 🏆 匹配集团分布 Top 10")
    brand_counts = df_show["所属集团"].value_counts().head(10)
    st.dataframe(brand_counts, use_container_width=True)

    # 网页数据预览（排版与 Excel 100% 相同）
    st.write("##### 🔍 最终数据前 20 行预览 (与导出 Excel 顺序 100% 一致)")
    st.dataframe(df_show.head(20), use_container_width=True)