# -*- coding: utf-8 -*-
"""
核心清洗引擎（全国通用标准版·分省强隔离+弹性匹配增强）
三大核心功能：1.行政区划匹配 2.集团品牌匹配(Excel驱动) 3.项目分类(户用/工商业)
增强特性：
- 严格校验【备案省份】列（缺失或全空直接报错截断）
- 按省份物理隔离清洗，支持多省混合数据安全并发
- 集团品牌关键词弹性正则编译（全半角括号、随机空格自动容错兼容）
"""

import io
import os
import re
import warnings
import numpy as np
import pandas as pd
from config_data import PROJECT_CATEGORIES

warnings.filterwarnings("ignore")


class NationalSolarExtractor:

    def __init__(
        self,
        df,
        brand_mapping_df,
        district_file_path="全国区划编码.csv",
        province=None,  # 兼容参数，实际清洗以数据内的【备案省份】为准
    ):
        """
        :param df: 上传的原始数据 DataFrame
        :param brand_mapping_df: 从 Excel 读入的集团映射表 DataFrame
        :param district_file_path: 区划编码路径
        :param province: 兼容保留字段
        """
        self.raw_df = df.copy()
        # 清除输入数据表头的首尾隐藏空格
        self.raw_df.columns = self.raw_df.columns.astype(str).str.strip()

        self.brand_mapping_df = brand_mapping_df.copy()
        self.brand_mapping_df.columns = (
            self.brand_mapping_df.columns.astype(str).str.strip()
        )

        self.district_file = district_file_path
        self.fallback_province = province
        self.df = None

        # 核心前置校验：检查【备案省份】列
        self._validate_and_prepare_provinces()

    # -------------------------------------------------------------
    # 核心前置校验：【备案省份】强校验与标准化
    # -------------------------------------------------------------
    def _validate_and_prepare_provinces(self):
        """严格校验【备案省份】列是否存在且有效"""
        target_col = "备案省份"

        # 1. 检查列是否存在
        if target_col not in self.raw_df.columns:
            similar_cols = [c for c in self.raw_df.columns if "省" in c]
            err_msg = f"【致命错误】原始数据中缺少必要的【{target_col}】列！"
            if similar_cols:
                err_msg += f" 检测到相似列: {similar_cols}，请核对表头标准名称。"
            raise ValueError(err_msg)

        # 2. 检查整列是否全为空值
        valid_series = self.raw_df[target_col].dropna().astype(str).str.strip()
        valid_series = valid_series[valid_series != ""]
        if valid_series.empty:
            raise ValueError(
                f"【致命错误】数据中的【{target_col}】列全为空值，无法进行分省清洗！"
            )

        # 3. 标准化省份名称（剔除“省”、“市”、“自治区”、空格）
        def clean_prov(p):
            if pd.isna(p):
                return ""
            return (
                str(p)
                .strip()
                .replace("省", "")
                .replace("市", "")
                .replace("自治区", "")
            )

        self.raw_df["_内部清洗省份"] = self.raw_df[target_col].apply(clean_prov)

        # 提示个别缺失行
        missing_count = (self.raw_df["_内部清洗省份"] == "").sum()
        if missing_count > 0:
            warnings.warn(
                f"【警告】检测到共有 {missing_count} 行数据的【{target_col}】为空，这些行将无法匹配特定省份规则！",
                UserWarning,
            )

    @staticmethod
    def _clean_prov_str(p):
        if pd.isna(p):
            return ""
        return (
            str(p)
            .strip()
            .replace("省", "")
            .replace("市", "")
            .replace("自治区", "")
        )

# -------------------------------------------------------------
    # 弹性正则编译（解决括号全半角、空格变形导致 Mapping 臃肿的问题）
    # -------------------------------------------------------------
    def _compile_group_patterns_for_prov(self, current_province):
        """针对特定省份，编译专属品牌正则（含全国/global）"""
        patterns = {}
        if self.brand_mapping_df.empty:
            return patterns

        cur_p = self._clean_prov_str(current_province)

        def is_match_prov(x):
            if pd.isna(x):
                return False
            val = self._clean_prov_str(x).lower()
            return (val == cur_p) or (val in ["global", "全国"])

        mask_prov = self.brand_mapping_df["适用省份"].apply(is_match_prov)
        df_target = self.brand_mapping_df[mask_prov]

        # 弹性正则生成器：使括号（全/半角）、空格具备自适应容错能力
        def make_elastic_pattern(kw):
            escaped = re.escape(kw)
            # 采用原生字符串替换，安全可靠，杜绝 bad escape 错误
            escaped = escaped.replace(r"\(", r"\s*[\(（]?\s*")
            escaped = escaped.replace(r"\（", r"\s*[\(（]?\s*")
            escaped = escaped.replace(r"\)", r"\s*[\)）]?\s*")
            escaped = escaped.replace(r"\）", r"\s*[\)）]?\s*")
            escaped = escaped.replace(r"\ ", r"\s*")
            return escaped

        for brand, group in df_target.groupby("集团品牌"):
            kws = (
                group["识别关键词"]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )
            kws = [k for k in kws if len(k) > 0]
            if not kws:
                continue

            # 按关键词长度倒序排序（长词优先匹配）
            sorted_kws = sorted(kws, key=len, reverse=True)
            elastic_patterns = [make_elastic_pattern(k) for k in sorted_kws]

            patterns[brand] = re.compile(
                f"({'|'.join(elastic_patterns)})",
                flags=re.IGNORECASE,
            )

        return patterns

    def _compile_category_patterns_for_prov(self, current_province):
        """编译特定省份的项目类型分类正则"""
        cat_patterns = {}
        cur_p = self._clean_prov_str(current_province)

        for cat, scope_dict in PROJECT_CATEGORIES.items():
            kws = scope_dict.get("global", []).copy()
            if cur_p in scope_dict:
                kws.extend(scope_dict[cur_p])

            if not kws:
                cat_patterns[cat] = re.compile(r"(匹配不到_#$)")
            else:
                sorted_kw = sorted(list(set(kws)), key=len, reverse=True)
                cat_patterns[cat] = re.compile(
                    f"({'|'.join(map(re.escape, sorted_kw))})"
                )
        return cat_patterns

    # -------------------------------------------------------------
    # 单省隔离处理器（三大核心功能）
    # -------------------------------------------------------------
    def _process_single_province(self, df_prov, cur_prov):
        df = df_prov.copy()
        clean_p = self._clean_prov_str(cur_prov)

        # 编译仅属于该省的规则字典
        group_patterns = self._compile_group_patterns_for_prov(clean_p)
        category_patterns = self._compile_category_patterns_for_prov(clean_p)

        # ---------------------------------------------------------
        # 1. 行政区划匹配 (仅湖南通过区划编码匹配，其余保留空白)
        # ---------------------------------------------------------
        ADMIN_MATCH_PROVINCES = {"湖南"}
        if clean_p not in ADMIN_MATCH_PROVINCES:
            for col in ["省", "地市", "区县"]:
                if col not in df.columns:
                    df[col] = clean_p if col == "省" else ""
                else:
                    df[col] = df[col].fillna("")
        else:
            if "备案项目代码" in df.columns and os.path.exists(self.district_file):
                df["区划编码"] = df["备案项目代码"].str.extract(
                    r"\d{4}-(\d{5,6})-\d{2}-\d{2}-\d{6}"
                )
                dist_df = pd.read_csv(self.district_file, dtype={"区划编码": str})
                dist_df.columns = dist_df.columns.str.strip()
                dist_df["区划编码"] = dist_df["区划编码"].str.strip()
                target_cols = [
                    c for c in ["区划编码", "省", "地市", "区县"] if c in dist_df.columns
                ]
                dist_df = dist_df[target_cols].drop_duplicates("区划编码")
                df = df.merge(dist_df, on="区划编码", how="left")

            for col in ["省", "地市", "区县"]:
                if col not in df.columns:
                    df[col] = "未知"
                else:
                    df[col] = df[col].fillna("未知")

        # ---------------------------------------------------------
        # 2. 集团品牌匹配 (基于弹性正则与归一化文本)
        # ---------------------------------------------------------
        df["所属集团"] = ""
        df["集团判断关键字"] = ""

        col_comp = (
            df["备案项目公司"].fillna("").astype(str)
            if "备案项目公司" in df.columns
            else ""
        )
        col_name = (
            df["备案项目名称"].fillna("").astype(str)
            if "备案项目名称" in df.columns
            else ""
        )

        raw_search_text = col_comp + " " + col_name
        # 归一化文本：将制表符及多个连续空格压缩为单个空格
        search_text = raw_search_text.str.replace(r"\s+", " ", regex=True)
        is_non_person = search_text.str.contains("非自然人", na=False)

        # 2.1 匹配品牌规则
        rem_mask = df["所属集团"] == ""
        for brand, pattern in group_patterns.items():
            if not rem_mask.any():
                break
            matches = search_text[rem_mask].str.extract(pattern)[0]
            hits = matches.dropna().index
            if not hits.empty:
                df.loc[hits, "所属集团"] = brand
                df.loc[hits, "集团判断关键字"] = matches[hits]
                rem_mask.loc[hits] = False

        # 2.2 自然人兜底（纯姓名格式）
        name_pat = r"^([一-龥]{2,3}|[一-龥]{1,2}\*)$"
        is_name = search_text[rem_mask].str.match(name_pat, na=False)
        name_idx = is_name[is_name].index
        if not name_idx.empty:
            df.loc[name_idx, "所属集团"] = "自然人"
            df.loc[name_idx, "集团判断关键字"] = "姓名格式"
            rem_mask.loc[name_idx] = False

        # 2.3 包含“自然人”字样
        is_person_str = search_text.str.contains("自然人", na=False)
        mask_person = rem_mask & is_person_str & (~is_non_person)
        df.loc[mask_person, "所属集团"] = "自然人"
        df.loc[mask_person, "集团判断关键字"] = "自然人"

        # 2.4 非自然人兜底
        mask_unresolved = rem_mask & is_non_person
        df.loc[mask_unresolved, "所属集团"] = "其他非自然人"
        df.loc[mask_unresolved, "集团判断关键字"] = "非自然人"

        # ---------------------------------------------------------
        # 3. 项目分类 (户用 / 工商业 / 集中式 / 公建)
        # ---------------------------------------------------------
        text = df["备案项目名称"].str.lower().fillna("")

        # 公建白名单控制
        PUBLIC_MODE_PROVINCES = {
            "重庆", "四川", "陕西", "贵州", "云南", "江苏", "浙江", "广东"
        }
        is_public_allowed = clean_p in PUBLIC_MODE_PROVINCES

        kw_res = text.str.extract(r"(户用|住宅|民居|自然人|自然村)")[0]
        kw_ind = text.str.extract(r"(工商业|厂房|工厂|商铺)")[0]

        found_cat = pd.Series(index=df.index, dtype=object)
        found_kw = pd.Series(index=df.index, dtype=object)

        public_cats = ["公建(户用模式)", "公建(需判断)", "陕西光伏+"]
        industrial_cats = [c for c in PROJECT_CATEGORIES.keys() if "工商业" in c]
        priority_cats = public_cats + ["集中式"] + industrial_cats + ["户用补强"]

        for cat in priority_cats:
            if cat not in category_patterns:
                continue
            pat = category_patterns[cat]
            mask = found_cat.isna()
            ext = text[mask].str.extract(pat)[0]
            hits = ext.dropna().index
            found_cat.loc[hits] = cat
            found_kw.loc[hits] = ext[hits]

        # 非白名单省份，公建强制降级为工商业
        disp_cat = found_cat.copy()
        if not is_public_allowed:
            mask_public = disp_cat.isin(["公建(户用模式)", "公建(需判断)"])
            disp_cat[mask_public] = "工商业"

        disp_type = disp_cat.copy()
        disp_type[found_cat == "户用补强"] = "户用"
        disp_type[found_cat.str.contains("工商业", na=False)] = "工商业"

        # 品牌驱动判定
        group_col = df["所属集团"].fillna("")
        is_effective_brand = (group_col != "") & (
            ~group_col.isin(["自然人", "其他非自然人", ""])
        )
        cond_allow_public = is_public_allowed & found_cat.isin(public_cats)

        conds = [
            cond_allow_public,  # 1. 白名单省份公建
            kw_ind.notna(),  # 2. 强工商业
            kw_res.notna(),  # 3. 强户用
            is_effective_brand,  # 4. 品牌商备案 -> 偏户用
            disp_type.notna(),  # 5. 命中分类词库
            text.str.contains("公司"),  # 6. 带公司后缀 -> 工商业
            text.str.contains(r"村|组|户", regex=True),  # 7. 村组后缀 -> 户用
        ]

        choices_type = [
            found_cat,
            "工商业",
            "户用",
            "户用",
            disp_type,
            "工商业",
            "户用",
        ]

        choices_kw = [
            found_kw,
            kw_ind,
            kw_res,
            "品牌集团",
            found_kw,
            "AA公司主体",
            "村组后缀",
        ]

        df["项目分类"] = np.select(conds, choices_type, default="待核查")
        df["类型判断关键字"] = np.select(conds, choices_kw, default="无匹配")

        # 自然人修正
        fix_person = (df["所属集团"] == "自然人") & (
            df["项目分类"].isin(["工商业", "待核查"])
        )
        df.loc[fix_person, "项目分类"] = "户用"
        df.loc[fix_person, "类型判断关键字"] = (
            df.loc[fix_person, "类型判断关键字"] + "(自然人修正)"
        )

        return df

    # -------------------------------------------------------------
    # 统一运行管线（按【备案省份】分省隔离执行，无损还原行序）
    # -------------------------------------------------------------
    def run_pipeline(self):
        processed_chunks = []

        # 按清洗后的省份分组独立处理
        for prov_key, group_df in self.raw_df.groupby(
            "_内部清洗省份", sort=False
        ):
            chunk_res = self._process_single_province(group_df, prov_key)
            processed_chunks.append(chunk_res)

        # 合并所有省份数据，并严格按原始输入的索引行序还原
        final_df = pd.concat(processed_chunks, axis=0).loc[self.raw_df.index]

        # 移除内部辅助列
        if "_内部清洗省份" in final_df.columns:
            final_df.drop(columns=["_内部清洗省份"], inplace=True)

        self.df = final_df
        return self.df

    # -------------------------------------------------------------
    # 内存导出 Excel
    # -------------------------------------------------------------
    @staticmethod
    def export_excel(df_to_save):
        output = io.BytesIO()
        target_cols = [
            "备案省份",
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
        final_cols = [c for c in target_cols if c in df_to_save.columns]
        other_cols = [c for c in df_to_save.columns if c not in final_cols]

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_to_save[final_cols + other_cols].to_excel(
                writer, index=False, sheet_name="清洗结果"
            )
        return output.getvalue()