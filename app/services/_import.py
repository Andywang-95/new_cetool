from copy import copy
from functools import partial
from zipfile import BadZipFile

import pandas as pd
from tabulate import tabulate

import app.services.utils as utils


class ImportService:
    def __init__(self, data, bom_path, log_func):
        self.database_path = data["database_path"]
        self.bom_path = bom_path
        self.log = partial(log_func, "import")
        _, self.bom_name, _ = utils.path_detail(bom_path)
        (
            (self.wb_mapping, self.ws_mapping),
            self.wb_maintain,
            self.mapping_path,
            self.maintain_path,
        ) = utils.read_db_files(self.database_path)
        self.bom_df, _, self.mapping_df = utils.read_files(
            self.bom_path, self.database_path
        )

    def run(self, method):
        """不同模式的共用流程"""
        msg = utils.check_bom(self.bom_path)
        if utils.check_and_log(msg, self.log):
            return

        # 檢查資料庫檔案狀態
        msg = utils.check_database(self.database_path)
        if utils.check_and_log(msg, self.log):
            return
        self.log(f"<p>Starting import from 【{self.bom_name}】...</p>")
        try:
            self._process(method)
        except BadZipFile:
            self.log(
                "<p>Import failed: <br/>&emsp;請確認 mapping.xlsx 和 maintain.xlsx 是否被加密或損毀</p><hr/>"
            )
            return
        except KeyError as e:
            self.log(
                f"<p>Import failed: <br/>&emsp;檔案規格不符：<br/>&emsp;&emsp;缺少必要欄位或辨識值 -> 【{e}】</p><hr/>"
            )
            return

    def _bom_detail(self, method):
        """不同 import 實作差異邏輯"""
        missing_pn = pd.Series()
        result_data = pd.DataFrame()
        if method == "main":
            bom_data = self.bom_df.iloc[6:].copy()
            bom_data.columns = self.bom_df.iloc[5]
            bom_data = bom_data[bom_data["Number"].str.len() == 16]
            bom_data["main_comment"] = bom_data["CE Comment"].where(
                bom_data["Action"] == "Add"
            )
            bom_data["main_comment"] = bom_data["main_comment"].ffill()
            bom_data.loc[bom_data["CE Comment"] == "同主料", "CE Comment"] = bom_data[
                "main_comment"
            ]
            missing_pn = bom_data.loc[
                bom_data["CE Comment"].isna() | (bom_data["CE Comment"] == ""), "Number"
            ]
            bom_data = bom_data[~bom_data["Number"].isin(missing_pn)]
            result_data = bom_data.loc[
                ~bom_data["Number"].isin(self.mapping_df["料號"]),
                ["Number", "Description", "Spec", "CE Comment"],
            ]
        elif method == "system":
            bom_data = self.bom_df.iloc[1:].copy()
            bom_data.columns = self.bom_df.iloc[0]
            bom_data = bom_data[bom_data["元件/替代料號"].str.len() == 16]
            bom_data["main_comment"] = bom_data["CE Comment"].where(
                bom_data["主件料號"].notna() & (bom_data["主件料號"] != "")
            )
            bom_data["main_comment"] = bom_data["main_comment"].ffill()
            bom_data.loc[bom_data["CE Comment"] == "同主料", "CE Comment"] = bom_data[
                "main_comment"
            ]
            missing_pn = bom_data.loc[
                bom_data["CE Comment"].isna() | (bom_data["CE Comment"] == ""),
                "元件/替代料號",
            ]
            bom_data = bom_data[~bom_data["元件/替代料號"].isin(missing_pn)]
            result_data = bom_data.loc[
                ~bom_data["元件/替代料號"].isin(self.mapping_df["料號"]),
                ["元件/替代料號", "品名規格", "規格", "CE Comment"],
            ]
        return missing_pn, result_data

    def _process(self, method):
        missing_pn, result_data = self._bom_detail(method)
        if not missing_pn.empty:
            self.log(f"<p>以下 {len(missing_pn)} 筆料號 CE Comment 缺失：</p>")
            for number in missing_pn:
                self.log(f"&emsp;{number}<br/>")
            self.log("<hr/>")

        wb_bom, ws_bom = utils.load(self.bom_path)
        if not result_data.empty:
            self.log(f"<p>新增 {len(result_data)} 筆新料號，開始匯入：</p>")
            for idx, r in result_data.iterrows():
                """mapping 料號開始匯入"""
                print(tuple(r))
                mapping_last_row = self.ws_mapping.max_row + 1
                self.ws_mapping.append(tuple(r))
                cell_mapping = self.ws_mapping.cell(mapping_last_row, 4)
                cell_bom = ws_bom.cell(idx + 1, ws_bom.max_column)  # type: ignore
                cell_mapping.fill = copy(cell_bom.fill)  # type: ignore
                """maintain 料號開始匯入"""
                utils.to_maintain(self.wb_maintain, r, cell_bom)
                # 將單列資料轉成 list，再用 tabulate 生成無 header 的 HTML 表格
                html_table = tabulate([r.tolist()], headers=[], tablefmt="html")
                self.log(html_table)
            self.ws_mapping.protection.enable()
            self.wb_mapping.save(self.mapping_path)
            self.wb_maintain.save(self.maintain_path)
            utils.clean_and_save_new_versions(self.wb_mapping, self.wb_maintain)
            self.log("<p>匯入完成</p><hr/>")
        else:
            self.log("<p>無新料號可匯入</p><hr/>")
