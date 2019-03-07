# coding=utf-8

from io import StringIO
import logging
import pandas as pd
import pptx_template.text as txt
import pptx_template.pyel as pyel

log = logging.getLogger()


def load_data_into_table(table, model):
    # テーブルIDを取得
    table_id = get_table_id(table)
    if not table_id:
        return
    log.debug("table_id:%s" % table_id)

    # 対象テーブルIDのモデル情報を取得
    table_setting = pyel.eval_el(table_id, model)
    log.debug(u" Found table_id: %s, table_setting: %s" % (table_id, table_setting))

    # データを流し込む
    _load_data(table, table_setting)

# cell(0, 0)からテーブルIDを取得
def get_table_id(table):
    text_0_0 = table.cell(0, 0).text_frame.text
    return txt.search_first_el(text_0_0)

# テーブルにデータを流し込む
def _load_data(table, table_setting):
    # テーブル流し込みはtsv_bodyのみ対応
    tsv_body = table_setting.get('tsv_body')
    if not tsv_body:
        return

    df = pd.read_csv(StringIO(tsv_body), delimiter='\t', index_col=False, header=None)

    # データ流し込み
    for (irow, row) in enumerate(table.rows):
        for (icell, cell) in enumerate(row.cells):
            for (iparagraph, paragraph) in enumerate(cell.text_frame.paragraphs):
                if iparagraph == 0:
                    for (irun, run) in enumerate(paragraph.runs):
                        if irun == 0:
                            val = df.iloc[irow, icell]
                            run.text = "" if pd.isnull(val) else str(val)
                        else:
                            run.text = ""
                else:
                    paragraph.clear()