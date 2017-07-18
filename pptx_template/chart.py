#
# coding=utf-8

import logging
import os.path
from io import StringIO

from pptx.shapes.graphfrm import GraphicFrame
from pptx.chart.data import ChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE as ct

import pandas as pd
import numpy as np

from six import string_types

import pptx_template.pyel as pyel
import pptx_template.text as txt
import pptx_template.pptx_util as util

log = logging.getLogger()

def _nan_to_none(x):
    # log.debug(u" type of x:%s is:%s" % (x, type(x)))
    if isinstance(x, np.generic):
        y = None if np.isnan(x) else x.item()
    elif isinstance(x, string_types):
        y = x if isinstance(x, type(u"a")) else unicode(x,'utf-8')
    else:
        y = x
    return y

def to_unicode(s):
    return s if isinstance(s, type(u"a")) else unicode(s,'utf-8')

def _build_xy_chart_data(csv):
    chart_data = XyChartData()
    for i in range(1, csv.columns.size):
        series = chart_data.add_series(csv.columns[i])
        xy_col = csv.iloc[:, [0, i]]
        for (_, row) in xy_col.iterrows():
            x, y = _nan_to_none(row[0]), _nan_to_none(row[1])
            log.debug(u" Adding xy %s,%s" % (y, x))
            series.add_data_point(y, x)
    return chart_data

def _build_chart_data(csv):
    chart_data = ChartData()
    categories = [_nan_to_none(x) for x in csv.iloc[:,0].values]
    categories = [u"%s" % x if x else u"c%d" % i for i,x in enumerate(categories)]
    log.debug(u" Setting categories with values:%s" % categories)
    chart_data.categories = categories

    for i in range(1, csv.columns.size):
        col = csv.iloc[:, i]
        values = [_nan_to_none(x) for x in col.values]
        name = to_unicode(col.name)
        log.debug(u" Adding series:%s values:%s" % (name, values))
        chart_data.add_series(name, values)
    return chart_data

def _is_xy_chart(chart):
    xy_charts = [ct.XY_SCATTER_LINES, ct.XY_SCATTER_LINES_NO_MARKERS, ct.XY_SCATTER, ct.XY_SCATTER_SMOOTH, ct.XY_SCATTER_SMOOTH_NO_MARKERS]
    return chart.chart_type in xy_charts

def _set_value_axis(chart, chart_id, chart_setting):
    max = chart_setting.get('value_axis_max')
    min = chart_setting.get('value_axis_min')
    if max or min:
        util.set_value_axis(chart, max = max, min = min)

def _load_csv_into_dataframe(chart_id, chart_setting):
    if 'body' in chart_setting:
        csv_body = chart_setting.get('body')
        log.info(u"Loading from CSV string: %s" % csv_body)
        return pd.read_csv(StringIO(csv_body))
    elif 'tsv_body' in chart_setting:
        tsv_body = chart_setting.get('tsv_body')
        log.info(u"Loading from TSV string: %s" % tsv_body)
        return pd.read_csv(StringIO(tsv_body), delimiter='\t')
    else:
        csv_file_name = chart_setting.get('file_name')
        if not csv_file_name:
            for ext in ['csv', 'tsv']:
                csv_file_name = "%s.%s" % (chart_id, ext)
                if os.path.isfile(csv_file_name):
                    break
            else:
                raise ValueError(u"File not found: csv or tsv for %s" % chart_id)

        log.info(u"Loading from csv file: %s" % csv_file_name)
        delimiter = '\t' if csv_file_name.endswith('.tsv') else ','
        return pd.read_csv(csv_file_name, delimiter=delimiter)


def _replace_chart_data_with_csv(chart, chart_id, chart_setting):
    """
        1つのチャートに対して指定されたCSVからデータを読み込む。
    """
    csv = _load_csv_into_dataframe(chart_id, chart_setting)

    if _is_xy_chart(chart):
        log.info(u"Setting csv/tsv into XY chart_id: %s" % chart_id)
        chart_data = _build_xy_chart_data(csv)
    else:
        log.info(u"Setting csv/tsv into chart_id: %s" % chart_id)
        chart_data = _build_chart_data(csv)

    chart.replace_data(chart_data)

    log.info(u"Completed chart data replacement.")

    return


def load_data_into_chart(chart, model):
    # チャートタイトルから {} で囲まれた文字列を探し、それをキーとしてチャート設定と紐付ける
    if not chart.has_title or not chart.chart_title.has_text_frame:
        return

    title_frame = chart.chart_title.text_frame
    chart_id = txt.search_first_el(title_frame.text)
    if not chart_id:
        return

    chart_setting = pyel.eval_el(chart_id, model)
    log.info(u"Found chart_id: %s, chart_setting: %s" % (chart_id, chart_setting))

    txt.replace_el_in_text_frame_with_str(title_frame, chart_id, '')
    _replace_chart_data_with_csv(chart, chart_id, chart_setting)
    _set_value_axis(chart, chart_id, chart_setting)

def select_all_chart_shapes(slide):
    return [ s.chart for s in slide.shapes if isinstance(s, GraphicFrame) and s.shape_type == 3 ]
