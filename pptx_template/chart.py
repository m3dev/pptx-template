#
# coding=utf-8

import logging
from io import StringIO

from pptx.shapes.graphfrm import GraphicFrame
from pptx.chart.data import ChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE as ct

import pandas as pd

import pptx_template.pyel as pyel
import pptx_template.text as txt
import pptx_template.pptx_util as util

log = logging.getLogger()

def _build_xy_chart_data(csv):
  chart_data = XyChartData()
  for i in range(1, csv.columns.size):
    series = chart_data.add_series(csv.columns[i])
    xy_col = csv.iloc[:, [0, i]]
    for (_, row) in xy_col.iterrows():
      log.debug(u"Adding xy %d,%d" % (row[1], row[0]))
      series.add_data_point(row[1], row[0])
  return chart_data

def _build_chart_data(csv):
  chart_data = ChartData()
  log.debug(u"Setting categories with values:%s" % (csv.iloc[:,0].values.tolist()))
  chart_data.categories = csv.iloc[:,0].values.tolist()

  for i in range(1, csv.columns.size):
    col = csv.iloc[:, i]
    log.debug(u"Adding series:%s" % (col.name))
    chart_data.add_series(col.name, col.values.tolist())
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
  csv_body = chart_setting.get('body')
  if csv_body:
    csv_file_name = StringIO(csv_body)
    log.info(u"Loading from csv string: %s" % csv_body)
  else:
    csv_file_name = chart_setting.get('file_name')
    if not csv_file_name:
      csv_file_name = "%s.csv" % chart_id
    log.info(u"Loading from csv file: %s" % csv_file_name)

  return pd.read_csv(csv_file_name)

def _replace_chart_data_with_csv(chart, chart_id, chart_setting):
  """
    1つのチャートに対して指定されたCSVからデータを読み込む。
  """
  csv = _load_csv_into_dataframe(chart_id, chart_setting)

  if _is_xy_chart(chart):
    log.info(u"setting csv into XY chart %s" % chart_id)
    chart_data = _build_xy_chart_data(csv)
  else:
    log.info(u"setting csv int chart %s" % chart_id)
    chart_data = _build_chart_data(csv)

  chart.replace_data(chart_data)

  log.info(u"chart data replacement completed.")

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
    log.info(u"found chart_id: %s. setting: %s" % (chart_id, chart_setting))

    txt.replace_el_in_text_frame_with_str(title_frame, chart_id, '')
    _replace_chart_data_with_csv(chart, chart_id, chart_setting)
    _set_value_axis(chart, chart_id, chart_setting)

def select_all_chart_shapes(slide):
  return [ s.chart for s in slide.shapes if isinstance(s, GraphicFrame) and s.shape_type == 3 ]
