# csv2pptx2 - import csv to powerpoint template
# coding=utf-8

from pptx import Presentation
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.table import Table
from pptx.chart.data import ChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE as ct

from io import StringIO
import sys
import codecs
import re
import logging
import pandas as pd

import pptx_template.pyel as pyel
import pptx_template.pptx_util as pptx_util

log = logging.getLogger()

id_regex = re.compile(r"\{([A-Za-z0-9._\-]+)\}")

def search_id(src):
   text_id_match = id_regex.search(src)
   if text_id_match:
     return text_id_match.group(1)
   return None


def select_all_text_shapes(slide):
  return [ s for s in slide.shapes if s.shape_type in [1,14,17] ]


def select_all_chart_shapes(slide):
  return [ s.chart for s in slide.shapes if isinstance(s, GraphicFrame) and s.shape_type == 3 ]


def select_all_tables(slide):
  return [ s.table for s in slide.shapes if isinstance(s, GraphicFrame) and s.shape_type == 19 ]


def replace_el_in_table(table, model):
  """
   table の各セルの中に EL 形式があれば、それを model の該当する値と置き換える
  """
  for cell in [ cell for row in table.rows for cell in row.cells ]:
    replace_el_in_shape_text(cell.text_frame, model)


def replace_el_in_shape_text(shape, model):
  """
   shape.text の中に EL 形式が一つ以上あれば、それを model の該当する値と置き換える
  """

  while True:
    text_id = search_id(shape.text)
    if not text_id:
      return
    log.info(u"found text_id: %s. replacing: %s" % (text_id, shape.text))
    replacing_text = pyel.eval_el(text_id, model)
    if isinstance(replacing_text, str):
      shape.text = shape.text.replace(u"{%s}" % text_id, replacing_text)
    else:
       log.error("Invalid value for id:%s, model:%s" % (text_id, replacing_text))
       return

def _build_xy_chart_data(csv):
  chart_data = XyChartData()
  for i in range(1, csv.columns.size):
    series = chart_data.add_series(csv.columns[i])
    xy_col = csv.ix[:, [0, i]]
    for (_, row) in xy_col.iterrows():
      log.debug(u"adding xy %d,%d" % (row[1], row[0]))
      series.add_data_point(row[1], row[0])
  return chart_data

def _build_chart_data(csv):
  chart_data = ChartData()
  for i in range(1, csv.columns.size):
    col = csv.ix[:, i]
    log.debug(u"adding series %s" % (col.name))
    chart_data.add_series(col.name, col.values.tolist())
  return chart_data

def _is_xy_chart(chart):
  xy_charts = [ct.XY_SCATTER_LINES, ct.XY_SCATTER_LINES_NO_MARKERS, ct.XY_SCATTER, ct.XY_SCATTER_SMOOTH, ct.XY_SCATTER_SMOOTH_NO_MARKERS]
  return chart.chart_type in xy_charts

def set_value_axis(chart, chart_id, chart_setting):
  max = chart_setting.get('value_axis_max')
  min = chart_setting.get('value_axis_min')

  pptx_util.set_value_axis(chart, max = max, min = min)


def load_csv_into_dataframe(chart_id, chart_setting):
  csv_body = chart_setting.get('body')
  if csv_body:
    csv_file_name = StringIO(csv_body)
    log.info(u"loading from csv string: %s" % csv_body)
  else:
    csv_file_name = chart_setting.get('file_name')
    if not csv_file_name:
      csv_file_name = "%s.csv" % chart_id
    log.info(u"loading from csv file: %s" % csv_file_name)

  return pd.read_csv(csv_file_name)

def replace_chart_data_with_csv(chart, chart_id, chart_setting):
  """
    1つのチャートに対して指定されたCSVからデータを読み込む。
  """
  csv = load_csv_into_dataframe(chart_id, chart_setting)

  if _is_xy_chart(chart):
    log.info(u"setting csv into XY chart %s" % chart_id)
    chart_data = _build_xy_chart_data(csv)
  else:
    log.info(u"setting csv int chart %s" % chart_id)
    chart_data = _build_chart_data(csv)

  chart_data.categories = csv.index.values.tolist()
  chart.replace_data(chart_data)

  log.info(u"chart data replacement completed.")

  return


def edit_slide(slide, model):
  """
    1つのスライドに対して文字列置換およびチャートCSV設定を行う
    チャート設定や文字列置換は1スライドに対して複数持てる。配列やdictなどで引渡し、pptxからはEL式で特定する
    チャート設定なのか文字列置換なのかは、EL式の配置されているpptx内のオブジェクトで判断される
    文字列置換やチャートタイトルに EL式で {answer.0} のような形で指定する
    チャート設定として指定可能な項目:
     - file_name : CSVファイルの名前
     - body: CSVファイルの中身そのものを直接文字列で指定できる(file_name, file_encodingは無視される)
     - value_axis_max: チャート左側軸の最大値。(省略可)
     - value_axix_min: チャート左側軸の最小値。(省略可)
     - (TBI) file_encoding: CSVファイルのエンコーディング。省略時は utf-8
  """

  log.info("processing %s" % slide)

  # pptx内の TextFrame の EL表記を model の値で置換する
  for shape in select_all_text_shapes(slide):
    replace_el_in_shape_text(shape, model)
  for shape in select_all_tables(slide):
    replace_el_in_table(shape, model)

  # pptx内の 各チャートに対してcsvの値を設定する
  for chart in select_all_chart_shapes(slide):
    # チャートタイトルから {} で囲まれた文字列を探し、それをキーとしてチャート設定と紐付ける
    if not chart.has_title or not chart.chart_title.has_text_frame:
      continue

    title_frame = chart.chart_title.text_frame
    chart_id = search_id(title_frame.text)
    if not chart_id:
      continue
    chart_setting = pyel.eval_el(chart_id, model)
    log.info(u"found chart_id: %s. setting: %s" % (chart_id, chart_setting))

    title_frame.text = title_frame.text.replace("{%s}" % chart_id, '') # チャートタイトルからチャートID指定部分を削除し他を残す
    replace_chart_data_with_csv(chart, chart_id, chart_setting)
    set_value_axis(chart, chart_id, chart_setting)


def remove_slide(presentation, slide):
  """
   presentation から 指定した slide を削除する
  """
  pptx_util.remove_slide(presentation, slide)


def remove_slide_id(presentation, slide_id):
  """
     指定した id のスライドから {id:foobar} という形式の文字列を削除する
  """
  slide = get_slide(presentation, slide_id)
  for shape in select_all_text_shapes(slide):
    if u"{id:%s}" % slide_id == shape.text:
      shape.text = ''

def remove_all_slides_having_id(presentation):
  """
     {id:foobar} という文字列を持つすべてのスライドを削除する
  """
  id_frame_regex = re.compile(r"\{id:([A-Za-z0-9._\-]+)\}")
  unused_slides = []
  for slide in presentation.slides:
    for shape in select_all_text_shapes(slide):
      if id_frame_regex.search(shape.text):
          unused_slides.append(slide)
          break
  for slide in unused_slides:
    remove_slide(presentation, slide)


def get_slide(presentation, slide_id):
  """
     指定した id に対して {id:foobar} という TextFrame を持つスライドを探す
  """
  for slide in presentation.slides:
    for shape in select_all_text_shapes(slide):
      if u"{id:%s}" % slide_id == shape.text:
        return slide
  raise ValueError(u"slide id:%s not found" % slide_id)
